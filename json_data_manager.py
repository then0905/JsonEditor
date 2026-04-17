import pandas as pd
import json
import os
import gc
import re
import warnings

warnings.filterwarnings("ignore")


class JsonDataManager:
    def __init__(self, config_path="config.json"):
        self.config_path = config_path
        self._full_config = self._load_config(config_path)
        self.config = {}          # current file config
        self.json_path = None
        self.tables = {}          # {table_name: DataFrame}
        self.sub_tables = {}      # {table_name.key_name: DataFrame}
        self.need_config_alert = False
        self.dirty = False
        self.dirty_cells = set()   # {(table_name, row_idx, col_name)}
        self._recent_files = self._full_config.get("_recent_files", [])
        self._search_index = {}   # {table_name: {token: set of (table_name, row_idx)}}
        self._ref_cache    = {}   # {norm_abs_path: [list of row dicts]}
        self._ref_dict     = {}   # {(norm_abs_path, key_col, val_col): {key_str: val_str}}

    # ──────────────────────────────────────────────
    # Config I/O
    # ──────────────────────────────────────────────

    def _load_config(self, path):
        if not os.path.exists(path):
            return {}
        try:
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def save_config(self):
        self._full_config["_recent_files"] = self._recent_files
        try:
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self._full_config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"儲存 config 失敗: {e}")

    # ──────────────────────────────────────────────
    # Recent files
    # ──────────────────────────────────────────────

    def _add_recent_file(self, path):
        norm = os.path.normpath(path)
        if norm in self._recent_files:
            self._recent_files.remove(norm)
        self._recent_files.insert(0, norm)
        self._recent_files = self._recent_files[:10]

    # ──────────────────────────────────────────────
    # Search index
    # ──────────────────────────────────────────────

    def _build_search_index(self, table_name, df):
        """Build inverted index for fast search: {token → set of (table_name, idx)}"""
        index = {}
        for col in df.columns:
            for idx, val in df[col].items():
                normalized = str(val).lower().strip()
                if not normalized:
                    continue
                # Index the full value and each word token
                for token in set([normalized] + normalized.split()):
                    if token not in index:
                        index[token] = set()
                    index[token].add((table_name, int(idx)))
        self._search_index[table_name] = index

    def search_index(self, query):
        """Index-based search, returns list of (table_name, is_sub, row_idx, matched_cols)."""
        q = query.lower().strip()
        if not q:
            return []
        # Collect candidate (table_name, row_idx) pairs
        candidates = set()
        for table_name, index in self._search_index.items():
            for token, positions in index.items():
                if q in token:
                    candidates.update(positions)

        results = []
        limit = 200
        for table_name, row_idx in candidates:
            if len(results) >= limit:
                break
            is_sub = table_name in self.sub_tables
            df = self.sub_tables.get(table_name) or self.tables.get(table_name)
            if df is None or row_idx not in df.index:
                continue
            # Get matched columns for display
            matched = {col: str(df.at[row_idx, col]) for col in df.columns
                       if q in str(df.at[row_idx, col]).lower()}
            if matched:
                results.append((table_name, is_sub, row_idx, matched))
        return results

    # ──────────────────────────────────────────────
    # Text reference lookup (for text_ref column type)
    # ──────────────────────────────────────────────

    def get_ref_text(self, ref_json_abs: str, ref_key: str, key_val, ref_val: str) -> str:
        norm     = os.path.normpath(ref_json_abs)
        dict_key = (norm, ref_key, ref_val)

        if dict_key not in self._ref_dict:
            # Load raw rows if not yet loaded
            if norm not in self._ref_cache:
                try:
                    with open(norm, 'r', encoding='utf-8-sig') as f:
                        raw = json.load(f)
                    if isinstance(raw, list):
                        rows = [r for r in raw if isinstance(r, dict)]
                    elif isinstance(raw, dict):
                        rows = []
                        for v in raw.values():
                            if isinstance(v, list):
                                rows = [r for r in v if isinstance(r, dict)]
                                break
                    else:
                        rows = []
                    self._ref_cache[norm] = rows
                except Exception:
                    self._ref_cache[norm] = []

            # Build lookup dict: {str(key_col_value): str(val_col_value)}
            rows = self._ref_cache[norm]
            self._ref_dict[dict_key] = {
                str(row[ref_key]): str(row[ref_val])
                for row in rows
                if ref_key in row and ref_val in row
            }

        return self._ref_dict[dict_key].get(str(key_val), "")

    def invalidate_ref_cache(self, ref_json_abs: str | None = None) -> None:
        if ref_json_abs is None:
            self._ref_cache.clear()
            self._ref_dict.clear()
        else:
            norm = os.path.normpath(ref_json_abs)
            self._ref_cache.pop(norm, None)
            for k in [k for k in self._ref_dict if k[0] == norm]:
                del self._ref_dict[k]

    # ──────────────────────────────────────────────
    # Load
    # ──────────────────────────────────────────────

    def load_json(self, file_path):
        """
        Load a .json file.
        Supports:
          - Array of objects: [{...}, {...}] → single table named after file stem
          - Object of arrays: {"skills": [...], ...} → one table per key
          - Single object: {...} → wrap as [{...}]

        Nested array-of-objects → extracted as sub_table with FK column.
        Nested array-of-primitives → joined as comma-separated string.
        """
        self.json_path = file_path
        self.need_config_alert = False
        self.tables = {}
        self.sub_tables = {}
        self._search_index = {}

        file_key = os.path.normpath(file_path)
        if file_key not in self._full_config:
            self._full_config[file_key] = {}
        self.config = self._full_config[file_key]

        try:
            with open(file_path, 'r', encoding='utf-8-sig') as f:
                raw = json.load(f)
        except Exception as e:
            print(f"載入 JSON 失敗: {e}")
            raise

        stem = os.path.splitext(os.path.basename(file_path))[0]

        # Detect top-level structure
        if isinstance(raw, list):
            # Array of objects (or empty)
            source_tables = {stem: raw}
        elif isinstance(raw, dict):
            # Check if values are lists → multi-table object
            # Otherwise single object
            has_list_values = any(isinstance(v, list) for v in raw.values())
            if has_list_values:
                source_tables = {}
                for k, v in raw.items():
                    if isinstance(v, list):
                        source_tables[k] = v
                    else:
                        # scalar value at top level – wrap into a single-row table
                        source_tables[k] = [{"value": v}]
            else:
                # Single object
                source_tables = {stem: [raw]}
        else:
            # Scalar – unlikely but handle gracefully
            source_tables = {stem: [{"value": raw}]}

        # Process each table
        for table_name, rows in source_tables.items():
            if not isinstance(rows, list):
                rows = [rows] if rows is not None else []

            main_rows = []
            sub_accumulators = {}  # {key_name: [rows]}

            # Determine primary key (first column of the first non-empty row)
            pk_key = None

            for row in rows:
                if not isinstance(row, dict):
                    row = {"value": row}

                flat_row = {}
                for k, v in row.items():
                    if isinstance(v, list):
                        if v and isinstance(v[0], dict):
                            # nested array-of-objects → sub_table
                            sub_accumulators.setdefault(k, [])
                            # Will be handled below with FK injection
                            flat_row[k] = None  # placeholder, removed later
                        else:
                            # array-of-primitives → join as string
                            flat_row[k] = ", ".join(str(x) for x in v)
                    elif isinstance(v, dict):
                        # Flatten nested dict with dot notation
                        for nk, nv in v.items():
                            flat_row[f"{k}.{nk}"] = nv
                    else:
                        flat_row[k] = v

                main_rows.append(flat_row)

            # Build main DataFrame
            if main_rows:
                df = pd.DataFrame(main_rows)
                # Remove sub_table placeholder columns (all-None means it was a nested array)
                cols_to_drop = [col for col in df.columns
                                if col in sub_accumulators and df[col].isna().all()]
                if cols_to_drop:
                    df = df.drop(columns=cols_to_drop)
                # Single-pass conversion: fillna + str + strip "None" artifacts
                df = df.where(df.notna(), "").astype(str).replace("None", "")
                df, _ = self._drop_empty_rows(df)
            else:
                df = pd.DataFrame()

            # Determine pk_key for FK injection
            if len(df.columns) > 0:
                pk_key = df.columns[0]

            self.tables[table_name] = df

            # Build search index for main table
            self._build_search_index(table_name, df)

            # Build sub_tables
            for key_name, _ in sub_accumulators.items():
                sub_full_name = f"{table_name}.{key_name}"
                sub_row_list = []
                for row in rows:
                    if not isinstance(row, dict):
                        continue
                    nested = row.get(key_name, [])
                    if not isinstance(nested, list):
                        continue
                    # Get FK value from parent row
                    fk_val = ""
                    if pk_key and pk_key in row:
                        fk_val = str(row[pk_key]) if row[pk_key] is not None else ""
                    for sub_row in nested:
                        if not isinstance(sub_row, dict):
                            sub_row = {"value": sub_row}
                        new_row = {pk_key: fk_val}
                        new_row.update(sub_row)
                        sub_row_list.append(new_row)

                if sub_row_list:
                    sub_df = pd.DataFrame(sub_row_list)
                    sub_df = sub_df.where(sub_df.notna(), "").astype(str).replace("None", "")
                    sub_df, _ = self._drop_empty_rows(sub_df)
                else:
                    cols = [pk_key] if pk_key else []
                    sub_df = pd.DataFrame(columns=cols)

                self.sub_tables[sub_full_name] = sub_df

                # Build search index for sub table
                self._build_search_index(sub_full_name, sub_df)

            # Init config for this table if not present
            if table_name not in self.config:
                self.config[table_name] = self._make_default_table_config(table_name, df)
                self.need_config_alert = True
            else:
                # Ensure sub_tables key exists
                self.config[table_name].setdefault("sub_tables", {})

            # Init sub_table configs
            for key_name in sub_accumulators:
                sub_full_name = f"{table_name}.{key_name}"
                sub_cfg = self.config[table_name].setdefault("sub_tables", {})
                if key_name not in sub_cfg:
                    sub_df = self.sub_tables.get(sub_full_name, pd.DataFrame())
                    sub_cfg[key_name] = {
                        "foreign_key": pk_key or "",
                        "columns": {col: {"type": "string"} for col in sub_df.columns}
                    }
                    self.need_config_alert = True

        if self.need_config_alert:
            self.save_config()

        self._add_recent_file(file_path)
        self.dirty = False
        self.dirty_cells.clear()

    def _make_default_table_config(self, table_name, df):
        pk = df.columns[0] if len(df.columns) > 0 else ""
        cls = df.columns[0] if len(df.columns) > 0 else ""
        return {
            "use_icon": False,
            "image_path": "",
            "classification_key": cls,
            "primary_key": pk,
            "columns": {col: {"type": self._infer_col_type(df, col)} for col in df.columns},
            "sub_tables": {}
        }

    # ──────────────────────────────────────────────
    # Save
    # ──────────────────────────────────────────────

    def save_json(self, path=None):
        """
        Save back to JSON.
        - Backup original as filename.json.bak
        - Write to .tmp then rename (safe save)
        - Reassemble sub_tables back into nested arrays
        """
        target = path or self.json_path
        if not target:
            return

        file_key = os.path.normpath(self.json_path)
        stem = os.path.splitext(os.path.basename(target))[0]

        # Determine if original was multi-table object or single-table array
        # We reconstruct based on what tables we have
        table_names = list(self.tables.keys())

        def _df_to_records(table_name, df):
            """Convert DataFrame to list of dicts, reassembling sub_tables."""
            cfg = self.config.get(table_name, {})
            pk_key = cfg.get("primary_key", df.columns[0] if len(df.columns) > 0 else None)

            # Find all sub_tables for this table
            sub_table_keys = [s for s in self.sub_tables if s.startswith(table_name + ".")]

            records = []
            for _, row in df.iterrows():
                record = {}
                for col in df.columns:
                    val = row[col]
                    record[col] = _convert_value(table_name, col, val)

                # Reassemble sub_tables
                for sub_full in sub_table_keys:
                    key_name = sub_full.rsplit(".", 1)[1]
                    sub_df = self.sub_tables[sub_full]
                    sub_cfg = cfg.get("sub_tables", {}).get(key_name, {})
                    fk_key = sub_cfg.get("foreign_key", pk_key)

                    if fk_key and fk_key in sub_df.columns and pk_key:
                        pk_val = str(row[pk_key]) if pk_key in row.index else ""
                        mask = sub_df[fk_key].astype(str) == pk_val
                        matched = sub_df[mask]
                    else:
                        matched = pd.DataFrame()

                    sub_records = []
                    for _, sub_row in matched.iterrows():
                        sub_rec = {}
                        for col in sub_df.columns:
                            if col == fk_key:
                                continue  # Remove FK column from output
                            sub_rec[col] = _convert_value_sub(table_name, key_name, col, sub_row[col])
                        sub_records.append(sub_rec)

                    record[key_name] = sub_records

                records.append(record)
            return records

        def _convert_value(table_name, col, val):
            cfg = self.config.get(table_name, {})
            col_type = cfg.get("columns", {}).get(col, {}).get("type", "string")
            return _coerce(val, col_type)

        def _convert_value_sub(table_name, sub_name, col, val):
            cfg = self.config.get(table_name, {})
            col_type = cfg.get("sub_tables", {}).get(sub_name, {}).get("columns", {}).get(col, {}).get("type", "string")
            return _coerce(val, col_type)

        def _coerce(val, col_type):
            if val == "" or val is None:
                if col_type in ("int", "float"):
                    return None
                return val
            try:
                if col_type == "int":
                    return int(float(str(val)))
                elif col_type == "float":
                    fv = float(str(val))
                    return fv
                elif col_type == "bool":
                    return str(val).lower() in ('true', '1', 'yes')
                else:
                    return str(val)
            except (ValueError, TypeError):
                return str(val) if val is not None else ""

        # Build output structure
        if len(table_names) == 1 and table_names[0] == stem:
            # Single table named after file → output as array
            output = _df_to_records(table_names[0], self.tables[table_names[0]])
        elif len(table_names) == 1:
            output = _df_to_records(table_names[0], self.tables[table_names[0]])
        else:
            # Multi-table → output as object
            output = {}
            for tn in table_names:
                output[tn] = _df_to_records(tn, self.tables[tn])

        # Backup
        if os.path.exists(target):
            bak = target + ".bak"
            try:
                import shutil
                shutil.copy2(target, bak)
            except Exception as e:
                print(f"備份失敗: {e}")

        # Write to .tmp then rename
        tmp = target + ".tmp"
        try:
            with open(tmp, 'w', encoding='utf-8') as f:
                json.dump(output, f, indent=2, ensure_ascii=False)
            os.replace(tmp, target)
        except Exception as e:
            if os.path.exists(tmp):
                try:
                    os.remove(tmp)
                except Exception:
                    pass
            print(f"儲存失敗: {e}")
            raise

        self.dirty = False
        self.dirty_cells.clear()

    # ──────────────────────────────────────────────
    # Schema inference
    # ──────────────────────────────────────────────

    def infer_schema(self, table_name, df):
        """Scan DataFrame and infer column types (string/int/float/bool/enum)."""
        result = {}
        for col in df.columns:
            result[col] = {"type": self._infer_col_type(df, col)}
        return result

    def _infer_col_type(self, df, col):
        if col not in df.columns or df[col].empty:
            return "string"
        # Sample at most 500 rows to keep inference fast on large datasets
        sample = df[col].dropna()
        if len(sample) > 500:
            sample = sample.iloc[:500]
        sample = sample[sample.astype(str).str.strip() != ""]
        if sample.empty:
            return "string"

        str_vals = sample.astype(str).str.strip()
        lower_vals = str_vals.str.lower()

        # bool detection — vectorized isin (O(n) hash lookup)
        bool_set = {"true", "false", "1", "0", "yes", "no"}
        if lower_vals.isin(bool_set).all():
            return "bool"

        # int detection — pd.to_numeric is ~50x faster than apply()
        numeric = pd.to_numeric(str_vals, errors='coerce')
        if numeric.notna().all():
            # Check all are integers (no decimal point in original strings)
            if str_vals.str.match(r'^-?\d+$').all():
                return "int"
            return "float"

        # enum detection: ≤12 unique values AND has repeated values
        unique = str_vals.unique()
        if len(unique) <= 12 and len(str_vals) >= 2 and len(unique) < len(str_vals):
            return "enum"

        return "string"

    # ──────────────────────────────────────────────
    # Cell update
    # ──────────────────────────────────────────────

    def update_cell(self, table_name, row_idx, col_name, value):
        """Update a cell in tables or sub_tables. Detects which dict by checking both."""
        if table_name in self.tables:
            target_df = self.tables[table_name]
            col_type = self.config.get(table_name, {}).get("columns", {}).get(col_name, {}).get("type", "string")
        elif table_name in self.sub_tables:
            target_df = self.sub_tables[table_name]
            # Determine col type from config
            parts = table_name.rsplit(".", 1)
            if len(parts) == 2:
                master_name, sub_name = parts
                col_type = (self.config.get(master_name, {})
                            .get("sub_tables", {})
                            .get(sub_name, {})
                            .get("columns", {})
                            .get(col_name, {})
                            .get("type", "string"))
            else:
                col_type = "string"
        else:
            return

        # Type coercion
        try:
            if col_type == "int":
                value = int(float(str(value)))
            elif col_type == "float":
                value = float(str(value))
            elif col_type == "bool":
                if isinstance(value, str):
                    value = value.lower() in ['true', '1', 'yes']
                else:
                    value = bool(value)
            else:
                value = str(value) if value is not None else ""
        except (ValueError, TypeError):
            value = str(value) if value is not None else ""

        if row_idx in target_df.index and col_name in target_df.columns:
            target_df.at[row_idx, col_name] = value
            self.dirty = True
            self.dirty_cells.add((table_name, row_idx, col_name))

    # ──────────────────────────────────────────────
    # Column management
    # ──────────────────────────────────────────────

    def add_column(self, table_name, col_name, col_type="string", default=""):
        """Add a new column to a table."""
        if table_name in self.tables:
            self.tables[table_name][col_name] = default
            cfg = self.config.setdefault(table_name, {}).setdefault("columns", {})
            cfg[col_name] = {"type": col_type}
        elif table_name in self.sub_tables:
            self.sub_tables[table_name][col_name] = default
            parts = table_name.rsplit(".", 1)
            if len(parts) == 2:
                master_name, sub_name = parts
                cfg = (self.config.setdefault(master_name, {})
                       .setdefault("sub_tables", {})
                       .setdefault(sub_name, {})
                       .setdefault("columns", {}))
                cfg[col_name] = {"type": col_type}
        self.dirty = True

    def delete_column(self, table_name, col_name):
        """Delete a column from a table."""
        if table_name in self.tables:
            if col_name in self.tables[table_name].columns:
                self.tables[table_name].drop(columns=[col_name], inplace=True)
            cfg = self.config.get(table_name, {}).get("columns", {})
            cfg.pop(col_name, None)
        elif table_name in self.sub_tables:
            if col_name in self.sub_tables[table_name].columns:
                self.sub_tables[table_name].drop(columns=[col_name], inplace=True)
            parts = table_name.rsplit(".", 1)
            if len(parts) == 2:
                master_name, sub_name = parts
                cfg = (self.config.get(master_name, {})
                       .get("sub_tables", {})
                       .get(sub_name, {})
                       .get("columns", {}))
                cfg.pop(col_name, None)
        self.dirty = True

    def rename_column(self, table_name, old_name, new_name):
        """Rename a column."""
        if table_name in self.tables:
            df = self.tables[table_name]
            if old_name in df.columns:
                df.rename(columns={old_name: new_name}, inplace=True)
            cfg = self.config.get(table_name, {}).get("columns", {})
            if old_name in cfg:
                cfg[new_name] = cfg.pop(old_name)
        elif table_name in self.sub_tables:
            df = self.sub_tables[table_name]
            if old_name in df.columns:
                df.rename(columns={old_name: new_name}, inplace=True)
            parts = table_name.rsplit(".", 1)
            if len(parts) == 2:
                master_name, sub_name = parts
                cfg = (self.config.get(master_name, {})
                       .get("sub_tables", {})
                       .get(sub_name, {})
                       .get("columns", {}))
                if old_name in cfg:
                    cfg[new_name] = cfg.pop(old_name)
        self.dirty = True

    # ──────────────────────────────────────────────
    # Utility
    # ──────────────────────────────────────────────

    @staticmethod
    def _drop_empty_rows(df):
        """Remove rows where all values are blank strings."""
        if df.empty:
            return df, pd.Series(dtype=bool)
        try:
            stripped = df.apply(lambda s: s.astype(str).str.strip())
            mask = ~stripped.eq("").all(axis=1)
            return df[mask].reset_index(drop=True), mask
        except Exception:
            return df.reset_index(drop=True), pd.Series([True] * len(df))

    def cleanup(self):
        self.tables.clear()
        self.sub_tables.clear()
        self._search_index.clear()
        gc.collect()

    def __del__(self):
        self.cleanup()
