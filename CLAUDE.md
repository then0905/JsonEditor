# SkillExcelEditor — 開發規範

## 專案概述
- CTk 框架層 + 原生 tk 熱路徑元件的混合架構 Excel 編輯器
- Key files: `main.py` (UI), `data_manager.py` (data layer), `config.json` (per-Excel config)
- Excel sheets named `*.json` → 母表; sheets with `#` → 子表 (e.g. `SkillData.json#Operation`)
- Config 以 Excel 路徑為 key 存於 `config.json`，支援 classification_key / primary_key / column types / sub_sheets
- DataManager 只用 openpyxl 讀寫 Excel（不依賴 pandas I/O），儲存時保留儲存格格式
- UI: SheetEditor（左=分類、中=項目清單、右=母表欄位 + 子表 TabView）
- 支援外部連結文字表（.xlsx，TextID/TextContent 欄位）
- 開發目標：穩定、流暢、生產品質的遊戲資料編輯器（零 bug、零 UI 卡頓）

---

## 架構模式（必須遵守）

### 1. CTk vs tk 分界
- 迴圈 / 熱路徑（子表行、標題、scrollbar、空資料標籤）**一律用原生 tk/ttk**，搭配暗色常數（`_CELL_BG`, `_CELL_FG`, `_CELL_BORDER`, `_CELL_FOCUS_BORDER`, `_CELL_FONT`）
- CTk 元件**僅限**一次性建立且不在迴圈中的框架層級（CTkTabview、CTkFrame 外殼、CTkInputDialog、CTkToplevel）
- **禁止**在迴圈內使用：CTkEntry, CTkOptionMenu, CTkCheckBox, CTkLabel, CTkButton, CTkTextbox, CTkScrollbar

### 2. Suppress-flag 模式（取代 unbind/rebind）
- callback 內部讀取 mutable context dict，`suppress=True` 時直接跳過
- 子表行：`row_frame._ctxs[col] = {"suppress": bool, "sheet": str, "row_idx": int, "col": str, "key": str}`
- 母表：`self._master_suppress` flag，在 `_on_field_change` / `_on_linked_field_change` / `_on_linked_field_change_tb` 開頭檢查
- **嚴禁** `trace_remove` / `trace_add` / `unbind` / `bind` 循環來避免資料注入觸發 callback

### 3. Linked key 快取
- `textbox._linked_key` 在 `_update_editor_data` 注入值時同步設定
- callback 直接讀取 `_linked_key`，**不碰** key_entry 的 `configure(state=...)`

### 4. Tab lazy-load
- `_ensure_editor()` / `_on_main_tab_changed()`：只在 tab 被選中時才建立 SheetEditor
- `refresh_ui()` 差異更新 Tabs，只建立第一個可見 tab

### 5. Freeze/thaw 批次更新
- 子表 `frames['_freeze'] = True` 阻擋 `_update_widths` 等佈局回呼
- 批次完成後 `_freeze = False` + 呼叫一次 `_update_widths()` flush

### 6. 背景線程 I/O
- `load_file` / `save_file` 在 `threading.Thread(daemon=True)` 執行
- 主線程顯示 loading dialog（CTkToplevel + grab_set）
- 完成後透過 `self.after(0, callback)` 回到主線程

### 7. DataManager I/O 規範
- `load_excel`：用 `_ws_to_dataframe(ws)` 從 openpyxl Worksheet 直接轉全字串 DataFrame（單次開檔，不用 pd.read_excel）
- `load_external_text`：用 `load_workbook(read_only=True)` + `iter_rows()` 串流讀取
- `_save_external_text`：按 sheet 分組 modifications → 每 sheet 建一次 `{key_str: row_idx}` 索引 → O(1) 更新
- DataFrame 運算限定在 `data_manager.py`，UI 層不做 pandas 操作

### 8. No-op 跳過
- `_resize_text_cell`：比對 `tw._last_lines`，相同則不呼叫 `configure(height=...)`
- `_update_widths`：檢查 `_freeze` flag；`<Configure>` 回呼比對 `_prev_width`
- 所有 configure / resize 前先比對舊值，避免無意義的重繪

### 9. 查找效能
- 使用 dict / set 做 O(1) 查找，**禁止**在迴圈內線性掃描
- 例：`text_dict`、`row_index`、`cls_buttons`、`item_buttons`

---

## 測試要求
- 新功能上線前用生產級資料量（14+ sheets, 1000+ rows）測試
- 確認無 UI 凍結（>200ms 主線程阻塞）
- 切換分類 → 點選項目 → 子表行顯示正確、enum 下拉可用、bool checkbox 可勾選
- 連結欄位修改文字 → key 正確回寫
- 儲存 Excel + 文字表 → 重新載入 → 資料無損
- 視窗縮放 → 母表/子表高度正確自適應
