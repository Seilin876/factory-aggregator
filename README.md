# Factory Data Aggregator

自動化合併分散式產線日誌（.txt）至單一彙總報表（.csv）。

---

## 🚀 Dry Run 測試步驟

在明天上傳原始資料前，您可以先在您的 Mac Mini 環境中進行 Dry Run 測試：

### 1. 環境準備
此腳本依賴 `pandas` 處理大量數據。請先安裝必要套件：
```bash
pip install pandas
```

### 2. 準備測試資料
腳本會尋找符合 `{Date}_{IP}.txt` 格式的檔案。您可以手動建立一個測試檔：
```bash
mkdir -p data_source
echo "col1,col2\n1,2\n3,4" > data_source/20260209_10_184_137_46.txt
```

### 3. 修改設定檔
編輯 `config.ini`，將來源路徑指向剛才建立的測試目錄：
```ini
[Path]
Source_Folder = ./data_source
Output_Folder = ./Output_Summaries
```

### 4. 執行腳本
執行主程式（預設處理昨天的數據）：
```bash
python aggregator.py
```
執行完成後，檢查 `./Output_Summaries/` 目錄是否生成了合併後的 `.csv` 檔案。

---

## 🛠 技術規格
- **輸入**: `\\10.198.112.103\Noise_Data` 中的日誌檔案。
- **輸出**: 每日彙總 CSV。
- **映射邏輯**: 透過 `config.ini` 中的 `[Device_Mapping]` 區塊自動為數據注入產線與設備名稱。
