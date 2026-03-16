Markdown

# Factory Log Aggregator & Dashboard Generator

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## 專案概述 (Project Overview)
本專案為一套工業生產環境的自動化數據彙整工具 。旨在解決測試製程原始日誌 (Raw logs) 格式破碎、跨設備數據孤島 (Data Silos) 之問題。透過測試Program之數據，將分散的txt檔案自動彙整為視覺化的 Excel  Dashboard。

## 核心功能 (Core Features)
* **多源數據彙整 (Multi-source Aggregation)**：自動遍歷指定目錄，解析不同設備產出的異質 Log 格式。
* **數據清洗與標準化 (ETL Process)**：執行去重 (Deduplication)、缺失值處理及格式標準化作業。
* **自動化儀表板 (Automated Dashboarding)**：利用 `pandas` 與 `XlsxWriter` 引擎，生成內含樞紐分析與統計圖表之 Excel 報表。
* **二級思考架構 (Second-Order Thinking Implementation)**：預留錯誤捕捉機制，確保在 Log 格式突發性變動時仍能維持系統穩定性 (System Robustness)。

## 技術棧 (Technology Stack)
* **Language**: Python 3.x
* **Core Libraries**: 
    * `pandas`: 用於大規模數據操縱與分析 (Data Manipulation)。
    * `openpyxl` / `XlsxWriter`: 驅動 Excel 檔案生成的進階引擎。
    * `os` / `glob`: 執行系統層級的檔案檢索與路徑管理。

## 快速開始 (Quick Start)

### 1. 環境配置 (Environment Setup)
```bash
pip install pandas openpyxl xlsxwriter
2. 執行腳本 (Execution)
將 Log 檔案放置於專案預設的 /raw_data 目錄下，執行：

Bash

python aggregator.py
風險審計與限制 (Risk Audit & Constraints)
性能瓶頸 (Performance Bottleneck)：當 Log 數據量超過百萬等級 (Million-scale) 時，pandas 的記憶體消耗將呈線性增長，建議改採分塊讀取 (Chunking) 策略。

Schema 飄移 (Schema Drift)：若工廠端韌體更新導致 Log 欄位變更，腳本將拋出關鍵字錯誤 (KeyError)，需手動調整 config.json 中的對應表。

安全合規 (Compliance)：本腳本僅處理本地數據，不涉及雲端上傳，符合一般工廠對資安 (Information Security) 的內網物理隔離要求。
