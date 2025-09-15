# Office Add‑ins MCP Server
[English Version](./README.md)

## 簡介

這個儲存庫提供一個基於 **Model Context Protocol (MCP)** 的簡易伺服器，
使用官方 Python SDK 中的 **FastMCP** 類別快速建構。MCP 是一套標準化
協議，用來讓大型語言模型（LLM）與外部資源或工具溝通。
FastMCP 將複雜的 MCP 協議封裝在高階介面之中，讓開發者只需利用少數
修飾器即可把普通的 Python 函式轉換成 MCP 工具或資源。
此伺服器提供一個工具，可依據 **asset ID** 從 Microsoft Office Add‑ins API
查詢外掛詳細資訊。藉由這個 MCP 伺服器，具備 MCP 客戶端功能的 LLM
應用程式（例如 Claude Desktop）便能輕鬆呼叫 API 並取得資料。

## Features / 功能

- **標準化介面：** MCP 伺服器是一種標準化 API，使 LLM 能透過統一的界面
  存取內部工具或資料來源。
- **FastMCP 簡化開發：** FastMCP 採用修飾器與型別註解自動產生工具的
  schema，減少樣板程式碼，並支援同步與非同步函式。
- **Server‑Sent Events (SSE)：** 本範例使用 SSE 傳輸方式，可以作為 Web
  服務部署；若要在本地工具模式使用 STDIO 傳輸，只需修改
  `mcp.run()` 的 `transport` 參數。
- **非同步 API 呼叫：** 工具使用 `httpx.AsyncClient` 非同步呼叫
  Office Add‑ins API，確保 MCP 伺服器在處理網路請求時不會阻塞其他工作。

## 安裝與使用

以下步驟說明如何在本地環境安裝與執行此 MCP 伺服器。請先確保安裝有
Python 3.8 或以上版本。

### 1. 取得原始碼

```bash
git clone <repository-url> office-addins-mcp-server
cd office-addins-mcp-server
```

### 2. 使用 uv 管理相依套件

此專案建議使用 [uv](https://docs.astral.sh/uv/) 來管理 Python 專案。uv
將專案的依賴和鎖檔集中在 `pyproject.toml` 和 `uv.lock`，無需自行建立
虛擬環境。安裝 uv 後，執行以下步驟：

```bash
# 在專案根目錄初始化 uv 專案
uv init

# 新增依賴：MCP SDK 與 httpx
uv add "mcp[cli]"
uv add httpx

# 安裝所有依賴（當已存在 uv.lock 時）
uv install
```

執行 `uv init` 會產生 `pyproject.toml` 和 `uv.lock`。完成後，您可以使用
`uv run` 來執行 Python 程式，例如：

```bash
uv run python src/server.py
```

在 `pyproject.toml` 中，這兩個依賴將被列為必需套件：

* `mcp[cli]` – 官方 Model Context Protocol SDK，包括 FastMCP 伺服器類別與
  CLI 工具。
* `httpx` – 非同步 HTTP 用戶端，用來呼叫 Office Add‑ins API。

### 3. 執行 MCP 伺服器

進入專案根目錄後，執行下列指令啟動伺服器：

```bash
python src/server.py
```

預設狀態下伺服器會以 SSE 傳輸方式在 `0.0.0.0:8000` 上監聽。欲
修改埠號或使用 STDIO 模式，可在 `src/server.py` 中調整
`mcp.run()` 的參數。例如：

```python
mcp.run(transport="stdio")  # 用於本地 CLI 客戶端
```

### 4. 測試伺服器

要驗證伺服器是否正常工作，可以透過支援 MCP 的客戶端連線並呼叫
`get_addin_details` 工具。官方 SDK 提供多種方法，包括 `uv run mcp` 指令
或撰寫程式使用 `mcp.client`。以下示意使用命令列工具：

```bash
# 安裝 mcp[cli] 之後，使用 uv 執行伺服器
uv run mcp dev src/server.py
# 或安裝到 Claude Desktop 等應用程式
uv run mcp install src/server.py
```

上述指令會啟動本範例伺服器，並讓支援 MCP 的應用程式（例如
Claude Desktop 或 MCP Inspector）可以連線測試。使用者亦可參考
官方文件撰寫客戶端程式來呼叫工具。

## 專案結構

```text
office-addins-mcp-server/
├── docs/
│   └── execution_plan.md  # 英文執行計劃
├── src/
│   ├── __init__.py       # 使 src 成為可匯入的模組
│   └── server.py         # MCP 伺服器實作
├── requirements.txt       # Python 相依套件清單
├── README.md              # 英文專案說明
└── README_zh_tw.md        # 專案中文說明
```

## 後續改進

此專案目前僅實作一個簡單工具，可依需求新增更多工具或資源，如：

* **快取機制：** 為頻繁查詢的 add‑in 資訊加入快取，以減少對
  Office API 的重複呼叫。
* **錯誤處理強化：** 改進對異常情況的處理與回報機制。
* **其他工具：** 擴充為不同 Office API 功能（例如搜尋外掛、列出使用者
  安裝的外掛等）。

歡迎貢獻者提供建議或提交拉取請求，共同完善此 MCP 伺服器。
