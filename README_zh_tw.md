# Office Add‑ins MCP Server
[English Version](./README.md)

## 簡介

用於探索和管理跨 Word、Excel、PowerPoint、Outlook 和 Teams 的 Microsoft Office 外掛程式的 Model Context Protocol (MCP) 伺服器。

這個儲存庫提供一個基於 **Model Context Protocol (MCP)** 的伺服器實作，使用官方 Python SDK。MCP 標準化大型語言模型（LLM）與外部資料來源和工具的通訊方式。`FastMCP` 類別封裝了 MCP 的複雜性，讓開發者能夠以最少的樣板程式碼將普通的 Python 函式公開為 MCP 工具或資源。

目前，此伺服器提供基本的外掛程式詳細資訊檢索功能，並規劃在未來版本中提供全面的外掛程式管理功能（請參閱開發路線圖部分）。

## 安裝與設定

此專案使用 [uv](https://docs.astral.sh/uv/) 來管理 Python 依賴項目和虛擬環境，並包含 `pyproject.toml` 配置文件和 `uv.lock` 文件以確保跨不同環境的可重現構建。

### 專案設定

1. **下載專案**：

   ```bash
   git clone <repository-url> office-addins-mcp-server
   cd office-addins-mcp-server
   ```

2. **使用 uv 安裝**：

   ```bash
   # 這將創建虛擬環境並根據 pyproject.toml 和 uv.lock 安裝所有套件
   uv sync
   ```

### 執行伺服器

您可以用幾種方式執行伺服器：

```bash
# 選項 1：使用 uv run（推薦）
uv run python src/server.py

# 選項 2：激活虛擬環境後
source .venv/bin/activate
python src/server.py

# 選項 3：使用已安裝的腳本（待實現）
uv run office-addins-mcp-server
```

## 伺服器配置

伺服器支援多種傳輸類型，可以使用環境變數進行配置：

### 傳輸配置

在專案根目錄創建 `.env` 文件來配置伺服器：

```bash
# .env 文件

# 傳輸配置
TRANSPORT=stdio  # 預設：標準輸入/輸出
# TRANSPORT=sse   # Server-Sent Events 用於 Web 部署
# TRANSPORT=http  # 串流 HTTP 傳輸

# 網路配置（用於 SSE 和 HTTP 傳輸）
HOST=0.0.0.0     # 預設：監聽所有介面
PORT=8000        # 預設：Port 8000
PATH_PREFIX=/    # 預設：根路徑（用於未來的 HTTP 路由）
SSE_PATH=/sse    # 預設：SSE 端點路徑
```

**傳輸類型：**
- **`stdio`**（預設）：標準輸入/輸出傳輸，非常適合本地測試和 CLI 整合
- **`sse`**：Server-Sent Events 傳輸，適合 Web 服務部署
- **`http`**：串流 HTTP 傳輸，適用於基於 HTTP 的整合

### 配置選項

**可用的環境變數：**
- `TRANSPORT`：傳輸類型（`stdio`、`sse`、`http`）- 預設：`stdio`
- `HOST`：綁定的主機 - 預設：`0.0.0.0`（所有介面）
- `PORT`：監聽的端口 - 預設：`8000`
- `PATH_PREFIX`：HTTP 路由的路徑前綴 - 預設：`/`
- `SSE_PATH`：SSE 端點路徑 - 預設：`/sse`

### 預設設定

預設情況下，伺服器：
- 使用 STDIO 傳輸
- 監聽 `0.0.0.0:8000`（用於 SSE/HTTP 傳輸）
- 使用根路徑（`/`）進行路由
- 如果存在則從 `.env` 文件載入所有配置

## 測試伺服器

要驗證伺服器是否工作，請使用相容 MCP 的客戶端連接並調用 `get_addin_details` 工具。官方 SDK 通過 `uv run mcp` 提供 CLI 工具。例如：

```bash
# 在開發模式下啟動伺服器
uv run mcp dev src/server.py

# 或安裝到 Claude Desktop
uv run mcp install src/server.py
```

運行後，您可以使用 Claude Desktop 或 MCP Inspector 連接來測試工具。請參考官方文檔編寫自定義客戶端。

## 最新公告

🎉 2025-09-13：Office Add-ins MCP 伺服器建立。

## 開發路線圖

以下功能已規劃於未來版本開發：

1. **外掛搜尋功能** - 實作完整的 Office 外掛搜尋功能
2. **搜尋外掛功能提示** - 使用智能提示與篩選增強搜尋體驗
3. **OAuth2 身份驗證** - 實作 Microsoft Graph API 的安全身份驗證
4. **顯示已安裝外掛** - 顯示使用者在各 Office 應用程式中已安裝的外掛
5. **安裝/卸載外掛** - 程式化安裝與移除外掛功能
6. **提交自訂外掛** - 支援將自訂外掛提交至 Office Store
7. **驗證自訂外掛** - 自動化驗證與合規性檢查自訂外掛
8. **發布自訂外掛** - 簡化自訂外掛的發布流程
9. **M365 管理員推送外掛** - 允許 Microsoft 365 管理員集中部署外掛

歡迎貢獻者提供建議或提交拉取請求，共同完善此 MCP 伺服器。