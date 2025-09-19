# Office Add‑ins MCP Server
[English Version](./README.md)

## 簡介

用於探索和管理跨 Word、Excel、PowerPoint、Outlook 和 Teams 的 Microsoft Office 外掛程式的 Model Context Protocol (MCP) 伺服器。此伺服器讓 AI Agent 能夠搜尋外掛程式並檢索詳細的中繼資料、安裝或卸載外掛程式、處理自訂外掛程式的提交、驗證和發佈。

這個儲存庫提供一個基於 **Model Context Protocol (MCP)** 的完整伺服器實作，使用官方 Python SDK。MCP 標準化大型語言模型（LLM）與外部資料來源和工具的通訊方式。`FastMCP` 類別封裝了 MCP 的複雜性，讓開發者能夠以最少的樣板程式碼將普通的 Python 函式公開為 MCP 工具或資源。

目前，此伺服器提供基本的外掛程式詳細資訊檢索功能，並規劃在未來版本中提供全面的外掛程式管理功能（請參閱開發路線圖部分）。

## 安裝與設定（本地伺服器）

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
# 選項 1：使用已安裝的腳本（推薦）
uv run office-addins-mcp-server

# 選項 2：使用特定傳輸類型
uv run office-addins-mcp-server --transport stdio
uv run office-addins-mcp-server --transport sse
uv run office-addins-mcp-server --transport http

# 選項 3：使用 uv run 直接路徑
uv run python office_addins_mcp_server/server.py --transport stdio

# 選項 4：啟動虛擬環境後
source .venv/bin/activate
python office_addins_mcp_server/server.py --transport stdio
```

**傳輸類型：**
- **`stdio`**（預設）：標準輸入/輸出傳輸，非常適合本地測試和 CLI 整合
- **`sse`**：Server-Sent Events 傳輸，適合 Web 服務部署
- **`http`**：串流 HTTP 傳輸，適用於基於 HTTP 的整合

## 🧪 實驗性遠端伺服器

> **⚠️ 實驗性功能**：此 MCP 伺服器的遠端實例僅供測試使用。此服務不適用於生產環境，可能有有限的運行時間、速率限制，或可能在無預警的情況下終止服務。

**遠端 MCP 端點**：`https://app-api-gmqmpcvoduxtc.azurewebsites.net/addins/mcp`

**使用方式**：您可以使用此端點測試 MCP 協定，但請為任何正式工作部署您自己的實例。

## 測試伺服器

要驗證伺服器是否工作，請使用相容 MCP 的客戶端連接並調用 `get_addin_details` 工具。官方 SDK 通過 `uv run mcp` 提供 CLI 工具。例如：

```bash
# 在開發模式下啟動伺服器
uv run mcp dev office_addins_mcp_server/server.py

# 或安裝到 Claude Desktop
uv run mcp install office_addins_mcp_server/server.py

# 或使用已安裝的腳本
uv run office-addins-mcp-server
```

運行後，您可以使用 Claude Desktop 或 MCP Inspector 連接來測試工具。請參考官方文檔編寫自定義客戶端。


## Azure App Service 部署（快速開始）
使用 Azure Developer CLI (azd) 將 MCP 伺服器部署為 Azure App Service 上的 Web 服務。提供具有自動擴展和監控功能的生產就緒 HTTP 端點。

```bash
# 安裝 Azure Developer CLI
curl -fsSL https://aka.ms/install-azd.sh | bash

# 使用 Azure 進行身份驗證
azd auth login

# 部署到 Azure（首次）
azd up

# 變更後重新部署
azd deploy
```

📖 **[完整 Azure App Service 部署指南 →](./AZURE_APP_SERVICE.md)**

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