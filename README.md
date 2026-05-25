# MCH Assistant Web Service

門諾醫院 AI 語音助理 — 語音轉文字並生成專業醫療報告（支援分段錄音、Word 模板、Email 發送）。

## 功能

- 🎙️ **分段錄音**：多次錄音，每段獨立轉文字
- 🧠 **Whisper STT**：本地部署 Faster Whisper（small 模型）
- 🤖 **AI 彙整**：透過 OpenRouter API（預設 gpt-4o-mini，可透過 OPENROUTER_MODEL 環境變數自訂）將口語整理為正式報告
- 📄 **Word 下載**：支援模板（`{{content}}`、`{{date}}`、`{{report_type}}`）
- 📧 **Email 發送**：透過 Maton Gmail API 寄送報告

## 技術架構

```
用戶瀏覽器 → Zeabur → FastAPI
  → /segment    (Whisper 轉換)
  → /generate   (OpenRouter API)
  → /download   (Word 下載)
  → /email      (Maton Gmail API)
```

## 環境變數（Zeabur 設定）

| 變數 | 說明 |
|---|---|
| `OPENROUTER_API_KEY` | OpenRouter API Key（用於 AI 彙整，相容 OpenAI 格式） |
| `OPENROUTER_MODEL` | 自訂 AI 模型（預設 `openai/gpt-4o-mini`） |
| `ACCESS_PASSWORD` | 登入密碼（預設 `ABC1234`，請在 Zeabur 環境變數中設定） |
| `MATON_API_KEY` | Maton Gmail API Key（用於 Email 發送） |
| `HF_TOKEN` | HuggingFace Token（選用，加速 Whisper 模型下載） |

## 本地開發

```bash
pip install -r requirements.txt

export OPENROUTER_API_KEY="your-openrouter-api-key"
export MATON_API_KEY="your-maton-api-key"

uvicorn main:app --reload --port 8080
```

## Zeabur 部署

1. Fork 此倉庫到 GitHub
2. 在 Zeabur 連接倉庫
3. 設定環境變數（見上表）
4. **自動部署**後服務即上線

## API 端點

| 方法 | 路徑 | 說明 |
|---|---|---|
| GET | `/login` | 登入頁面 |
| POST | `/api/auth/login` | 登入驗證 |
| GET | `/dashboard` | 主控制台 |
| POST | `/api/sessions/{id}/segment` | 上傳錄音段落 |
| POST | `/api/sessions/{id}/clear` | 清空所有段落 |
| POST | `/api/sessions/{id}/generate` | 產生報告 |
| GET | `/api/sessions/{id}/download` | 下載報告 |
| POST | `/api/sessions/{id}/email` | Email 發送 |
| POST | `/api/sessions/{id}/template` | 上傳 Word 模板（唯一一份，覆蓋舊的） |
