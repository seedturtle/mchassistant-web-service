# mchassistant-web-service

## 專案目的
語音轉文字，再將逐段錄音/上傳音檔統整成正式 Word 報告的 FastAPI web service。部署目標目前在 Zeabur。

## 目前架構
- 入口：`main.py`
- 前端：目前主要由 `main.py` 的 `/dashboard` 內嵌 HTML/JS 提供；`templates/dashboard.html` 與 `static/js/app.js` 看起來像舊版/部分未使用的模板資源。
- STT：`faster-whisper` small model，本機 CPU/int8，函式 `transcribe_audio()`。
- 報告統整：MiniMax API (`minimax-m2.7`)，函式 `summarize_with_hermes()`。
- Word 輸出：`python-docx`，模板 placeholder 支援 `{{content}}`, `{{date}}`, `{{report_type}}` 與自訂欄位。
- Email/Google Drive：透過 Maton gateway API；`MATON_API_KEY` 必須在 Zeabur 環境變數。
- 部署：Dockerfile + `zeabur.json`，port 8080。

## Zeabur 服務
使用者提供的 Zeabur 服務頁：`https://zeabur.com/projects/6a004353e6a21fff4d961e3d/services/6a004354e6a21fff4d961e3e?envID=6a004354e5ed304c1d84596f`
目前瀏覽時導向 Zeabur Projects，需要登入或前端權限；可先以 GitHub repo 內容分析。

## 重要環境變數
- `MINIMAX_API_KEY`：AI 報告統整
- `MATON_API_KEY`：Gmail/Google Drive gateway
- `HF_TOKEN`：可選，Whisper 模型下載

## 注意事項
- 不要把密碼或 API key 寫入 repo。`ACCESS_PASSWORD` 目前硬編在 `main.py`（`ABC1234`），未來應改成環境變數。
- `sessions` 是記憶體資料；Zeabur 重啟後 session 與已上傳暫存音檔狀態會消失，但模板/報告類型會嘗試同步到 Drive。
- 若要修改 dashboard UI，優先確認目前實際使用的是 `main.py` 內嵌 HTML，而非 `templates/dashboard.html`。
- 做部署前先在本機測試：`uvicorn main:app --host 0.0.0.0 --port 8080`。

## 待研究/改善方向
1. 將 `ACCESS_PASSWORD` 改為 `ACCESS_PASSWORD` 環境變數。
2. 確認 `templates/dashboard.html` 與 `static/js/app.js` 是否可刪除或重整，避免雙版本混淆。
3. 評估 STT 是否改用 Gemini/Google Speech/OpenAI Whisper API，以減少 Zeabur CPU 與模型下載負擔。
4. 補上健康檢查端點與部署日誌診斷流程。
5. 為 Zeabur persistent storage / Drive sync 補強，避免 sessions 與 uploads 因重啟遺失。
6. 增加更明確的背景處理錯誤訊息與使用者可見 retry。 
