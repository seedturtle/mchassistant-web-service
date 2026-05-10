"""
MCH Assistant Web Service - 門諾醫院AI語音助理
語音轉文字並生成專業報告
"""

import os
import logging
from contextlib import asynccontextmanager
from datetime import datetime, timedelta
from typing import Optional
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request, Form
from fastapi.responses import HTMLResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import uuid

# =============================================================================
# Configuration
# =============================================================================

# 登入密碼（暫時設定）
ACCESS_PASSWORD = "ABC1234"

# Session storage (simple in-memory for demo)
sessions = {}  # {session_id: {"logged_in": True, "user": "admin"}}

# =============================================================================
# Pydantic Models
# =============================================================================

class AudioTranscribeRequest(BaseModel):
    audio_data: str  # Base64 encoded audio
    format: str = "webm"
    report_type: str = "general"

# =============================================================================
# Hermes Integration Service
# =============================================================================

class HermesService:
    """與 Hermes AI Agent 串接的服務"""
    
    def __init__(self):
        self.api_key = os.getenv("HERMES_API_KEY", "")
        self.model = os.getenv("HERMES_MODEL", "minimax/minimax-m2.7")
    
    async def generate_report(self, transcription: str, report_type: str = "general") -> str:
        """
        將文字轉換為專業報告
        """
        # 這裡是示範報告格式，實際使用時會串接 Hermes AI
        report_date = datetime.now().strftime('%Y年%m月%d日 %H:%M')
        
        if report_type == "medical":
            return f"""# 醫療報告
生成時間：{report_date}

## 語音轉文字內容
{transcription}

## 診斷摘要
本報告由 AI 語音助理根據語音輸入自動生成。
如有醫療需求，請諮詢專業醫護人員。

---
MCH Assistant 門諾醫院AI助理
"""
        elif report_type == "meeting":
            return f"""# 會議記錄
生成時間：{report_date}

## 語音轉文字內容
{transcription}

## 重點摘要
- 待補充

---
MCH Assistant 門諾醫院AI助理
"""
        else:
            return f"""# 報告
生成時間：{report_date}

## 語音轉文字內容
{transcription}

## 備註
本報告由 MCH Assistant 語音助理生成。

---
MCH Assistant 門諾醫院AI助理
"""

hermes_service = HermesService()

# =============================================================================
# Lifespan
# =============================================================================

@asynccontextmanager
async def lifespan(app: FastAPI):
    logging.info("MCH Assistant Web Service 啟動中...")
    yield
    logging.info("MCH Assistant Web Service 關閉中...")

# =============================================================================
# FastAPI Application
# =============================================================================

app = FastAPI(
    title="MCH Assistant 語音助理",
    description="門諾醫院AI語音助理 - 語音轉文字生成報告",
    version="1.0.0",
    lifespan=lifespan
)

# Static files
BASE_DIR = Path(__file__).parent
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")

# =============================================================================
# Helper Functions
# =============================================================================

def get_session_id(request: Request) -> Optional[str]:
    return request.cookies.get("session_id")

def validate_session(request: Request) -> bool:
    session_id = get_session_id(request)
    if session_id and session_id in sessions:
        return True
    return False

def create_session():
    session_id = str(uuid.uuid4())
    sessions[session_id] = {"logged_in": True}
    return session_id

# =============================================================================
# Pages (HTML)
# =============================================================================

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    """首頁 - 檢查登入狀態"""
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    return RedirectResponse(url="/dashboard", status_code=302)

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    """登入頁面"""
    if validate_session(request):
        return RedirectResponse(url="/dashboard", status_code=302)
    
    return """
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>MCH Assistant - 登入</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { 
                font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; 
                background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
                min-height: 100vh;
                display: flex;
                align-items: center;
                justify-content: center;
            }
            .login-container {
                background: rgba(255,255,255,0.1);
                backdrop-filter: blur(10px);
                border-radius: 20px;
                padding: 40px;
                width: 100%;
                max-width: 400px;
                box-shadow: 0 8px 32px rgba(0,0,0,0.3);
            }
            h1 { 
                color: #e94560; 
                text-align: center;
                margin-bottom: 10px;
                font-size: 28px;
            }
            .subtitle {
                color: #aaa;
                text-align: center;
                margin-bottom: 30px;
                font-size: 14px;
            }
            .input-group { margin-bottom: 20px; }
            label { 
                display: block; 
                color: #fff; 
                margin-bottom: 8px;
                font-size: 14px;
            }
            input { 
                width: 100%; 
                padding: 15px; 
                background: rgba(255,255,255,0.1);
                border: 2px solid transparent;
                border-radius: 10px;
                color: #fff;
                font-size: 16px;
                transition: all 0.3s;
            }
            input:focus {
                outline: none;
                border-color: #e94560;
                background: rgba(255,255,255,0.15);
            }
            .btn {
                width: 100%;
                padding: 15px;
                background: #e94560;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 18px;
                cursor: pointer;
                transition: all 0.3s;
                font-family: inherit;
            }
            .btn:hover { background: #d63b53; transform: translateY(-2px); }
            .error { 
                color: #ff6b6b; 
                text-align: center; 
                margin-bottom: 15px;
                display: none;
            }
            .info {
                color: #888;
                text-align: center;
                margin-top: 20px;
                font-size: 12px;
            }
        </style>
    </head>
    <body>
        <div class="login-container">
            <h1>🏥 MCH Assistant</h1>
            <p class="subtitle">門諾醫院 AI 語音助理</p>
            <p id="error" class="error">密碼錯誤，請重新輸入</p>
            <form id="loginForm">
                <div class="input-group">
                    <label for="password">🔐 請輸入存取密碼</label>
                    <input type="password" id="password" name="password" placeholder="輸入密碼" required>
                </div>
                <button type="submit" class="btn">登入</button>
            </form>
            <p class="info">僅限醫院內部人員使用</p>
        </div>
        <script>
            document.getElementById('loginForm').onsubmit = async (e) => {
                e.preventDefault();
                const password = document.getElementById('password').value;
                const response = await fetch('/api/auth/login', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({password: password})
                });
                if (response.ok) {
                    window.location.href = '/dashboard';
                } else {
                    document.getElementById('error').style.display = 'block';
                    document.getElementById('password').value = '';
                }
            };
        </script>
    </body>
    </html>
    """

@app.post("/api/auth/login")
async def login(request_data: dict):
    """驗證登入"""
    password = request_data.get("password", "")
    if password == ACCESS_PASSWORD:
        session_id = create_session()
        response = RedirectResponse(url="/dashboard", status_code=302)
        response.set_cookie(key="session_id", value=session_id, httponly=True, samesite="lax")
        return response
    raise HTTPException(status_code=401, detail="Invalid password")

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    """儀表板首頁"""
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    
    return """
    <!DOCTYPE html>
    <html lang="zh-TW">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>MCH Assistant - 語音助理</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; }
            body { 
                font-family: 'PingFang TC', 'Microsoft JhengHei', sans-serif; 
                background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
                min-height: 100vh;
                color: #fff;
            }
            .header {
                background: rgba(255,255,255,0.05);
                padding: 20px 40px;
                display: flex;
                justify-content: space-between;
                align-items: center;
                border-bottom: 1px solid rgba(255,255,255,0.1);
            }
            .logo { font-size: 24px; color: #e94560; }
            .logout { color: #888; text-decoration: none; font-size: 14px; }
            .container { max-width: 900px; margin: 40px auto; padding: 0 20px; }
            h1 { color: #e94560; margin-bottom: 10px; }
            .subtitle { color: #888; margin-bottom: 30px; }
            .card {
                background: rgba(255,255,255,0.05);
                border-radius: 15px;
                padding: 30px;
                margin-bottom: 20px;
            }
            .card h2 { color: #e94560; margin-bottom: 15px; }
            .card p { color: #aaa; line-height: 1.6; }
            .btn-start {
                display: inline-block;
                padding: 15px 40px;
                background: #e94560;
                color: white;
                border-radius: 10px;
                text-decoration: none;
                font-size: 18px;
                margin-top: 20px;
                transition: all 0.3s;
            }
            .btn-start:hover { background: #d63b53; transform: translateY(-2px); }
            .recorder-box {
                background: rgba(233,69,96,0.1);
                border: 2px dashed #e94560;
                border-radius: 15px;
                padding: 40px;
                text-align: center;
                margin: 20px 0;
            }
            .record-btn {
                font-size: 64px;
                cursor: pointer;
                background: none;
                border: none;
                transition: transform 0.3s;
            }
            .record-btn:hover { transform: scale(1.1); }
            .status { margin: 20px 0; font-size: 18px; color: #aaa; }
            .waveform {
                height: 60px;
                background: rgba(255,255,255,0.05);
                border-radius: 10px;
                margin: 20px 0;
                display: flex;
                align-items: center;
                justify-content: center;
                color: #666;
            }
            .report-type {
                margin: 20px 0;
                display: flex;
                gap: 10px;
                flex-wrap: wrap;
            }
            .report-type label { color: #fff; margin-right: 10px; }
            .report-type select {
                padding: 10px 20px;
                background: rgba(255,255,255,0.1);
                color: #fff;
                border: none;
                border-radius: 5px;
                font-size: 16px;
            }
            .btn-generate {
                padding: 15px 40px;
                background: #e94560;
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 18px;
                cursor: pointer;
                margin-top: 20px;
                display: none;
            }
            .btn-generate:hover { background: #d63b53; }
            #result {
                background: rgba(255,255,255,0.05);
                border-radius: 10px;
                padding: 20px;
                margin-top: 20px;
                display: none;
            }
            #result pre { white-space: pre-wrap; line-height: 1.6; }
            .nav { margin-bottom: 30px; }
            .nav a {
                color: #e94560;
                text-decoration: none;
                margin-right: 20px;
                font-size: 14px;
            }
        </style>
    </head>
    <body>
        <div class="header">
            <div class="logo">🏥 MCH Assistant</div>
            <a href="/api/auth/logout" class="logout">登出</a>
        </div>
        <div class="container">
            <div class="nav">
                <a href="/dashboard">🏠 首頁</a>
                <a href="/record">🎙️ 錄音</a>
            </div>
            <h1>🎙️ 語音錄製</h1>
            <p class="subtitle">錄製語音並轉換為專業報告</p>
            <div class="recorder-box">
                <button class="record-btn" id="recordBtn">🎤</button>
                <div class="status" id="status">點擊麥克風開始錄音</div>
                <div class="waveform" id="waveform">錄音波形區域</div>
                <div class="report-type">
                    <label for="reportType">報告類型：</label>
                    <select id="reportType">
                        <option value="general">一般報告</option>
                        <option value="medical">醫療報告</option>
                        <option value="meeting">會議記錄</option>
                    </select>
                </div>
                <button class="btn-generate" id="submitBtn">生成報告</button>
            </div>
            <div id="result"></div>
        </div>
        <script>
            let mediaRecorder;
            let audioChunks = [];
            let isRecording = false;
            let audioBlob;
            const recordBtn = document.getElementById('recordBtn');
            const status = document.getElementById('status');
            const submitBtn = document.getElementById('submitBtn');
            const result = document.getElementById('result');

            recordBtn.onclick = async () => {
                if (!isRecording) {
                    const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
                    mediaRecorder = new MediaRecorder(stream);
                    audioChunks = [];
                    mediaRecorder.ondataavailable = e => audioChunks.push(e.data);
                    mediaRecorder.onstop = () => {
                        audioBlob = new Blob(audioChunks, { type: 'audio/webm' });
                        submitBtn.style.display = 'inline-block';
                    };
                    mediaRecorder.start();
                    isRecording = true;
                    recordBtn.textContent = '⏹️';
                    status.textContent = '錄音中... 再次點擊停止';
                    document.getElementById('waveform').textContent = '▓▓▓▓▓▓▓▓▓▓';
                } else {
                    mediaRecorder.stop();
                    isRecording = false;
                    recordBtn.textContent = '🎤';
                    status.textContent = '錄音完成';
                    document.getElementById('waveform').textContent = '錄音已完成';
                }
            };

            submitBtn.onclick = async () => {
                if (!audioBlob) {
                    alert('請先錄音');
                    return;
                }
                submitBtn.textContent = '處理中...';
                submitBtn.disabled = true;
                
                const reader = new FileReader();
                reader.readAsDataURL(audioBlob);
                reader.onloadend = async () => {
                    const base64 = reader.result.split(',')[1];
                    const reportType = document.getElementById('reportType').value;
                    
                    try {
                        const response = await fetch('/api/audio/transcribe', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ 
                                audio_data: base64, 
                                format: 'webm',
                                report_type: reportType 
                            })
                        });
                        const data = await response.json();
                        result.style.display = 'block';
                        result.innerHTML = '<pre>' + data.report + '</pre>';
                    } catch (err) {
                        result.style.display = 'block';
                        result.innerHTML = '錯誤: ' + err.message;
                    }
                    submitBtn.textContent = '生成報告';
                    submitBtn.disabled = false;
                };
            };
        </script>
    </body>
    </html>
    """

@app.get("/api/auth/logout")
async def logout():
    """登出"""
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie("session_id")
    return response

# =============================================================================
# API Endpoints
# =============================================================================

@app.post("/api/audio/transcribe")
async def transcribe_audio(req: AudioTranscribeRequest, request: Request):
    """接收音訊資料，轉換為文字並生成報告"""
    if not validate_session(request):
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    # 這裡是示範用的即興文字，實際會串接語音轉文字服務
    transcription = "這是語音轉換的文字內容。在實際部署時，這裡會串接專業的語音辨識服務（如 Whisper、Azure Speech 等）來將音訊轉換為文字。\n\n虛擬文字內容：\n- 病人主訴症狀\n- 醫療處置建議\n- 用藥注意事項"
    
    report = await hermes_service.generate_report(transcription, req.report_type)
    
    return {"transcription": transcription, "report": report}

@app.get("/health")
async def health():
    return {"status": "healthy", "service": "MCH Assistant"}

@app.get("/api/status")
async def api_status():
    return {
        "service": "MCH Assistant",
        "version": "1.0.0",
        "status": "running"
    }

# =============================================================================
# Run with uvicorn
# =============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port)