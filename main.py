"""
MCH Assistant Web Service - 門諾醫院AI語音助理
語音轉文字並生成專業報告（支援分段錄音、Word模板）
"""

import os
import io
import uuid
import logging
import base64
import tempfile
import json
from datetime import datetime
from typing import Optional
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor

from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
from docx import Document

# =============================================================================
# Configuration
# =============================================================================

ACCESS_PASSWORD = "ABC1234"
UPLOAD_DIR = Path("./uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# Maton Gmail API (for sending emails)
MATON_API_KEY = os.getenv("MATON_API_KEY", "v2.zlcGqf1ftNATtrNkwmXFa8snundIGVsq_5-fzjBn9BArZalW1bk5IiPZgK9TSyu5ADpZ4hM08OlBHHdPTstn5f4xdUoA9GikGmBmPIf2GGZcb5Nsa5mDOMgR")
GMAIL_GATEWAY = "https://gateway.maton.ai/google-mail/gmail/v1/users/me/messages/send"

# Session storage
# {session_id: {"report_type": str, "template_path": str, "segments": [], "created_at": datetime}}
sessions = {}
executor = ThreadPoolExecutor(max_workers=2)

# =============================================================================
# Pydantic Models
# =============================================================================

class AudioSegmentRequest(BaseModel):
    audio_data: str  # Base64 encoded
    format: str = "webm"

class EmailRequest(BaseModel):
    to_email: str
    subject: str = "MCH Assistant 報告"
    body: str = ""

# =============================================================================
# Lifespan
# =============================================================================

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
    version="2.0.0",
    lifespan=lifespan
)

BASE_DIR = Path(__file__).parent
app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")

# =============================================================================
# Helper Functions
# =============================================================================

def get_session_id(request: Request) -> Optional[str]:
    return request.cookies.get("session_id")

def validate_session(request: Request) -> bool:
    session_id = get_session_id(request)
    return session_id and session_id in sessions

def get_or_create_session() -> str:
    """Get existing session or create new one"""
    session_id = str(uuid.uuid4())[:8]
    sessions[session_id] = {
        "report_type": "general",
        "template_path": None,
        "template_name": None,
        "segments": [],
        "generated_docx": None,
        "created_at": datetime.now()
    }
    return session_id

def transcribe_audio(audio_bytes: bytes) -> str:
    """Transcribe audio using Whisper (local)"""
    try:
        import whisper
        model = whisper.load_model("base")
        with tempfile.NamedTemporaryFile(suffix=".webm", delete=False) as f:
            f.write(audio_bytes)
            temp_path = f.name
        try:
            result = model.transcribe(temp_path, language="zh")
            return result["text"].strip()
        finally:
            os.unlink(temp_path)
    except Exception as e:
        logging.error(f"Whisper transcription error: {e}")
        return f"[轉換失敗: {str(e)}]"

def fill_template(template_path: str, segments: list, report_type: str) -> bytes:
    """Fill Word template with transcribed text"""
    doc = Document(template_path)
    
    # Build content from all segments
    full_text = "\n\n".join([
        f"【段落{i+1}】\n{seg['transcription']}" 
        for i, seg in enumerate(segments)
    ])
    
    report_date = datetime.now().strftime('%Y年%m月%d日')
    report_name = get_report_type_name(report_type)
    
    # Replace placeholders in paragraphs
    for para in doc.paragraphs:
        if "{{content}}" in para.text:
            para.text = para.text.replace("{{content}}", full_text)
        if "{{date}}" in para.text:
            para.text = para.text.replace("{{date}}", report_date)
        if "{{report_type}}" in para.text:
            para.text = para.text.replace("{{report_type}}", report_name)
    
    # Also replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "{{content}}" in cell.text:
                    cell.text = cell.text.replace("{{content}}", full_text)
                if "{{date}}" in cell.text:
                    cell.text = cell.text.replace("{{date}}", report_date)
                if "{{report_type}}" in cell.text:
                    cell.text = cell.text.replace("{{report_type}}", report_name)
    
    # Save to bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def get_report_type_name(report_type: str) -> str:
    names = {
        "general": "一般報告",
        "medical": "醫療報告", 
        "meeting": "會議記錄",
        "swallow": "吞嚥評估",
        "ent": "耳鼻喉科報告"
    }
    return names.get(report_type, report_type)

# =============================================================================
# Pages (HTML)
# =============================================================================

@app.get("/", response_class=HTMLResponse)
async def root(request: Request):
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    return RedirectResponse(url="/dashboard", status_code=302)

@app.get("/login", response_class=HTMLResponse)
async def login_page(request: Request):
    if validate_session(request):
        return RedirectResponse(url="/dashboard", status_code=302)
    
    # Inline HTML to avoid Jinja2 template issues
    html = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MCH Assistant - 登入</title>
    <link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
    <div class="login-container">
        <h1>🏥 MCH Assistant</h1>
        <p class="subtitle">門諾醫院 AI 語音助理</p>
        <p id="error" class="error" style="display:none">密碼錯誤，請重新輸入</p>
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
</html>"""
    return HTMLResponse(content=html)

@app.post("/api/auth/login")
async def login(request_data: dict):
    password = request_data.get("password", "")
    if password != ACCESS_PASSWORD:
        raise HTTPException(status_code=401, detail="Invalid password")
    
    session_id = get_or_create_session()
    response = RedirectResponse(url="/dashboard", status_code=302)
    response.set_cookie(key="session_id", value=session_id, httponly=True, samesite="lax")
    return response

@app.get("/api/auth/logout")
async def logout():
    response = RedirectResponse(url="/login", status_code=302)
    response.delete_cookie("session_id")
    return response

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    
    session_id = get_session_id(request)
    
    # Inline dashboard HTML
    html = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MCH Assistant - 語音助理</title>
    <link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
    <div class="header">
        <div class="logo">🏥 MCH Assistant</div>
        <div>
            <span class="session-id">Session: """ + session_id + """</span>
            <a href="/api/auth/logout" class="logout">登出</a>
        </div>
    </div>
    
    <div class="container">
        <div class="nav">
            <a href="/dashboard">🏠 主頁</a>
            <a href="/new">➕ 新建報告</a>
        </div>
        
        <h1>🎙️ 語音錄製</h1>
        <p class="subtitle">分段錄音，完成後合併產生報告</p>
        
        <div class="card">
            <h3>📋 報告類型</h3>
            <select id="reportType" class="select-full">
                <option value="general">一般報告</option>
                <option value="medical">醫療報告</option>
                <option value="meeting">會議記錄</option>
                <option value="swallow">吞嚥評估</option>
                <option value="ent">耳鼻喉科報告</option>
            </select>
        </div>
        
        <div class="card">
            <h3>📄 Word 模板</h3>
            <p class="hint">上傳 Word 模板（.docx），使用 {{content}}、{{date}}、{{report_type}} 作為佔位符</p>
            <input type="file" id="templateFile" accept=".docx" class="file-input">
            <div id="templateStatus" class="template-status"></div>
        </div>
        
        <div class="card recorder-card">
            <h3>🎤 錄音區段</h3>
            <div class="recorder-box">
                <button class="btn-record" id="recordBtn">🎤</button>
                <div class="status" id="status">點擊麥克風開始錄音</div>
            </div>
            
            <div id="segments" class="segments-container"></div>
            
            <button class="btn btn-secondary" id="addSegmentBtn" disabled>+ 新增段落</button>
        </div>
        
        <div class="actions">
            <button class="btn btn-generate" id="generateBtn" disabled>📝 產生報告</button>
            <button class="btn btn-download" id="downloadBtn" disabled>📥 下載</button>
            <button class="btn btn-email" id="emailBtn" disabled>📧 Email</button>
        </div>
        
        <div id="result" class="result-box"></div>
    </div>
    
    <script>
    const SESSION_ID = \"""" + session_id + """\";
    let mediaRecorder;
    let audioChunks = [];
    let isRecording = false;
    let segments = [];
    let hasTemplate = false;
    
    const recordBtn = document.getElementById('recordBtn');
    const status = document.getElementById('status');
    const segmentsDiv = document.getElementById('segments');
    const addSegmentBtn = document.getElementById('addSegmentBtn');
    const generateBtn = document.getElementById('generateBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const emailBtn = document.getElementById('emailBtn');
    const templateFile = document.getElementById('templateFile');
    const templateStatus = document.getElementById('templateStatus');
    const reportType = document.getElementById('reportType');
    const result = document.getElementById('result');
    
    templateFile.onchange = async () => {
        if (templateFile.files.length > 0) {
            const file = templateFile.files[0];
            const formData = new FormData();
            formData.append('file', file);
            formData.append('session_id', SESSION_ID);
            
            try {
                const res = await fetch('/api/sessions/' + SESSION_ID + '/template', {
                    method: 'POST',
                    body: formData
                });
                const data = await res.json();
                if (data.success) {
                    hasTemplate = true;
                    templateStatus.innerHTML = '<span class="success">✓ 已上傳：' + file.name + '</span>';
                    checkGenerateReady();
                } else {
                    templateStatus.innerHTML = '<span class="error">上傳失敗：' + data.error + '</span>';
                }
            } catch (e) {
                templateStatus.innerHTML = '<span class="error">上傳失敗</span>';
            }
        }
    };
    
    recordBtn.onclick = async () => {
        if (!isRecording) {
            const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
            mediaRecorder = new MediaRecorder(stream);
            audioChunks = [];
            mediaRecorder.ondataavailable = e => audioChunks.push(e.data);
            mediaRecorder.onstop = async () => {
                const blob = new Blob(audioChunks, { type: 'audio/webm' });
                const reader = new FileReader();
                reader.readAsDataURL(blob);
                reader.onloadend = async () => {
                    const base64 = reader.result.split(',')[1];
                    try {
                        const res = await fetch('/api/sessions/' + SESSION_ID + '/segment', {
                            method: 'POST',
                            headers: { 'Content-Type': 'application/json' },
                            body: JSON.stringify({ audio_data: base64, format: 'webm' })
                        });
                        const data = await res.json();
                        if (data.success) {
                            segments.push({ id: data.segment_id, text: data.transcription });
                            renderSegments();
                            addSegmentBtn.disabled = false;
                            checkGenerateReady();
                        }
                    } catch (e) {
                        alert('轉換失敗');
                    }
                    stream.getTracks().forEach(t => t.stop());
                };
            };
            mediaRecorder.start();
            isRecording = true;
            recordBtn.textContent = '⏹️';
            status.textContent = '錄音中...';
        } else {
            mediaRecorder.stop();
            isRecording = false;
            recordBtn.textContent = '🎤';
            status.textContent = '錄音完成';
        }
    };
    
    function renderSegments() {
        segmentsDiv.innerHTML = segments.map((seg, i) => 
            '<div class="segment-item"><strong>段落' + (i+1) + ':</strong> ' + seg.text + '</div>'
        ).join('');
    }
    
    function checkGenerateReady() {
        const ready = segments.length > 0;
        generateBtn.disabled = !ready;
    }
    
    addSegmentBtn.onclick = () => {
        addSegmentBtn.disabled = true;
        status.textContent = '點擊麥克風開始錄音';
    };
    
    generateBtn.onclick = async () => {
        generateBtn.textContent = '處理中...';
        generateBtn.disabled = true;
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ report_type: reportType.value })
            });
            const data = await res.json();
            if (data.success) {
                result.innerHTML = '<div class="success">✓ 報告已產生！</div>';
                downloadBtn.disabled = false;
                emailBtn.disabled = false;
            } else {
                result.innerHTML = '<div class="error">產生失敗：' + data.error + '</div>';
            }
        } catch (e) {
            result.innerHTML = '<div class="error">錯誤：' + e.message + '</div>';
        }
        generateBtn.textContent = '📝 產生報告';
        generateBtn.disabled = false;
    };
    
    downloadBtn.onclick = () => {
        window.location.href = '/api/sessions/' + SESSION_ID + '/download';
    };
    
    emailBtn.onclick = async () => {
        const email = prompt('請輸入收件者 Email:');
        if (email) {
            emailBtn.textContent = '發送中...';
            emailBtn.disabled = true;
            try {
                const res = await fetch('/api/sessions/' + SESSION_ID + '/email', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ to_email: email })
                });
                const data = await res.json();
                if (data.success) {
                    alert('✓ 郵件已發送至 ' + email);
                } else {
                    alert('發送失敗：' + data.error);
                }
            } catch (e) {
                alert('錯誤：' + e.message);
            }
            emailBtn.textContent = '📧 Email';
            emailBtn.disabled = false;
        }
    };
    </script>
</body>
</html>"""
    return HTMLResponse(content=html)

@app.get("/new")
async def new_session(request: Request):
    """Create new session"""
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    
    # Create new session
    new_id = get_or_create_session()
    response = RedirectResponse(url="/dashboard", status_code=302)
    response.set_cookie(key="session_id", value=new_id, httponly=True, samesite="lax")
    return response

# =============================================================================
# API Endpoints
# =============================================================================

@app.post("/api/sessions/{session_id}/template")
async def upload_template(session_id: str, request: Request, file: UploadFile = File(...)):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if not file.filename.endswith('.docx'):
        return {"success": False, "error": "請上傳 .docx 檔案"}
    
    # Save template
    template_path = UPLOAD_DIR / f"{session_id}_template.docx"
    with open(template_path, "wb") as f:
        content = await file.read()
        f.write(content)
    
    sessions[session_id]["template_path"] = str(template_path)
    sessions[session_id]["template_name"] = file.filename
    
    return {"success": True, "filename": file.filename}

@app.post("/api/sessions/{session_id}/segment")
async def add_segment(session_id: str, req: AudioSegmentRequest, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Decode audio
    audio_bytes = base64.b64decode(req.audio_data)
    
    # Save audio
    segment_id = str(uuid.uuid4())[:8]
    audio_path = UPLOAD_DIR / f"{session_id}_{segment_id}.webm"
    with open(audio_path, "wb") as f:
        f.write(audio_bytes)
    
    # Transcribe in background thread
    def do_transcribe():
        return transcribe_audio(audio_bytes)
    
    try:
        import asyncio
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        future = loop.run_in_executor(executor, do_transcribe)
        text = loop.run_until_complete(asyncio.wait_for(future, timeout=120))
    except asyncio.TimeoutError:
        text = "[轉換失敗: 逾時]"
    except Exception as e:
        text = f"[轉換失敗: {str(e)}]"
    
    # Store segment
    sessions[session_id]["segments"].append({
        "id": segment_id,
        "transcription": text,
        "audio_path": str(audio_path)
    })
    
    return {
        "success": True,
        "segment_id": segment_id,
        "transcription": text
    }

@app.post("/api/sessions/{session_id}/generate")
async def generate_report(session_id: str, request: Request, req: dict):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    
    if not session["segments"]:
        return {"success": False, "error": "尚無錄音內容"}
    
    report_type = req.get("report_type", "general")
    session["report_type"] = report_type
    
    # If template exists, fill it
    if session["template_path"] and Path(session["template_path"]).exists():
        try:
            docx_bytes = fill_template(
                session["template_path"],
                session["segments"],
                report_type
            )
            session["generated_docx"] = docx_bytes
            return {"success": True, "has_template": True}
        except Exception as e:
            return {"success": False, "error": str(e)}
    else:
        # No template - return text only
        full_text = "\n\n".join([
            f"【段落{i+1}】\n{seg['transcription']}"
            for i, seg in enumerate(session["segments"])
        ])
        return {"success": True, "has_template": False, "text": full_text}

@app.get("/api/sessions/{session_id}/download")
async def download_report(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    
    if session.get("generated_docx"):
        return Response(
            content=session["generated_docx"],
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={
                "Content-Disposition": f"attachment; filename=MCH_報告_{datetime.now().strftime('%Y%m%d')}.docx"
            }
        )
    elif session["segments"]:
        # Return text if no docx
        full_text = "\n\n".join([
            f"【段落{i+1}】\n{seg['transcription']}"
            for i, seg in enumerate(session["segments"])
        ])
        return Response(
            content=full_text,
            media_type="text/plain; charset=utf-8",
            headers={
                "Content-Disposition": f"attachment; filename=MCH_報告_{datetime.now().strftime('%Y%m%d')}.txt"
            }
        )
    else:
        raise HTTPException(status_code=404, detail="No report to download")

@app.post("/api/sessions/{session_id}/email")
async def email_report(session_id: str, request: Request, req: EmailRequest):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    
    if not session.get("generated_docx") and not session["segments"]:
        return {"success": False, "error": "無報告可發送"}
    
    try:
        import urllib.request
        
        # Build email content
        report_date = datetime.now().strftime('%Y年%m月%d日')
        report_type_name = get_report_type_name(session.get("report_type", "general"))
        
        # Create email
        msg = MIMEMultipart()
        msg['to'] = req.to_email
        msg['subject'] = f"MCH 報告 - {report_date}"
        
        # Email body
        body = f"""
您好，

這是來自 MCH Assistant 的語音報告。

報告日期：{report_date}
報告類型：{report_type_name}

--- 錄音內容 ---
{chr(10).join([f"【段落{i+1}】{seg['transcription']}" for i, seg in enumerate(session['segments'])])}

---
此郵件由 MCH Assistant 自動發送
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # If we have a generated docx, attach it
        if session.get("generated_docx"):
            from email.mime.base import MIMEBase
            from email import encoders
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(session["generated_docx"])
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename=MCH_報告_{report_date.replace("年", "").replace("月", "").replace("日", "")}.docx')
            msg.attach(part)
        
        # Encode to base64url
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode().rstrip('=')
        
        # Send via Maton Gmail API
        data = json.dumps({"raw": raw}).encode()
        gmail_req = urllib.request.Request(GMAIL_GATEWAY, data=data, method='POST')
        gmail_req.add_header('Authorization', f'Bearer {MATON_API_KEY}')
        gmail_req.add_header('Content-Type', 'application/json')
        
        with urllib.request.urlopen(gmail_req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
            return {
                "success": True,
                "message": f"報告已發送至 {req.to_email}",
                "email_id": result.get("id")
            }
    except urllib.error.HTTPError as e:
        error_body = e.read().decode()
        return {"success": False, "error": f"郵件發送失敗：{error_body}"}
    except Exception as e:
        logging.error(f"Email error: {e}")
        return {"success": False, "error": str(e)}

@app.get("/health")
async def health():
    return {"status": "healthy", "service": "MCH Assistant", "version": "2.0.0"}

@app.get("/api/status")
async def api_status():
    return {
        "service": "MCH Assistant",
        "version": "2.0.0",
        "status": "running",
        "sessions": len(sessions)
    }

# =============================================================================
# Run
# =============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port)