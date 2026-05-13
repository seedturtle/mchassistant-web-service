"""
MCH Assistant Web Service - 門諾醫院AI語音助理
語音轉文字並生成專業報告（支援分段錄音、Word模板）
"""

import os
import io
import re
import uuid
import logging
import base64
import tempfile
import json
from datetime import datetime
from typing import Optional, List
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, TimeoutError

from fastapi import FastAPI, HTTPException, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse
from pydantic import BaseModel
from docx import Document
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# =============================================================================
# Configuration
# =============================================================================

ACCESS_PASSWORD = "ABC1234"
UPLOAD_DIR = Path("./uploads")
UPLOAD_DIR.mkdir(exist_ok=True)

# Maton Gmail API (for sending emails)
MATON_API_KEY=os.getenv("MATON_API_KEY", "")
GMAIL_GATEWAY = "https://gateway.maton.ai/google-mail/gmail/v1/users/me/messages/send"

# HuggingFace token (for Faster Whisper model download)
HF_TOKEN=os.getenv("HF_TOKEN", "")

# MiniMax API (for report summarization)
MINIMAX_API_KEY=os.getenv("MINIMAX_API_KEY", "")
MINIMAX_API_GATEWAY="https://api.minimax.io/v1/text/chatcompletion_v2"

# Session storage
# {session_id: {"report_type": str, "segments": [], "created_at": datetime}}
sessions = {}
executor = ThreadPoolExecutor(max_workers=2)

# Persistent template storage
TEMPLATES_DIR = Path("./templates")
TEMPLATES_DIR.mkdir(exist_ok=True)

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
        "segments": [],
        "generated_docx": None,
        "created_at": datetime.now(),
        "upload_files": {}  # {filename: {status, transcription, segment_id, error}}
    }
    return session_id

def transcribe_audio(audio_bytes: bytes, file_ext: str = ".webm") -> str:
    """Transcribe audio using Faster Whisper (local)
    
    Args:
        audio_bytes: Raw audio data
        file_ext: File extension (e.g. .wav, .mp3, .m4a) for temp file
    """
    try:
        from faster_whisper import WhisperModel
        import os as os_module
        # Set HF_TOKEN for faster model download
        if HF_TOKEN:
            os_module.environ["HF_TOKEN"] = HF_TOKEN
        model = WhisperModel("small", device="cpu", compute_type="int8")
        with tempfile.NamedTemporaryFile(suffix=file_ext, delete=False) as f:
            f.write(audio_bytes)
            temp_path = f.name
        try:
            segments, info = model.transcribe(temp_path, language="zh")
            text = "".join([s.text for s in segments]).strip()
            return text if text else "[無辨識結果]"
        finally:
            os.unlink(temp_path)
    except Exception as e:
        logging.error(f"Faster Whisper transcription error: {e}")
        return f"[轉換失敗: {str(e)}]"

def summarize_with_hermes(transcribed_text: str, report_type: str, placeholders: list = None) -> any:
    """Use MiniMax API to summarize and organize transcribed text into a professional report.
    
    If placeholders is provided, returns a dict {field_name: content}.
    Otherwise returns a plain summarized string.
    """
    logging.info(f"[summarize_with_hermes] Called with report_type={report_type}, text_length={len(transcribed_text)}, placeholders={placeholders}")
    if not MINIMAX_API_KEY:
        logging.warning("[summarize_with_hermes] No API key, returning original text")
        if placeholders:
            return {p: transcribed_text for p in placeholders}
        return transcribed_text
    
    try:
        import urllib.request
        import urllib.error
        
        report_type_names = {
            "general": "一般報告",
            "medical": "醫療報告",
            "meeting": "會議記錄",
            "swallow": "吞嚥評估",
            "ent": "耳鼻喉科報告"
        }
        type_name = report_type_names.get(report_type, "報告")
        
        if placeholders:
            # === 結構化模式：Template has specific fields ===
            fields_str = "、".join(placeholders)
            user_prompt = f"""報告類型：{type_name}

模板欄位：{fields_str}

錄音內容：
{transcribed_text}

請根據錄音內容，為以上每個模板欄位產生合適的文字內容。
以 JSON 格式回傳，不要有其他文字。範例：
{{"{placeholders[0]}": "填寫內容", "{placeholders[-1]}": "填寫內容"}}"""
            system_prompt = "你是專業的醫療報告整理助理。你只回傳純 JSON，不附加任何說明文字。"
        else:
            # === 自由模式：No specific fields, just summarize ===
            system_prompt = "你是專業的醫療報告整理助理。請將下面的口語錄音整理成正式格式，直接輸出結果，不需要標記【段落】。"
            user_prompt = f"報告類型：{type_name}\n\n錄音內容：\n{transcribed_text}"
        
        payload = json.dumps({
            "model": "minimax-m2.7",
            "messages": [
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            "temperature": 0.3,
            "max_tokens": 4000
        }).encode()
        
        req = urllib.request.Request(MINIMAX_API_GATEWAY, data=payload, method='POST')
        req.add_header('Authorization', f'Bearer {MINIMAX_API_KEY}')
        req.add_header('Content-Type', 'application/json')
        
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode())
            logging.warning(f"[summarize_with_hermes] MiniMax raw response: {str(result)[:500]}")
            if 'choices' in result and len(result['choices']) > 0:
                content = result['choices'][0]['message']['content'].strip()
                logging.warning(f"[summarize_with_hermes] Extracted content: {content[:300]}")
                
                if placeholders:
                    # Try to parse as JSON
                    try:
                        # Strip markdown code fences if present
                        clean = content.strip()
                        if clean.startswith("```"):
                            clean = clean.split("\n", 1)[1] if "\n" in clean else clean
                            clean = clean.rsplit("```", 1)[0] if "```" in clean else clean
                            clean = clean.strip()
                        fields_result = json.loads(clean)
                        # Ensure all placeholders have values
                        for p in placeholders:
                            if p not in fields_result:
                                fields_result[p] = f"[待補充：{p}]"
                        return fields_result
                    except (json.JSONDecodeError, Exception) as e:
                        logging.error(f"[summarize_with_hermes] JSON parse error: {e}, content={content[:200]}")
                        # Fallback: put everything in first placeholder
                        return {p: content if i == 0 else f"[待補充：{p}]" for i, p in enumerate(placeholders)}
                else:
                    return content
            
            logging.warning("[summarize_with_hermes] No choices in response")
            if placeholders:
                return {p: f"[待補充：{p}]" for p in placeholders}
            return transcribed_text
    except Exception as e:
        logging.error(f"MiniMax summarization error: {e}")
        if placeholders:
            return {p: transcribed_text for p in placeholders}
        return transcribed_text

def fill_template(template_path: str, segments: list, report_type: str, summarized_text: str = None, fields_dict: dict = None) -> bytes:
    """Fill Word template with transcribed text and structured field data"""
    doc = Document(template_path)
    
    # Build full text from segments
    if summarized_text:
        full_text = summarized_text
    else:
        full_text = "\n\n".join([
            f"【段落{i+1}】\n{seg['transcription']}" 
            for i, seg in enumerate(segments)
        ])
    
    report_date = datetime.now().strftime('%Y年%m月%d日')
    report_name = get_report_type_name(report_type)
    
    def replace_paragraph_text(para):
        """Replace all placeholder text in a paragraph, handling runs properly"""
        # Get all paragraph text by joining runs
        full_para_text = ""
        for run in para.runs:
            full_para_text += run.text
        
        new_text = full_para_text
        if "{{content}}" in new_text:
            new_text = new_text.replace("{{content}}", full_text)
            nonlocal content_replaced
            content_replaced = True
        if "{{date}}" in new_text:
            new_text = new_text.replace("{{date}}", report_date)
        if "{{report_type}}" in new_text:
            new_text = new_text.replace("{{report_type}}", report_name)
        if fields_dict:
            for key, value in fields_dict.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in new_text:
                    new_text = new_text.replace(placeholder, str(value))
        
        # Only modify if something changed
        if new_text != full_para_text:
            # Keep first run's formatting, clear all text, set new text in first run
            for i, run in enumerate(para.runs):
                if i == 0:
                    run.text = new_text
                else:
                    run.text = ""
    
    # Replace in paragraphs
    content_replaced = False
    for para in doc.paragraphs:
        replace_paragraph_text(para)
    
    # Replace in table cells
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_paragraph_text(para)
    
    # If no {{content}} placeholder was found, append content at end
    if not content_replaced and full_text:
        doc.add_paragraph("")
        doc.add_paragraph(full_text)
    
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

def extract_placeholders(doc_path: str) -> list:
    """Scan Word template and extract all {{placeholder}} patterns (except built-in ones)"""
    doc = Document(doc_path)
    placeholders = set()
    for para in doc.paragraphs:
        found = re.findall(r'\{\{(.*?)\}\}', para.text)
        placeholders.update(found)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                found = re.findall(r'\{\{(.*?)\}\}', cell.text)
                placeholders.update(found)
    # Built-in fields we auto-fill — exclude from dynamic fields
    auto_fill = {"date", "report_type", "content"}
    return [p for p in placeholders if p not in auto_fill]

def _process_uploaded_file(session_id: str, filepath: str, filename: str, ext: str):
    """Background task: transcribe audio and store result in session"""
    session = sessions.get(session_id)
    if not session:
        try:
            os.unlink(filepath)
        except:
            pass
        return
    
    session["upload_files"][filename]["status"] = "processing"
    
    try:
        with open(filepath, "rb") as f:
            audio_bytes = f.read()
        text = transcribe_audio(audio_bytes, file_ext=ext)
        
        if text.startswith("[轉換失敗"):
            session["upload_files"][filename]["status"] = "error"
            session["upload_files"][filename]["error"] = text
        else:
            segment_id = str(uuid.uuid4())[:8]
            session["segments"].append({
                "id": segment_id,
                "transcription": text,
                "audio_path": filename
            })
            session["upload_files"][filename]["status"] = "done"
            session["upload_files"][filename]["segment_id"] = segment_id
            session["upload_files"][filename]["transcription"] = text
    except Exception as e:
        logging.error(f"Background transcription error for {filename}: {e}")
        session["upload_files"][filename]["status"] = "error"
        session["upload_files"][filename]["error"] = str(e)
    finally:
        try:
            os.unlink(filepath)
        except:
            pass


@app.post("/api/sessions/{session_id}/upload")
async def upload_audio_files(session_id: str, request: Request, files: List[UploadFile] = File(...)):
    """Upload audio files for background transcription (async, non-blocking)"""
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if not files:
        return {"success": False, "error": "請選擇至少一個音檔"}
    
    allowed_extensions = {".wav", ".mp3", ".m4a", ".ogg", ".webm", ".flac", ".aac", ".wma", ".opus"}
    max_bytes = 50 * 1024 * 1024  # 50MB per file
    results = []
    
    for file in files:
        # Validate file extension
        ext = Path(file.filename).suffix.lower() if file.filename else ".webm"
        if ext not in allowed_extensions:
            results.append({
                "filename": file.filename,
                "success": False,
                "status": "error",
                "error": f"不支援的格式: {ext}"
            })
            continue
        
        content = await file.read()
        
        # Validate file size
        if len(content) > max_bytes:
            file_size_mb = len(content) / (1024 * 1024)
            results.append({
                "filename": file.filename,
                "success": False,
                "status": "error",
                "error": f"檔案過大: {file_size_mb:.1f}MB，上限 50MB"
            })
            continue
        
        # Save to disk for background processing
        safe_name = f"{session_id}_{uuid.uuid4().hex[:8]}{ext}"
        filepath = UPLOAD_DIR / safe_name
        with open(filepath, "wb") as f:
            f.write(content)
        
        # Register in session with pending status
        sessions[session_id]["upload_files"][file.filename] = {
            "status": "pending",
            "transcription": None,
            "segment_id": None,
            "error": None
        }
        
        # Submit to background thread
        executor.submit(_process_uploaded_file, session_id, str(filepath), file.filename, ext)
        
        results.append({
            "filename": file.filename,
            "success": True,
            "status": "pending"
        })
    
    return {"success": True, "results": results}


@app.get("/api/sessions/{session_id}/upload-status")
async def get_upload_status(session_id: str, request: Request):
    """Get background upload processing status"""
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    files = sessions[session_id].get("upload_files", {})
    
    # Count statuses
    status_counts = {"pending": 0, "processing": 0, "done": 0, "error": 0}
    for fname, info in files.items():
        s = info.get("status", "pending")
        if s in status_counts:
            status_counts[s] += 1
    
    return {
        "success": True,
        "files": files,
        "counts": status_counts,
        "total": len(files),
        "all_done": status_counts["done"] + status_counts["error"] == len(files) if files else True
    }

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
        
        <div class="card instructions-card">
            <h3>📖 使用說明</h3>
            <ul class="instructions-list">
                <li>🎤 按下錄音鈕開始錄音，再按一下停止</li>
                <li>➕ 停止後可繼續錄下一段，段落沒有限制</li>
                <li>📁 也可上傳音檔（MP3/WAV/M4A 等），每檔上限 50MB，支援拖放或多選</li>
                <li>🔄 錄音與上傳可混合使用，段落會累積，不限制次數</li>
                <li>⏳ 上傳後的辨識在背景自動執行，可關閉網頁，完成後再回來查看</li>
                <li>📋 上傳區會即時顯示每檔的處理狀態（等待中→辨識中→完成）</li>
                <li>📥 所有段落都辨識完畢後，再按下「產生報告」</li>
                <li>📥 報告產出成功後，可以下載或是以 Email 發送</li>
                <li>📄 可上傳 Word 模板，AI 彙整內容會自動填入</li>
            </ul>
        </div>
        
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
            <p class="hint">使用 {{content}}、{{date}}、{{report_type}} 作為佔位符</p>
            <div id="templateStatus" class="template-status"></div>
            <input type="file" id="templateFile" accept=".docx" class="file-input">
            <p class="hint" id="templateOverwriteHint" style="display:none; margin-top:4px;">上傳新模板將會覆蓋現有模板</p>
        </div>
        
        <div class="card recorder-card">
            <h3>🎤 錄音區段</h3>
            <div class="recorder-box">
                <button class="btn-record" id="recordBtn">🎤</button>
                <div class="status" id="status">點擊麥克風開始錄音</div>
            </div>
            
            <div id="segments" class="segments-container"></div>
            
            <button class="btn btn-secondary" id="clearBtn">🗑 清空重置</button>
        </div>
        
        <div class="card upload-card">
            <h3>📁 上傳音檔</h3>
            <p class="hint">支援格式：MP3、WAV、M4A、OGG、FLAC、AAC、WebM、Opus ｜ 每檔上限 50MB</p>
            <div class="upload-box" id="uploadBox">
                <div class="upload-icon">📂</div>
                <div class="upload-text">點擊選擇檔案 或 拖放音檔到此處</div>
                <div class="upload-hint">可選擇多個檔案一次上傳，辨識在背景執行</div>
                <input type="file" id="fileInput" multiple accept=".wav,.mp3,.m4a,.ogg,.webm,.flac,.aac,.wma,.opus" class="file-input-hidden">
            </div>
            <div id="uploadProgress" class="upload-progress" style="display:none">
                <div class="progress-bar"><div class="progress-fill" id="progressFill"></div></div>
                <div class="progress-text" id="progressText">上傳中...</div>
            </div>
            <div id="uploadResults" class="upload-results"></div>
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
    const clearBtn = document.getElementById('clearBtn');
    const generateBtn = document.getElementById('generateBtn');
    const downloadBtn = document.getElementById('downloadBtn');
    const emailBtn = document.getElementById('emailBtn');
    const templateFile = document.getElementById('templateFile');
    const templateStatus = document.getElementById('templateStatus');
    const templateOverwriteHint = document.getElementById('templateOverwriteHint');
    const reportType = document.getElementById('reportType');
    const result = document.getElementById('result');
    
    // Load template status on page load
    async function loadTemplateStatus() {
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/template');
            const data = await res.json();
            if (data.success && data.has_template) {
                hasTemplate = true;
                templateStatus.innerHTML = '<span class="success">✓ 目前模板：' + data.filename + '</span>';
                templateOverwriteHint.style.display = 'block';
            } else {
                hasTemplate = false;
                templateStatus.innerHTML = '<span class="hint">目前無模板，可以上傳模板</span>';
                templateOverwriteHint.style.display = 'none';
            }
        } catch (e) {
            templateStatus.innerHTML = '<span class="hint">目前無模板，可以上傳模板</span>';
        }
    }
    loadTemplateStatus();
    
    // Template upload — upload once, persists on server, re-upload overwrites
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
                    templateStatus.innerHTML = '<span class="success">✓ 已儲存：' + data.filename + '</span>';
                    templateOverwriteHint.style.display = 'block';
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
    
    clearBtn.onclick = async () => {
        if (segments.length === 0) {
            segments = [];
            renderSegments();
            status.textContent = '已清空';
            return;
        }
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/clear', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' }
            });
            const data = await res.json();
            if (data.success) {
                segments = [];
                renderSegments();
                downloadBtn.disabled = true;
                emailBtn.disabled = true;
                generateBtn.disabled = true;
                status.textContent = '已清空重置';
                document.getElementById('result').innerHTML = '';
            }
        } catch (e) {
            alert('清除失敗：' + e.message);
        }
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

    // ========== File Upload ==========
    const uploadBox = document.getElementById('uploadBox');
    const fileInput = document.getElementById('fileInput');
    const uploadProgress = document.getElementById('uploadProgress');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    const uploadResults = document.getElementById('uploadResults');

    // Click to select files
    uploadBox.addEventListener('click', () => fileInput.click());

    // Drag & drop
    uploadBox.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadBox.classList.add('drag-over');
    });
    uploadBox.addEventListener('dragleave', () => {
        uploadBox.classList.remove('drag-over');
    });
    uploadBox.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadBox.classList.remove('drag-over');
        if (e.dataTransfer.files.length > 0) {
            fileInput.files = e.dataTransfer.files;
            handleFiles(fileInput.files);
        }
    });

    // File selected via dialog
    fileInput.addEventListener('change', () => {
        if (fileInput.files.length > 0) {
            handleFiles(fileInput.files);
        }
    });

    let uploadPollTimer = null;

    function renderFileStatus(fileInfo, filename) {
        const statusMap = {
            'pending': '<span class="badge-pending">⏳ 等待中</span>',
            'processing': '<span class="badge-processing">🔄 辨識中</span>',
            'done': '<span class="badge-done">✅ 完成</span>',
            'error': '<span class="badge-error">❌ 失敗</span>'
        };
        const badge = statusMap[fileInfo.status] || statusMap['pending'];
        let extra = '';
        if (fileInfo.status === 'done' && fileInfo.transcription) {
            extra = '<div class="file-preview">' + escapeHtml(fileInfo.transcription.substring(0, 80)) + (fileInfo.transcription.length > 80 ? '...' : '') + '</div>';
        } else if (fileInfo.status === 'error') {
            extra = '<div class="file-error-text">' + escapeHtml(fileInfo.error || '未知錯誤') + '</div>';
        }
        return '<div class="file-status-item status-' + fileInfo.status + '">' +
            '<div class="file-status-header">' + badge + ' <span class="result-filename">' + escapeHtml(filename) + '</span></div>' +
            extra + '</div>';
    }

    function renderUploadPanel(files) {
        let html = '';
        for (const [fname, info] of Object.entries(files)) {
            html += renderFileStatus(info, fname);
        }
        uploadResults.innerHTML = html;
    }

    async function pollUploadStatus() {
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/upload-status');
            const data = await res.json();
            if (!data.success) return;

            renderUploadPanel(data.files);

            // Update progress
            const done = data.counts.done + data.counts.error;
            const total = data.total;
            const pct = total > 0 ? Math.round(done / total * 100) : 0;
            progressFill.style.width = pct + '%';

            if (data.all_done) {
                progressText.textContent = '✓ ' + done + '/' + total + ' 處理完成';
                if (uploadPollTimer) {
                    clearInterval(uploadPollTimer);
                    uploadPollTimer = null;
                }
                // Rebuild segments from completed uploads + check generate readiness
                segments = [];
                for (const [fname, info] of Object.entries(data.files)) {
                    if (info.status === 'done' && info.segment_id) {
                        segments.push({ id: info.segment_id, text: info.transcription });
                    }
                }
                renderSegments();
                checkGenerateReady();
                status.textContent = '✓ ' + segments.length + ' 個段落辨識完成（可關閉此頁，回來繼續操作）';
            } else {
                progressText.textContent = '處理中 ' + done + '/' + total + '（可關閉此頁，完成後回來查看）';
                status.textContent = '🔄 語音辨識背景處理中...';
            }
        } catch (e) {
            // Silently retry on next poll
        }
    }

    async function handleFiles(files) {
        const MAX_MB = 50;
        // Frontend check: file size limit
        for (const file of files) {
            if (file.size > MAX_MB * 1024 * 1024) {
                uploadResults.innerHTML = '<div class="file-status-item status-error"><span class="badge-error">❌</span> <span class="result-filename">' + escapeHtml(file.name) + '</span> 超過 ' + MAX_MB + 'MB 限制（' + (file.size / 1024 / 1024).toFixed(1) + 'MB）</div>';
                status.textContent = '❌ 部分檔案超過大小限制';
                return;
            }
        }
        
        const formData = new FormData();
        for (const file of files) {
            formData.append('files', file);
        }

        // Show progress area
        uploadProgress.style.display = 'block';
        progressFill.style.width = '5%';
        progressText.textContent = '正在上傳 ' + files.length + ' 個檔案...';
        uploadResults.innerHTML = '';

        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/upload', {
                method: 'POST',
                body: formData
            });
            const data = await res.json();

            if (data.success) {
                progressFill.style.width = '10%';
                progressText.textContent = '已接收，背景辨識中...';
                status.textContent = '🔄 語音辨識背景處理中，可關閉此頁...';

                // Initial render
                const initFiles = {};
                data.results.forEach(r => {
                    if (r.success) {
                        initFiles[r.filename] = { status: 'pending', transcription: null, segment_id: null, error: null };
                    } else {
                        initFiles[r.filename] = { status: 'error', error: r.error };
                    }
                });
                renderUploadPanel(initFiles);

                // Start polling
                if (uploadPollTimer) clearInterval(uploadPollTimer);
                uploadPollTimer = setInterval(pollUploadStatus, 2500);
            } else {
                uploadResults.innerHTML = '<div class="file-status-item status-error"><span class="badge-error">❌</span> 上傳失敗：' + escapeHtml(data.error || '未知錯誤') + '</div>';
            }
        } catch (e) {
            progressFill.style.width = '0%';
            progressText.textContent = '上傳錯誤';
            uploadResults.innerHTML = '<div class="file-status-item status-error"><span class="badge-error">❌</span> 錯誤：' + escapeHtml(e.message) + '</div>';
        }

        // Reset file input
        fileInput.value = '';
    }
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

@app.get("/api/sessions/{session_id}/template")
async def get_template_status(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    template_path = TEMPLATES_DIR / "template.docx"
    if template_path.exists():
        return {"success": True, "has_template": True, "filename": "template.docx"}
    else:
        return {"success": True, "has_template": False}

@app.post("/api/sessions/{session_id}/template")
async def upload_template(session_id: str, request: Request, file: UploadFile = File(...)):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    if not file.filename.endswith('.docx'):
        return {"success": False, "error": "請上傳 .docx 檔案"}
    
    # Save as the one persistent template (overwrite previous)
    template_path = TEMPLATES_DIR / "template.docx"
    with open(template_path, "wb") as f:
        content = await file.read()
        f.write(content)
    
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
        future = executor.submit(do_transcribe)
        text = future.result(timeout=120)
    except TimeoutError:
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

@app.post("/api/sessions/{session_id}/clear")
async def clear_session(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    # Clear segments + upload files
    sessions[session_id]["segments"] = []
    sessions[session_id]["generated_docx"] = None
    sessions[session_id]["upload_files"] = {}
    
    return {"success": True}

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
    
    # Build full text from all segments
    full_text = "\n\n".join([
        f"【段落{i+1}】\n{seg['transcription']}" 
        for i, seg in enumerate(session["segments"])
    ])
    
    # Check if template has specific fields
    template_path = str(TEMPLATES_DIR / "template.docx")
    template_fields = []
    template_usable = False
    if Path(template_path).exists():
        try:
            template_fields = extract_placeholders(template_path)
            template_usable = True
            logging.warning(f"[Generate] Template fields found: {template_fields}")
        except Exception as e:
            logging.error(f"[Generate] Template corrupted, skipping: {e}")
    
    # Summarize with MiniMax AI
    api_key_preview = MINIMAX_API_KEY[:10] + "..." if MINIMAX_API_KEY else "EMPTY"
    logging.warning(f"[Generate] MINIMAX_API_KEY: {api_key_preview}, length: {len(MINIMAX_API_KEY) if MINIMAX_API_KEY else 0}")
    
    fields_dict = None
    summarized_text = None
    
    if MINIMAX_API_KEY:
        logging.warning(f"[Generate] Calling MiniMax API with placeholders={template_fields}")
        result = summarize_with_hermes(full_text, report_type, placeholders=template_fields if template_fields else None)
        
        if template_fields:
            # Structured mode — MiniMax returned a dict
            fields_dict = result
            logging.warning(f"[Generate] MiniMax returned fields: {list(fields_dict.keys()) if fields_dict else 'none'}")
        else:
            # Free mode — MiniMax returned a string
            summarized_text = result
            logging.warning(f"[Generate] MiniMax returned {len(summarized_text)} chars")
    else:
        logging.warning("[Generate] MINIMAX_API_KEY is empty - skipping AI summarization")
        summarized_text = full_text
    
    # If template exists and is usable, fill it
    if template_usable:
        try:
            docx_bytes = fill_template(
                template_path,
                session["segments"],
                report_type,
                summarized_text=summarized_text,
                fields_dict=fields_dict
            )
            session["generated_docx"] = docx_bytes
            return {"success": True, "has_template": True}
        except Exception as e:
            logging.error(f"[Generate] Template fill failed, falling back: {e}")
            session["generated_docx"] = None
            return {"success": True, "has_template": False, "text": summarized_text}
    else:
        # No template - return summarized text only
        return {"success": True, "has_template": False, "text": summarized_text}

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
                "Content-Disposition": f"attachment; filename*=UTF-8''MCH_%E5%A0%B1%E5%91%8A_{datetime.now().strftime('%Y%m%d')}.docx"
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
                "Content-Disposition": f"attachment; filename*=UTF-8''MCH_%E5%A0%B1%E5%91%8A_{datetime.now().strftime('%Y%m%d')}.txt"
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
        import urllib.error
        
        # Build email content
        report_date = datetime.now().strftime('%Y年%m月%d日')
        report_type_name = get_report_type_name(session.get("report_type", "general"))
        
        # Try to get summarized text from segments or build a cleaner body
        # Build a cleaner body using all segments (but without the 【段落】 markers)
        body_lines = []
        for i, seg in enumerate(session["segments"]):
            body_lines.append(f"【記錄 {i+1}】\n{seg['transcription']}")
        segment_text = "\n\n".join(body_lines)
        
        # Create email
        msg = MIMEMultipart('mixed')
        msg['to'] = req.to_email
        msg['subject'] = f"MCH {report_type_name} - {report_date}"
        
        # Email body (brief intro + mention attachment)
        body = f"""您好，

這是由 MCH Assistant 產生的 {report_type_name}，日期：{report_date}。

彙整報告已作為 Word 檔案 (.docx) 附加於此郵件中，請直接下載開啟。

---
此郵件由 MCH Assistant 自動發送
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        
        # Attach the Word docx if available
        if session.get("generated_docx"):
            from email.mime.base import MIMEBase
            from email import encoders
            clean_date = report_date.replace("年", "").replace("月", "").replace("日", "")
            filename = f"MCH_{report_type_name}_{clean_date}.docx"
            
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
            part.set_payload(session["generated_docx"])
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
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