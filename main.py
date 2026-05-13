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

from fastapi import FastAPI, HTTPException, Request, UploadFile, File, Form
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
RECORDING_DIR = Path("./recordings")
RECORDING_DIR.mkdir(exist_ok=True)

MATON_API_KEY = os.getenv("MATON_API_KEY", "")
GMAIL_GATEWAY = "https://gateway.maton.ai/google-mail/gmail/v1/users/me/messages/send"

HF_TOKEN = os.getenv("HF_TOKEN", "")
MINIMAX_API_KEY = os.getenv("MINIMAX_API_KEY", "")
MINIMAX_API_GATEWAY = "https://api.minimax.io/v1/text/chatcompletion_v2"

sessions = {}
executor = ThreadPoolExecutor(max_workers=2)

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
    version="2.1.0",
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
    session_id = str(uuid.uuid4())[:8]
    sessions[session_id] = {
        "report_type": "general",
        "segments": [],            # final transcribed text segments
        "audio_files": {},         # {uid: {filename, filepath, ext, source: "upload"|"record"}}
        "generated_docx": None,
        "created_at": datetime.now(),
        "auto_email": "",
        "auto_email_sent": False,
        "auto_email_message": None,
        "auto_email_error": None,
        "processing": False,
        "processing_done": False,
        "processing_error": None,
        "processing_progress": None,  # {transcribed, total, current_file, stage}
    }
    return session_id

def transcribe_audio(audio_bytes: bytes, file_ext: str = ".webm") -> str:
    try:
        from faster_whisper import WhisperModel
        import os as os_module
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
    logging.info(f"[summarize_with_hermes] report_type={report_type}, text_length={len(transcribed_text)}")
    if not MINIMAX_API_KEY:
        if placeholders:
            return {p: transcribed_text for p in placeholders}
        return transcribed_text
    
    try:
        import urllib.request
        import urllib.error
        
        report_type_names = {
            "general": "一般報告", "medical": "醫療報告", "meeting": "會議記錄",
            "swallow": "吞嚥評估", "ent": "耳鼻喉科報告"
        }
        type_name = report_type_names.get(report_type, "報告")
        
        if placeholders:
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
            if 'choices' in result and len(result['choices']) > 0:
                content = result['choices'][0]['message']['content'].strip()
                if placeholders:
                    try:
                        clean = content.strip()
                        if clean.startswith("```"):
                            clean = clean.split("\n", 1)[1] if "\n" in clean else clean
                            clean = clean.rsplit("```", 1)[0] if "```" in clean else clean
                            clean = clean.strip()
                        fields_result = json.loads(clean)
                        for p in placeholders:
                            if p not in fields_result:
                                fields_result[p] = f"[待補充：{p}]"
                        return fields_result
                    except (json.JSONDecodeError, Exception) as e:
                        return {p: content if i == 0 else f"[待補充：{p}]" for i, p in enumerate(placeholders)}
                else:
                    return content
            if placeholders:
                return {p: f"[待補充：{p}]" for p in placeholders}
            return transcribed_text
    except Exception as e:
        logging.error(f"MiniMax summarization error: {e}")
        if placeholders:
            return {p: transcribed_text for p in placeholders}
        return transcribed_text

def fill_template(template_path: str, segments: list, report_type: str, summarized_text: str = None, fields_dict: dict = None) -> bytes:
    doc = Document(template_path)
    if summarized_text:
        full_text = summarized_text
    else:
        full_text = "\n\n".join([f"【段落{i+1}】\n{seg['transcription']}" for i, seg in enumerate(segments)])
    report_date = datetime.now().strftime('%Y年%m月%d日')
    report_name = get_report_type_name(report_type)
    
    def replace_paragraph_text(para):
        full_para_text = "".join(run.text for run in para.runs)
        new_text = full_para_text
        nonlocal content_replaced
        if "{{content}}" in new_text:
            new_text = new_text.replace("{{content}}", full_text)
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
        if new_text != full_para_text:
            for i, run in enumerate(para.runs):
                if i == 0:
                    run.text = new_text
                else:
                    run.text = ""
    
    content_replaced = False
    for para in doc.paragraphs:
        replace_paragraph_text(para)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    replace_paragraph_text(para)
    if not content_replaced and full_text:
        doc.add_paragraph("")
        doc.add_paragraph(full_text)
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

def get_report_type_name(report_type: str) -> str:
    names = {"general": "一般報告", "medical": "醫療報告", "meeting": "會議記錄", "swallow": "吞嚥評估", "ent": "耳鼻喉科報告"}
    return names.get(report_type, report_type)

def extract_placeholders(doc_path: str) -> list:
    doc = Document(doc_path)
    placeholders = set()
    for para in doc.paragraphs:
        placeholders.update(re.findall(r'\{\{(.*?)\}\}', para.text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                placeholders.update(re.findall(r'\{\{(.*?)\}\}', cell.text))
    auto_fill = {"date", "report_type", "content"}
    return [p for p in placeholders if p not in auto_fill]

def _send_email_sync(session: dict, to_email: str) -> tuple:
    try:
        import urllib.request, urllib.error
        report_date = datetime.now().strftime('%Y年%m月%d日')
        report_type_name = get_report_type_name(session.get("report_type", "general"))
        msg = MIMEMultipart('mixed')
        msg['to'] = to_email
        msg['subject'] = f"MCH {report_type_name} - {report_date}"
        body = f"""您好，

這是由 MCH Assistant 產生的 {report_type_name}，日期：{report_date}。

彙整報告已作為 Word 檔案 (.docx) 附加於此郵件中，請直接下載開啟。

---
此郵件由 MCH Assistant 自動發送
"""
        msg.attach(MIMEText(body, 'plain', 'utf-8'))
        if session.get("generated_docx"):
            from email.mime.base import MIMEBase
            from email import encoders
            clean_date = report_date.replace("年", "").replace("月", "").replace("日", "")
            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
            part.set_payload(session["generated_docx"])
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="MCH_{report_type_name}_{clean_date}.docx"')
            msg.attach(part)
        raw = base64.urlsafe_b64encode(msg.as_bytes()).decode().rstrip('=')
        data = json.dumps({"raw": raw}).encode()
        gmail_req = urllib.request.Request(GMAIL_GATEWAY, data=data, method='POST')
        gmail_req.add_header('Authorization', f'Bearer {MATON_API_KEY}')
        gmail_req.add_header('Content-Type', 'application/json')
        with urllib.request.urlopen(gmail_req, timeout=30) as resp:
            return True, f"報告已自動寄送至 {to_email}"
    except Exception as e:
        return False, str(e)

# =============================================================================
# Background Processing Pipeline
# =============================================================================

def _process_generate(session_id: str):
    """Background pipeline: transcribe all → summarize → Word → Email"""
    session = sessions.get(session_id)
    if not session:
        return
    
    session["processing"] = True
    session["processing_done"] = False
    session["segments"] = []
    
    audio_files = session.get("audio_files", {})
    total = len(audio_files)
    session["processing_progress"] = {"stage": "transcribing", "transcribed": 0, "total": total, "current_file": "", "error": None}
    
    try:
        # --- Step 1: Transcribe all audio files ---
        for uid, info in sorted(audio_files.items(), key=lambda x: x[1].get("created_at", 0)):
            session["processing_progress"]["current_file"] = info["filename"]
            try:
                with open(info["filepath"], "rb") as f:
                    audio_bytes = f.read()
                text = transcribe_audio(audio_bytes, file_ext=info["ext"])
            except Exception as e:
                text = f"[轉換失敗: {str(e)}]"
            
            session["segments"].append({
                "id": uid,
                "transcription": text,
                "audio_path": info["filename"]
            })
            session["processing_progress"]["transcribed"] += 1
        
        # --- Step 2: Summarize with MiniMax ---
        session["processing_progress"]["stage"] = "summarizing"
        full_text = "\n\n".join([f"【段落{i+1}】\n{seg['transcription']}" for i, seg in enumerate(session["segments"])])
        report_type = session.get("report_type", "general")
        
        template_path = str(TEMPLATES_DIR / "template.docx")
        template_fields = []
        template_usable = False
        if Path(template_path).exists():
            try:
                template_fields = extract_placeholders(template_path)
                template_usable = True
            except Exception:
                pass
        
        fields_dict = None
        summarized_text = None
        
        if MINIMAX_API_KEY:
            result = summarize_with_hermes(full_text, report_type, placeholders=template_fields if template_fields else None)
            if template_fields:
                fields_dict = result
            else:
                summarized_text = result
        else:
            summarized_text = full_text
        
        # --- Step 3: Generate Word ---
        session["processing_progress"]["stage"] = "generating"
        if template_usable:
            docx_bytes = fill_template(template_path, session["segments"], report_type,
                                       summarized_text=summarized_text, fields_dict=fields_dict)
            session["generated_docx"] = docx_bytes
        
        # --- Step 4: Send email if configured ---
        target_email = session.get("auto_email", "")
        if target_email:
            session["processing_progress"]["stage"] = "emailing"
            success, msg = _send_email_sync(session, target_email)
            if success:
                session["auto_email_sent"] = True
                session["auto_email_message"] = msg
            else:
                session["auto_email_error"] = msg
        
        session["processing_progress"]["stage"] = "done"
        session["processing_done"] = True
        
    except Exception as e:
        logging.error(f"Background generate error: {e}")
        session["processing_error"] = str(e)
        if session.get("processing_progress"):
            session["processing_progress"]["error"] = str(e)
            session["processing_progress"]["stage"] = "error"
        session["processing_done"] = True
    finally:
        session["processing"] = False


# =============================================================================
# API Endpoints
# =============================================================================

# --- Upload audio files (save only, no transcription) ---
@app.post("/api/sessions/{session_id}/upload")
async def upload_audio_files(
    session_id: str,
    request: Request,
    files: List[UploadFile] = File(...),
    email: str = Form("")
):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    if not files:
        return {"success": False, "error": "請選擇至少一個音檔"}
    
    if email:
        sessions[session_id]["auto_email"] = email
        sessions[session_id]["auto_email_sent"] = False
        sessions[session_id]["auto_email_message"] = None
        sessions[session_id]["auto_email_error"] = None
    
    allowed_extensions = {".wav", ".mp3", ".m4a", ".ogg", ".webm", ".flac", ".aac", ".wma", ".opus"}
    max_bytes = 50 * 1024 * 1024
    results = []
    
    for file in files:
        ext = Path(file.filename).suffix.lower() if file.filename else ".webm"
        if ext not in allowed_extensions:
            results.append({"filename": file.filename, "success": False, "error": f"不支援的格式: {ext}"})
            continue
        content = await file.read()
        if len(content) > max_bytes:
            results.append({"filename": file.filename, "success": False, "error": f"檔案過大: {len(content)/(1024*1024):.1f}MB，上限 50MB"})
            continue
        
        uid = uuid.uuid4().hex[:8]
        safe_name = f"{session_id}_{uid}{ext}"
        filepath = UPLOAD_DIR / safe_name
        with open(filepath, "wb") as f:
            f.write(content)
        
        sessions[session_id]["audio_files"][uid] = {
            "filename": file.filename, "filepath": str(filepath), "ext": ext,
            "size": len(content), "source": "upload", "created_at": datetime.now().timestamp()
        }
        results.append({"uid": uid, "filename": file.filename, "success": True})
    
    return {"success": True, "results": results, "auto_email": email if email else None}


# --- Record audio segment (save only, no transcription) ---
@app.post("/api/sessions/{session_id}/segment")
async def add_segment(session_id: str, request: Request, seg: AudioSegmentRequest):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    audio_bytes = base64.b64decode(seg.audio_data)
    uid = uuid.uuid4().hex[:8]
    ext = f".{seg.format}" if seg.format else ".webm"
    safe_name = f"{session_id}_record_{uid}{ext}"
    filepath = RECORDING_DIR / safe_name
    with open(filepath, "wb") as f:
        f.write(audio_bytes)
    
    sessions[session_id]["audio_files"][uid] = {
        "filename": f"錄音段落{list(sessions[session_id]['audio_files'].keys()).count('') + 1}",
        "filepath": str(filepath), "ext": ext,
        "size": len(audio_bytes), "source": "record", "created_at": datetime.now().timestamp()
    }
    # Give it a friendly name
    idx = len([v for v in sessions[session_id]["audio_files"].values() if v["source"] == "record"])
    sessions[session_id]["audio_files"][uid]["filename"] = f"錄音段落 {idx}"
    
    return {"success": True, "segment_id": uid}


# --- Get list of audio files in session ---
@app.get("/api/sessions/{session_id}/audio-files")
async def get_audio_files(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    return {
        "success": True,
        "files": sessions[session_id].get("audio_files", {}),
        "total": len(sessions[session_id].get("audio_files", {}))
    }


# --- Generate report (start background pipeline) ---
@app.post("/api/sessions/{session_id}/generate")
async def generate_report(session_id: str, request: Request, req: dict):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    session = sessions[session_id]
    audio_files = session.get("audio_files", {})
    if not audio_files:
        return {"success": False, "error": "尚無音檔，請先錄音或上傳"}
    
    if session.get("processing", False):
        return {"success": False, "error": "已有處理程序正在執行"}
    
    report_type = req.get("report_type", "general")
    session["report_type"] = report_type
    
    # Reset processing state
    session["processing"] = True
    session["processing_done"] = False
    session["processing_error"] = None
    session["processing_progress"] = {"stage": "queued", "transcribed": 0, "total": len(audio_files), "current_file": ""}
    
    # Submit to background
    executor.submit(_process_generate, session_id)
    
    return {"success": True, "message": f"背景處理已啟動，共 {len(audio_files)} 個音檔", "total": len(audio_files)}


# --- Get generation progress ---
@app.get("/api/sessions/{session_id}/generate-status")
async def get_generate_status(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    
    s = sessions[session_id]
    return {
        "success": True,
        "processing": s.get("processing", False),
        "processing_done": s.get("processing_done", False),
        "processing_error": s.get("processing_error"),
        "processing_progress": s.get("processing_progress"),
        "segments_count": len(s.get("segments", [])),
        "has_docx": s.get("generated_docx") is not None,
        "auto_email": s.get("auto_email") or None,
        "auto_email_sent": s.get("auto_email_sent", False),
        "auto_email_message": s.get("auto_email_message"),
        "auto_email_error": s.get("auto_email_error")
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
    html = """<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MCH Assistant - 登入</title><link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
    <div class="login-container">
        <h1>🏥 MCH Assistant</h1><p class="subtitle">門諾醫院 AI 語音助理</p>
        <p id="error" class="error" style="display:none">密碼錯誤，請重新輸入</p>
        <form id="loginForm">
            <div class="input-group"><label for="password">🔐 請輸入存取密碼</label>
            <input type="password" id="password" name="password" placeholder="輸入密碼" required></div>
            <button type="submit" class="btn">登入</button>
        </form>
        <p class="info">僅限醫院內部人員使用</p>
    </div>
    <script>
    document.getElementById('loginForm').onsubmit = async (e) => {
        e.preventDefault();
        const pwd = document.getElementById('password').value;
        const res = await fetch('/api/auth/login', {method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({password:pwd})});
        if (res.ok) window.location.href='/dashboard';
        else document.getElementById('error').style.display='block';
    };
    </script>
</body></html>"""
    return HTMLResponse(content=html)

@app.get("/dashboard", response_class=HTMLResponse)
async def dashboard(request: Request):
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    session_id = get_session_id(request)
    
    html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
    <meta charset="UTF-8"><meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>MCH Assistant - 儀表板</title><link rel="stylesheet" href="/static/css/style.css">
</head>
<body>
    <div class="header">
        <div class="logo">🏥 MCH Assistant</div>
        <div><span class="session-id">Session: {session_id}</span><a href="/new" class="logout">🔄 新開工作</a></div>
    </div>
    <div class="container">
        <h1>🎙️ 語音報告助理</h1>
        <p class="subtitle">錄音或上傳音檔，完成後按下「🔄 批次產出報告」</p>
        
        <div class="card instructions-card">
            <h3>📖 使用說明</h3>
            <ul class="instructions-list">
                <li>🎤 按錄音鈕開始，再按一下停止，段落會自動列出</li>
                <li>📁 也可上傳音檔（MP3/WAV/M4A 等），支援拖放或多選，每檔上限 50MB</li>
                <li>🔄 錄音與上傳的音檔會累積在列表中，可混合使用</li>
                <li>⚡ 按下「即時處理」在頁面上觀看進度，完成後手動下載或寄送</li>
                <li>📧 按下「背景處理並寄送」需先輸入 Email，完成後自動寄送到信箱，可關閉網頁</li>
                <li>📥 處理完成後也可手動下載或按「手動寄送」</li>
                <li>📄 可上傳 Word 模板（{{content}}、{{date}}、{{report_type}}），AI 內容自動填入</li>
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
            <h3>📄 Word 模板（選填）</h3>
            <p class="hint">使用 {{content}}、{{date}}、{{report_type}} 作為佔位符</p>
            <div id="templateStatus" class="template-status"></div>
            <input type="file" id="templateFile" accept=".docx" class="file-input">
        </div>
        
        <div class="card recorder-card">
            <h3>🎤 即時錄音</h3>
            <div class="recorder-box">
                <button class="btn-record" id="recordBtn">🎤</button>
                <div class="status" id="status">點擊麥克風開始錄音</div>
            </div>
        </div>
        
        <div class="card upload-card">
            <h3>📁 上傳音檔</h3>
            <p class="hint">支援格式：MP3、WAV、M4A、OGG、FLAC、AAC、WebM、Opus ｜ 每檔上限 50MB</p>
            <div class="email-input-group">
                <label for="emailInput">📧 背景模式 Email（必填）：按「背景處理並寄送」時，完成後自動寄至此信箱</label>
                <input type="email" id="emailInput" placeholder="example@mch.org.tw" class="input-full">
            </div>
            <div class="upload-box" id="uploadBox">
                <div class="upload-icon">📂</div>
                <div class="upload-text">點擊選擇檔案 或 拖放音檔到此處</div>
                <div class="upload-hint">可選擇多個檔案一次上傳</div>
                <input type="file" id="fileInput" multiple accept=".wav,.mp3,.m4a,.ogg,.webm,.flac,.aac,.wma,.opus" class="file-input-hidden">
            </div>
        </div>
        
        <div class="card">
            <h3>📋 音檔列表 <span id="fileCount" class="file-count">(0)</span></h3>
            <div id="audioFileList" class="audio-file-list"></div>
            <div class="clear-btn-wrap"><button class="btn btn-secondary" id="clearBtn">🗑 清空重置</button></div>
        </div>
        
        <div id="processingArea" class="card processing-card" style="display:none">
            <h3>⚙️ 批次處理進度</h3>
            <div class="progress-bar"><div class="progress-fill" id="genProgressFill"></div></div>
            <div class="progress-text" id="genProgressText">準備中...</div>
            <div id="genSegments" class="segments-container" style="max-height:200px;margin-top:12px"></div>
            <div id="genResult" style="margin-top:8px"></div>
        </div>
        
        <div class="actions">
            <button class="btn btn-email-mode" id="bgEmailBtn">📧 背景處理並寄送</button>
            <button class="btn btn-generate" id="generateBtn">⚡ 即時處理</button>
            <button class="btn btn-download" id="downloadBtn" disabled>📥 下載報告</button>
            <button class="btn btn-email" id="emailBtn" disabled>📧 手動寄送</button>
        </div>
        
        <div id="result" class="result-box"></div>
    </div>
    
    <script>
    let mediaRecorder, audioChunks = [], isRecording = false;
    let SESSION_ID = '{session_id}';
    let recordingTimer = null;
    let genPollTimer = null;
    let isProcessing = false;

    function escapeHtml(t) {{ const d = document.createElement('div'); d.textContent = t; return d.innerHTML; }}

    // Refresh audio file list from server
    async function refreshFileList() {{
        try {{
            const res = await fetch('/api/sessions/' + SESSION_ID + '/audio-files');
            const data = await res.json();
            if (!data.success) return;
            const files = data.files || {{}};
            const countEl = document.getElementById('fileCount');
            if (countEl) countEl.textContent = '(' + Object.keys(files).length + ')';
            const listEl = document.getElementById('audioFileList');
            if (!listEl) return;
            let html = '';
            for (const [uid, info] of Object.entries(files)) {{
                const icon = info.source === 'record' ? '🎤' : '📁';
                const sizeKB = (info.size / 1024).toFixed(1);
                html += '<div class="audio-file-item"><span class="audio-file-icon">' + icon + '</span>' +
                    '<span class="audio-file-name">' + escapeHtml(info.filename) + '</span>' +
                    '<span class="audio-file-size">' + sizeKB + 'KB</span></div>';
            }}
            listEl.innerHTML = html || '<div class="audio-file-empty">尚無音檔，請錄音或上傳</div>';
        }} catch(e) {{}}
    }}

    function disableInputsDuringProcessing(disabled) {{
        document.getElementById('recordBtn').disabled = disabled;
        document.getElementById('uploadBox').style.pointerEvents = disabled ? 'none' : 'auto';
        document.getElementById('uploadBox').style.opacity = disabled ? '0.4' : '1';
        document.getElementById('fileInput').disabled = disabled;
        document.getElementById('emailInput').disabled = disabled;
        document.getElementById('templateFile').disabled = disabled;
        document.getElementById('bgEmailBtn').disabled = disabled;
        document.getElementById('generateBtn').disabled = disabled;
    }}

    // Poll generation status
    async function pollGenerateStatus() {{
        try {{
            const res = await fetch('/api/sessions/' + SESSION_ID + '/generate-status');
            const data = await res.json();
            if (!data.success) return;
            
            const progress = data.processing_progress || {{}};
            const total = progress.total || 0;
            const transcribed = progress.transcribed || 0;
            const stage = progress.stage || '';
            
            let pct = 0;
            let stageLabel = '';
            if (stage === 'transcribing') {{
                pct = total > 0 ? Math.round(transcribed / total * 70) : 0;
                stageLabel = '🔄 語音辨識中 ' + transcribed + '/' + total;
            }} else if (stage === 'summarizing') {{
                pct = 75;
                stageLabel = '🤖 AI 彙整中';
            }} else if (stage === 'generating') {{
                pct = 90;
                stageLabel = '📄 產出 Word 報告中';
            }} else if (stage === 'emailing') {{
                pct = 95;
                stageLabel = '📧 寄送 Email 中';
            }} else if (stage === 'done') {{
                pct = 100;
                stageLabel = '✅ 完成';
            }} else if (stage === 'error') {{
                pct = 0;
                stageLabel = '❌ 處理失敗';
            }}
            
            document.getElementById('genProgressFill').style.width = pct + '%';
            document.getElementById('genProgressText').textContent = stageLabel;
            
            if (data.processing_done) {{
                if (genPollTimer) {{ clearInterval(genPollTimer); genPollTimer = null; }}
                isProcessing = false;
                
                if (data.processing_error) {{
                    document.getElementById('genResult').innerHTML = '<div class="error">❌ ' + escapeHtml(data.processing_error) + '</div>';
                }} else {{
                    if (data.auto_email_sent) {{
                        document.getElementById('genResult').innerHTML = '<div class="success">✅ 報告已自動寄送至 ' + escapeHtml(data.auto_email) + '</div>';
                    }} else if (data.auto_email) {{
                        document.getElementById('genResult').innerHTML = '<div class="warning">⚠️ 已產生報告，但自動寄信失敗：' + escapeHtml(data.auto_email_error || '') + '</div>';
                    }} else {{
                        document.getElementById('genResult').innerHTML = '<div class="success">✅ 報告已產生！</div>';
                    }}
                    document.getElementById('downloadBtn').disabled = false;
                    document.getElementById('emailBtn').disabled = false;
                }}
                disableInputsDuringProcessing(false);
            }}
        }} catch(e) {{}}
    }}

    // ========== Template Upload ==========
    document.getElementById('templateFile').addEventListener('change', async () => {{
        const file = document.getElementById('templateFile').files[0];
        if (!file) return;
        const fd = new FormData(); fd.append('file', file); fd.append('session_id', SESSION_ID);
        const res = await fetch('/api/sessions/' + SESSION_ID + '/template', {{method:'POST', body:fd}});
        const data = await res.json();
        document.getElementById('templateStatus').innerHTML = data.success 
            ? '<span class="success">✓ 已儲存：' + data.filename + '</span>' 
            : '<span class="error">上傳失敗</span>';
    }});

    // ========== Recording ==========
    document.getElementById('recordBtn').addEventListener('click', async function() {{
        if (isProcessing) {{ alert('處理中，無法錄音'); return; }}
        if (!isRecording) {{
            try {{
                const stream = await navigator.mediaDevices.getUserMedia({{audio:true}});
                mediaRecorder = new MediaRecorder(stream);
                audioChunks = [];
                mediaRecorder.ondataavailable = e => audioChunks.push(e.data);
                mediaRecorder.onstop = async () => {{
                    const blob = new Blob(audioChunks, {{type:'audio/webm'}});
                    const reader = new FileReader();
                    reader.readAsDataURL(blob);
                    reader.onloadend = async () => {{
                        document.getElementById('status').textContent = '儲存中...';
                        const res = await fetch('/api/sessions/' + SESSION_ID + '/segment', {{
                            method:'POST', headers:{{'Content-Type':'application/json'}},
                            body:JSON.stringify({{audio_data: reader.result.split(',')[1], format:'webm'}})
                        }});
                        const data = await res.json();
                        document.getElementById('status').textContent = data.success ? '✓ 錄音已儲存' : '儲存失敗';
                        await refreshFileList();
                        clearInterval(recordingTimer);
                        stream.getTracks().forEach(t => t.stop());
                    }};
                }};
                mediaRecorder.start(); isRecording = true;
                this.textContent = '⏹️';
                let remaining = 300;
                recordingTimer = setInterval(() => {{
                    remaining--;
                    const m = Math.floor(remaining/60), s = remaining%60;
                    document.getElementById('status').textContent = '錄音中... 再按停止（還有 ' + m + ':' + s.toString().padStart(2,'0') + '）';
                    if (remaining <= 0) {{ mediaRecorder.stop(); clearInterval(recordingTimer); }}
                }}, 1000);
            }} catch(e) {{ alert('無法存取麥克風：' + e.message); }}
        }} else {{
            mediaRecorder.stop(); isRecording = false; this.textContent = '🎤';
        }}
    }});

    // ========== Upload ==========
    const uploadBox = document.getElementById('uploadBox');
    const fileInput = document.getElementById('fileInput');
    uploadBox.addEventListener('click', () => {{ if (!isProcessing) fileInput.click(); else alert('處理中，無法上傳'); }});
    uploadBox.addEventListener('dragover', e => {{ e.preventDefault(); if(!isProcessing) uploadBox.classList.add('drag-over'); }});
    uploadBox.addEventListener('dragleave', () => uploadBox.classList.remove('drag-over'));
    uploadBox.addEventListener('drop', e => {{
        e.preventDefault(); uploadBox.classList.remove('drag-over');
        if (isProcessing) {{ alert('處理中，無法上傳'); return; }}
        if (e.dataTransfer.files.length > 0) {{ fileInput.files = e.dataTransfer.files; doUpload(fileInput.files); }}
    }});
    fileInput.addEventListener('change', () => {{ if(fileInput.files.length > 0) doUpload(fileInput.files); }});

    async function doUpload(files) {{
        for (const f of files) {{
            if (f.size > 50*1024*1024) {{ alert(f.name + ' 超過 50MB 限制'); return; }}
        }}
        const fd = new FormData();
        for (const f of files) fd.append('files', f);
        const email = document.getElementById('emailInput').value.trim();
        if (email) fd.append('email', email);
        
        document.getElementById('status').textContent = '上傳中...';
        try {{
            const res = await fetch('/api/sessions/' + SESSION_ID + '/upload', {{method:'POST', body:fd}});
            const data = await res.json();
            if (data.success) {{
                document.getElementById('status').textContent = '✓ ' + files.length + ' 個檔案上傳成功';
                await refreshFileList();
            }} else {{
                document.getElementById('status').textContent = '上傳失敗：' + data.error;
            }}
        }} catch(e) {{ document.getElementById('status').textContent = '上傳錯誤'; }}
        fileInput.value = '';
    }}

    // ========== Generate Report ==========
    function startProcessing(mode) {{
        document.getElementById('processingArea').style.display = 'block';
        document.getElementById('genProgressFill').style.width = '2%';
        document.getElementById('genProgressText').textContent = '啟動中...';
        document.getElementById('genResult').innerHTML = '';
        document.getElementById('genSegments').innerHTML = '';
        
        document.getElementById('bgEmailBtn').textContent = '⏳ 處理中...';
        document.getElementById('generateBtn').textContent = '⏳ 處理中...';
        
        genPollTimer = setInterval(pollGenerateStatus, 3000);
        
        if (mode === 'bg') {{
            document.getElementById('genProgressText').textContent = '📧 背景處理中，完成後將自動寄送至信箱（可關閉此頁）';
        }} else {{
            document.getElementById('genProgressText').textContent = '⚡ 即時處理中，請稍候...';
        }}
    }}

    function stopProcessing() {{
        if (genPollTimer) {{ clearInterval(genPollTimer); genPollTimer = null; }}
        isProcessing = false;
        disableInputsDuringProcessing(false);
        document.getElementById('bgEmailBtn').textContent = '📧 背景處理並寄送';
        document.getElementById('generateBtn').textContent = '⚡ 即時處理';
    }}

    // ========== 背景處理並寄送 ==========
    document.getElementById('bgEmailBtn').addEventListener('click', async function() {{
        if (isProcessing) return;
        await refreshFileList();
        const listEl = document.getElementById('audioFileList');
        if (!listEl || listEl.textContent.includes('尚無音檔')) {{
            alert('請先錄音或上傳音檔'); return;
        }}
        
        const email = document.getElementById('emailInput').value.trim();
        if (!email) {{
            alert('請先輸入 Email 再使用背景處理並寄送');
            document.getElementById('emailInput').focus();
            return;
        }}
        
        // Save email to server first by doing a dummy upload with just the email
        // Actually, the upload endpoint also accepts email. Let's set it via a simple API call
        // Use upload with empty files just to set the email... OR we can just pass it in the generate request
        
        // Simpler: set email by making a dummy API call
        // The email is stored on the upload endpoint. Let me just set it.
        // Actually, the /generate endpoint doesn't accept email. Let me store it directly.
        const setRes = await fetch('/api/sessions/' + SESSION_ID + '/set-email', {{
            method:'POST', headers:{{'Content-Type':'application/json'}},
            body:JSON.stringify({{email: email}})
        }});
        const setData = await setRes.json();
        if (!setData.success) {{
            alert('設定 Email 失敗'); return;
        }}
        
        isProcessing = true;
        disableInputsDuringProcessing(true);
        startProcessing('bg');
        
        try {{
            const res = await fetch('/api/sessions/' + SESSION_ID + '/generate', {{
                method:'POST', headers:{{'Content-Type':'application/json'}},
                body:JSON.stringify({{report_type: document.getElementById('reportType').value}})
            }});
            const data = await res.json();
            if (!data.success) {{
                document.getElementById('genResult').innerHTML = '<div class="error">啟動失敗：' + escapeHtml(data.error) + '</div>';
                stopProcessing();
            }}
        }} catch(e) {{
            document.getElementById('genResult').innerHTML = '<div class="error">錯誤：' + escapeHtml(e.message) + '</div>';
            stopProcessing();
        }}
    }});

    // ========== 即時處理 ==========
    document.getElementById('generateBtn').addEventListener('click', async function() {{
        if (isProcessing) return;
        await refreshFileList();
        const listEl = document.getElementById('audioFileList');
        if (!listEl || listEl.textContent.includes('尚無音檔')) {{
            alert('請先錄音或上傳音檔'); return;
        }}
        
        isProcessing = true;
        disableInputsDuringProcessing(true);
        startProcessing('online');
        
        try {{
            const res = await fetch('/api/sessions/' + SESSION_ID + '/generate', {{
                method:'POST', headers:{{'Content-Type':'application/json'}},
                body:JSON.stringify({{report_type: document.getElementById('reportType').value}})
            }});
            const data = await res.json();
            if (!data.success) {{
                document.getElementById('genResult').innerHTML = '<div class="error">啟動失敗：' + escapeHtml(data.error) + '</div>';
                stopProcessing();
            }}
        }} catch(e) {{
            document.getElementById('genResult').innerHTML = '<div class="error">錯誤：' + escapeHtml(e.message) + '</div>';
            stopProcessing();
        }}
    }});

    // ========== Clear ==========
    document.getElementById('clearBtn').addEventListener('click', async () => {{
        if (isProcessing) {{ alert('處理中，無法清空'); return; }}
        const res = await fetch('/api/sessions/' + SESSION_ID + '/clear', {{method:'POST', headers:{{'Content-Type':'application/json'}}}});
        const data = await res.json();
        if (data.success) {{
            document.getElementById('downloadBtn').disabled = true;
            document.getElementById('emailBtn').disabled = true;
            document.getElementById('bgEmailBtn').textContent = '📧 背景處理並寄送';
            document.getElementById('generateBtn').textContent = '⚡ 即時處理';
            document.getElementById('result').innerHTML = '';
            document.getElementById('genResult').innerHTML = '';
            document.getElementById('processingArea').style.display = 'none';
            document.getElementById('audioFileList').innerHTML = '';
            document.getElementById('fileCount').textContent = '(0)';
            document.getElementById('status').textContent = '已清空重置';
        }}
    }});

    // ========== Download & Email ==========
    document.getElementById('downloadBtn').addEventListener('click', () => {{
        window.location.href = '/api/sessions/' + SESSION_ID + '/download';
    }});
    document.getElementById('emailBtn').addEventListener('click', async () => {{
        const email = prompt('請輸入收件者 Email:');
        if (!email) return;
        document.getElementById('emailBtn').textContent = '發送中...'; document.getElementById('emailBtn').disabled = true;
        const res = await fetch('/api/sessions/' + SESSION_ID + '/email', {{
            method:'POST', headers:{{'Content-Type':'application/json'}},
            body:JSON.stringify({{to_email: email}})
        }});
        const data = await res.json();
        alert(data.success ? '✓ 郵件已發送' : '發送失敗：' + data.error);
        document.getElementById('emailBtn').textContent = '📧 手動寄送 Email';
        document.getElementById('emailBtn').disabled = false;
    }});

    // ========== Init ==========
    refreshFileList();
    </script>
</body></html>"""
    return HTMLResponse(content=html)


@app.get("/new")
async def new_session(request: Request):
    if not validate_session(request):
        return RedirectResponse(url="/login", status_code=302)
    new_id = get_or_create_session()
    response = RedirectResponse(url="/dashboard", status_code=302)
    response.set_cookie(key="session_id", value=new_id, httponly=True, samesite="lax")
    return response


# =============================================================================
# Other API Endpoints
# =============================================================================

@app.post("/api/sessions/{session_id}/set-email")
async def set_email(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    data = await request.json()
    email = data.get("email", "")
    if not email:
        return {"success": False, "error": "Email 不得為空"}
    sessions[session_id]["auto_email"] = email
    sessions[session_id]["auto_email_sent"] = False
    sessions[session_id]["auto_email_message"] = None
    sessions[session_id]["auto_email_error"] = None
    return {"success": True, "email": email}


@app.get("/api/sessions/{session_id}/template")
async def get_template_status(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    return {"exists": (TEMPLATES_DIR / "template.docx").exists()}

@app.post("/api/sessions/{session_id}/template")
async def upload_template(session_id: str, request: Request, file: UploadFile = File(...)):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if not file.filename.endswith('.docx'):
        raise HTTPException(status_code=400, detail="Only .docx files are accepted")
    content = await file.read()
    with open(TEMPLATES_DIR / "template.docx", "wb") as f:
        f.write(content)
    return {"success": True, "filename": file.filename}

@app.post("/api/auth/login")
async def login(request: Request):
    data = await request.json()
    if data.get("password") == ACCESS_PASSWORD:
        session_id = get_or_create_session()
        response = JSONResponse({"success": True})
        response.set_cookie(key="session_id", value=session_id, httponly=True, samesite="lax")
        return response
    raise HTTPException(status_code=401, detail="Invalid password")

@app.post("/api/sessions/{session_id}/clear")
async def clear_session(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    # Reset session state while keeping session alive
    s = sessions[session_id]
    s["segments"] = []
    s["audio_files"] = {}
    s["generated_docx"] = None
    s["auto_email"] = ""
    s["auto_email_sent"] = False
    s["auto_email_message"] = None
    s["auto_email_error"] = None
    s["processing"] = False
    s["processing_done"] = False
    s["processing_error"] = None
    s["processing_progress"] = None
    return {"success": True}

@app.get("/api/sessions/{session_id}/download")
async def download_report(session_id: str, request: Request):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    session = sessions[session_id]
    if session.get("generated_docx"):
        return Response(content=session["generated_docx"],
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''MCH_報名_{datetime.now().strftime('%Y%m%d')}.docx"})
    elif session.get("segments"):
        return Response(content="\n\n".join([s["transcription"] for s in session["segments"]]),
            media_type="text/plain; charset=utf-8")
    raise HTTPException(status_code=404, detail="No report to download")

@app.post("/api/sessions/{session_id}/email")
async def email_report(session_id: str, request: Request, req: EmailRequest):
    if not validate_session(request) or get_session_id(request) != session_id:
        raise HTTPException(status_code=401, detail="Not authenticated")
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    session = sessions[session_id]
    if not session.get("generated_docx") and not session.get("segments"):
        return {"success": False, "error": "無報告可發送"}
    success, msg = _send_email_sync(session, req.to_email)
    return {"success": success, "message" if success else "error": msg}

@app.get("/health")
async def health():
    return {"status": "healthy", "service": "MCH Assistant", "version": "2.1.0"}

@app.get("/api/status")
async def api_status():
    return {"service": "MCH Assistant", "version": "2.1.0", "status": "running", "sessions": len(sessions)}

# =============================================================================
# Run
# =============================================================================

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", "8080"))
    uvicorn.run(app, host="0.0.0.0", port=port)
