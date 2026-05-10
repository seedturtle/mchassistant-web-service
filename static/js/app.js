// MCH Assistant Web Service - JavaScript (Multi-Segment Recording)

let mediaRecorder;
let audioChunks = [];
let isRecording = false;
let segments = [];
let hasTemplate = false;
let SESSION_ID = '';
let recordingTimer = null;

// DOM Elements
const recordBtn = document.getElementById('recordBtn');
const status = document.getElementById('status');
const segmentsDiv = document.getElementById('segments');
const clearBtn = document.getElementById('clearBtn');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const emailBtn = document.getElementById('emailBtn');
const templateFile = document.getElementById('templateFile');
const templateStatus = document.getElementById('templateStatus');
const templateList = document.getElementById('templateList');
const uploadTemplateBtn = document.getElementById('uploadTemplateBtn');
const reportType = document.getElementById('reportType');
const result = document.getElementById('result');

let selectedTemplate = null;

// Load saved templates
async function loadTemplates() {
    try {
        const res = await fetch('/api/templates');
        const data = await res.json();
        if (data.success) {
            renderTemplates(data.templates);
        }
    } catch (e) {
        console.error('Failed to load templates:', e);
    }
}

function renderTemplates(templates) {
    if (!templateList) return;
    if (templates.length === 0) {
        templateList.innerHTML = '<div class="template-empty">尚未上傳任何模板</div>';
        return;
    }
    templateList.innerHTML = templates.map(t => 
        `<div class="template-item${selectedTemplate === t.name ? ' active' : ''}" data-name="${t.name}">
            <span class="name">📄 ${t.name}</span>
            <button class="delete-btn" data-name="${t.name}" title="刪除模板">✕</button>
        </div>`
    ).join('');
    
    templateList.querySelectorAll('.template-item').forEach(el => {
        el.addEventListener('click', (e) => {
            if (e.target.classList.contains('delete-btn')) return;
            selectTemplate(el.dataset.name);
        });
    });
    
    templateList.querySelectorAll('.delete-btn').forEach(btn => {
        btn.addEventListener('click', async (e) => {
            e.stopPropagation();
            const name = btn.dataset.name;
            if (!confirm(`刪除模板「${name}」？`)) return;
            try {
                const res = await fetch('/api/templates/' + encodeURIComponent(name), {
                    method: 'DELETE'
                });
                const data = await res.json();
                if (data.success) {
                    if (selectedTemplate === name) {
                        selectedTemplate = null;
                        hasTemplate = false;
                        templateStatus.innerHTML = '';
                    }
                    await loadTemplates();
                }
            } catch (e) {
                alert('刪除失敗：' + e.message);
            }
        });
    });
}

async function selectTemplate(name) {
    try {
        const res = await fetch('/api/sessions/' + SESSION_ID + '/template/select', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ template_name: name })
        });
        const data = await res.json();
        if (data.success) {
            selectedTemplate = name;
            hasTemplate = true;
            templateStatus.innerHTML = '<span class="success">✓ 已選取：' + name + '</span>';
            checkGenerateReady();
        } else {
            templateStatus.innerHTML = '<span class="error">選取失敗</span>';
        }
    } catch (e) {
        templateStatus.innerHTML = '<span class="error">選取失敗：' + e.message + '</span>';
    }
    renderTemplates(await fetchSavedTemplates());
}

async function fetchSavedTemplates() {
    try {
        const res = await fetch('/api/templates');
        const data = await res.json();
        if (data.success) return data.templates;
    } catch(e) {}
    return [];
}

// Template upload
if (templateFile && uploadTemplateBtn) {
    templateFile.addEventListener('change', () => {
        uploadTemplateBtn.disabled = templateFile.files.length === 0;
    });
    
    uploadTemplateBtn.addEventListener('click', async () => {
        if (templateFile.files.length === 0) return;
        const file = templateFile.files[0];
        const formData = new FormData();
        formData.append('file', file);
        formData.append('session_id', SESSION_ID);
        
        uploadTemplateBtn.textContent = '上傳中...';
        uploadTemplateBtn.disabled = true;
        
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/template', {
                method: 'POST',
                body: formData
            });
            const data = await res.json();
            if (data.success) {
                selectedTemplate = data.filename;
                hasTemplate = true;
                templateStatus.innerHTML = '<span class="success">✓ 已上傳並選取：' + data.filename + '</span>';
                templateFile.value = '';
                uploadTemplateBtn.disabled = true;
                checkGenerateReady();
                await loadTemplates();
            } else {
                templateStatus.innerHTML = '<span class="error">上傳失敗：' + data.error + '</span>';
            }
        } catch (e) {
            templateStatus.innerHTML = '<span class="error">上傳失敗</span>';
        }
        uploadTemplateBtn.textContent = '上傳';
        uploadTemplateBtn.disabled = templateFile.files.length === 0;
    });
}

// Load templates on startup
loadTemplates();

// Recording Control
if (recordBtn) {
    recordBtn.addEventListener('click', async () => {
        if (!isRecording) {
            try {
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
                            status.textContent = '轉換中...';
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
                                status.textContent = '錄音完成';
                            } else {
                                alert('轉換失敗');
                                status.textContent = '轉換失敗';
                            }
                        } catch (e) {
                            alert('錯誤：' + e.message);
                            status.textContent = '錯誤';
                        }
                        clearInterval(recordingTimer);
                        stream.getTracks().forEach(t => t.stop());
                    };
                };
                
                mediaRecorder.start();
                isRecording = true;
                recordBtn.textContent = '⏹️';
                
                // 5分鐘計時器
                let remaining = 300;
                const timerDisplay = setInterval(() => {
                    remaining--;
                    const mins = Math.floor(remaining / 60);
                    const secs = remaining % 60;
                    status.textContent = `錄音中... 再次點擊停止（還有 ${mins}:${secs.toString().padStart(2, '0')}）`;
                    if (remaining <= 0) {
                        mediaRecorder.stop();
                        clearInterval(timerDisplay);
                    }
                }, 1000);
                recordingTimer = timerDisplay;
            } catch (e) {
                alert('無法存取麥克風：' + e.message);
            }
        } else {
            mediaRecorder.stop();
            isRecording = false;
            recordBtn.textContent = '🎤';
        }
    });
}

// Render Segments
function renderSegments() {
    if (!segmentsDiv) return;
    segmentsDiv.innerHTML = segments.map((seg, i) => 
        '<div class="segment-item"><strong>段落' + (i+1) + ':</strong><br>' + escapeHtml(seg.text) + '</div>'
    ).join('');
}

// Escape HTML
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// Check if ready to generate
function checkGenerateReady() {
    if (!generateBtn) return;
    generateBtn.disabled = segments.length === 0;
}

// Clear Button
if (clearBtn) {
    clearBtn.addEventListener('click', async () => {
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
                if (result) result.innerHTML = '';
            }
        } catch (e) {
            alert('清除失敗：' + e.message);
        }
    });
}

// Generate Report
if (generateBtn) {
    generateBtn.addEventListener('click', async () => {
        generateBtn.textContent = '處理中...';
        generateBtn.disabled = true;
        
        try {
            const res = await fetch('/api/sessions/' + SESSION_ID + '/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ report_type: reportType ? reportType.value : 'general' })
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
    });
}

// Download Report
if (downloadBtn) {
    downloadBtn.addEventListener('click', () => {
        window.location.href = '/api/sessions/' + SESSION_ID + '/download';
    });
}

// Email Report
if (emailBtn) {
    emailBtn.addEventListener('click', async () => {
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
    });
}

// Initialize session ID from page
document.addEventListener('DOMContentLoaded', () => {
    const sessionMatch = document.body.innerHTML.match(/SESSION_ID = "([^"]+)"/);
    if (sessionMatch) {
        SESSION_ID = sessionMatch[1];
    }
});