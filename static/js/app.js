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
const addSegmentBtn = document.getElementById('addSegmentBtn');
const generateBtn = document.getElementById('generateBtn');
const downloadBtn = document.getElementById('downloadBtn');
const emailBtn = document.getElementById('emailBtn');
const templateFile = document.getElementById('templateFile');
const templateStatus = document.getElementById('templateStatus');
const reportType = document.getElementById('reportType');
const result = document.getElementById('result');

// Template Upload
if (templateFile) {
    templateFile.addEventListener('change', async () => {
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
    });
}

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
                                addSegmentBtn.disabled = false;
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

// Add Segment Button
if (addSegmentBtn) {
    addSegmentBtn.addEventListener('click', () => {
        addSegmentBtn.disabled = true;
        status.textContent = '點擊麥克風開始錄音';
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