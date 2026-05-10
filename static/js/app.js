// MCH Assistant Web Service - JavaScript

// Recording functionality
let mediaRecorder;
let audioChunks = [];
let isRecording = false;
let audioBlob;

const recordBtn = document.getElementById('recordBtn');
const status = document.getElementById('status');
const submitBtn = document.getElementById('submitBtn');
const result = document.getElementById('result');

if (recordBtn) {
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
            status.textContent = 'Recording... Click to stop';
        } else {
            mediaRecorder.stop();
            isRecording = false;
            recordBtn.textContent = '🎤';
            status.textContent = 'Recording complete';
        }
    };
}

if (submitBtn) {
    submitBtn.onclick = async () => {
        if (!audioBlob) {
            alert('Please record audio first');
            return;
        }
        
        submitBtn.textContent = 'Processing...';
        submitBtn.disabled = true;
        
        const reader = new FileReader();
        reader.readAsDataURL(audioBlob);
        reader.onloadend = async () => {
            const base64 = reader.result.split(',')[1];
            const reportType = document.getElementById('reportType')?.value || 'general';
            
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
                result.innerHTML = '<pre style="white-space: pre-wrap;">' + data.report + '</pre>';
            } catch (err) {
                result.style.display = 'block';
                result.innerHTML = 'Error: ' + err.message;
            }
            submitBtn.textContent = 'Generate Report';
            submitBtn.disabled = false;
        };
    };
}