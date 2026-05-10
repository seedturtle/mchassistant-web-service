# Hermes Web Service

A web-based service that connects to Hermes AI Agent for voice-to-report functionality.

## Features

- 🎙️ Voice recording and transcription
- 🤖 AI-powered report generation via Hermes Agent
- 📊 Multiple report templates
- 🔐 Simple user authentication

## Quick Start

### Local Development

```bash
# Install dependencies
pip install -r requirements.txt

# Set environment variables
export GITHUB_TOKEN="your-github-token"
export HERMES_API_KEY="your-hermes-api-key"

# Run the application
uvicorn main:app --reload --port 8000
```

### Zeabur Deployment

1. Fork this repository to GitHub
2. Connect to Zeabur
3. Add environment variables:
   - `GITHUB_TOKEN`: GitHub Personal Access Token
   - `HERMES_API_KEY`: Your Hermes API key
4. Deploy!

Your service will be available at: **https://mchassistantwebservice.zeabur.app**

## API Endpoints

- `POST /api/auth/login` - User login
- `POST /api/auth/register` - User registration
- `POST /api/audio/transcribe` - Upload audio for transcription
- `POST /api/reports/generate` - Generate report from text
- `GET /api/reports/{id}` - Get report by ID

## Project Structure

```
hermes-web-service/
├── main.py              # FastAPI application entry
├── requirements.txt     # Python dependencies
├── templates/           # HTML templates
│   ├── login.html
│   ├── dashboard.html
│   └── record.html
├── services/
│   ├── auth.py          # Authentication service
│   ├── stt.py           # Speech-to-text service
│   ├── hermes.py        # Hermes Agent integration
│   └── reports.py       # Report generation
├── static/
│   ├── css/
│   └── js/
└── README.md
```

## License

MIT