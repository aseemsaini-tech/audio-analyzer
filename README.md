# Audio Transcription & Analysis App

Transcribe audio with **AssemblyAI**, analyze with **Claude**, and export results as DOCX + CSV.

## Local Setup

```bash
pip install -r requirements.txt
export ASSEMBLYAI_API_KEY=your_assemblyai_key
export ANTHROPIC_API_KEY=your_anthropic_key
python app.py
```

Open **http://localhost:5000**

---

## Deploy to Render (Internet)

### Step 1: Push to GitHub

```bash
cd audio-analyzer
git init
git add .
git commit -m "Initial commit"
gh repo create audio-analyzer --public --source=. --push
```

### Step 2: Deploy on Render

1. Go to **https://dashboard.render.com**
2. Click **New** > **Web Service**
3. Connect your **audio-analyzer** GitHub repo
4. Render auto-detects `render.yaml` — settings are pre-configured
5. Add your **Environment Variables**:
   - `ASSEMBLYAI_API_KEY` = your AssemblyAI key
   - `ANTHROPIC_API_KEY` = your Anthropic key
6. Click **Create Web Service**
7. Wait 2-3 minutes for build and deploy
8. Your app is live at `https://audio-analyzer-xxxx.onrender.com`

### Alternative: One-Click Deploy

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy)

---

## Usage

1. **Upload** an audio file (MP3, WAV, M4A, etc.)
2. Click **Transcribe Audio** — AssemblyAI transcribes with speaker detection
3. **Paste your analysis prompt** — define what attributes to extract
4. Click **Run Analysis** — Claude analyzes and returns structured attributes
5. **Download** the DOCX transcript and/or CSV attributes

## Example Analysis Prompt

```
Analyze this transcript and extract:
- sentiment (positive / negative / neutral)
- main_topic (1-3 words)
- key_points (comma-separated)
- action_items (comma-separated)
- urgency (low / medium / high)
- summary (2-3 sentences)
```

## Getting API Keys

- **AssemblyAI**: Sign up at https://www.assemblyai.com — free tier includes hours of transcription
- **Anthropic**: Get a key at https://console.anthropic.com
