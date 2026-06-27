# Excel Auditor V2 Starter

This workspace contains a clean starter for the next version of the Excel financial model auditor.

## What is included

- `app.py`: Streamlit interface for upload, summary, issues, and AI-ready Q&A
- `workbook_parser.py`: workbook parsing and basic model classification
- `dependency_graph.py`: precedent/dependent graph construction
- `block_detector.py`: first-pass formula block grouping
- `detectors.py`: initial detectors for hidden pattern breaks and hardcoded issues
- `issue_pipeline.py`: orchestration and dedupe
- `ai_tools.py`: tool functions an AI agent can call
- `chat_agent.py`: lightweight placeholder Q&A layer
- `openai_agent.py`: optional OpenAI-backed grounded chat layer

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

To enable OpenAI-backed chat:

```bash
export OPENAI_API_KEY=your_key_here
streamlit run app.py
```

To protect a deployed private beta with a simple password:

```bash
export APP_PASSWORD=your_demo_password
streamlit run app.py
```

## Why this structure

The goal is to avoid another single-file app and make room for:

1. Stronger formula inconsistency detection
2. Better workbook navigation
3. OpenAI-backed chat and reasoning
4. Data cleanup and analysis workflows later

## Deploy on Render

1. Push this folder to GitHub.
2. Create a new Web Service on Render and point it at the repo.
3. Render will pick up `render.yaml`.
4. Set these environment variables in Render:
   - `OPENAI_API_KEY`
   - `APP_PASSWORD`
5. Deploy and share the private URL with analyst friends.

### Why Render

- easiest private beta path
- simple env var management
- easy redeploys as we keep updating the app
