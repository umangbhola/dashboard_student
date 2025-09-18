# Dashboard for Form Responses

This Streamlit app reads the Excel file's "Form Responses 1" sheet and renders an interactive dashboard to analyze answers per question.

## Quick start

1. Create and activate a virtual environment (recommended).
2. Install dependencies:

```bash
pip install -r requirements.txt
```

3. Run the app (from the project directory):

```bash
streamlit run app.py
```

The app auto-loads the included Excel file:

- `Program Manager - Fortnightly school audit checklist - HUHT project, June-Dec 2025 (Responses).xlsx`
- Sheet: `Form Responses 1`

You can change the file/sheet from the sidebar.

## Notes
- Requires Python 3.9+
- Large files will be cached to speed up repeated loads.

## Deploy

### Render (recommended)
- Connect repo on Render, create a Web Service
- Runtime: Python
- Build Command: `pip install -r requirements.txt`
- Start Command: `streamlit run app.py --server.port $PORT --server.address 0.0.0.0`
- Or use `render.yaml` in this repo for autodetect

### Heroku
- Install Heroku CLI and login
- Ensure `Procfile` exists
- Deploy:
```bash
heroku create your-app-name
heroku stack:set container -a your-app-name
git push heroku main
```
