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
