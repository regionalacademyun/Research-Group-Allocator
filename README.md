# RAUN Project Allocator

A Streamlit dashboard for allocating RAUN participants to research projects using:
- ranked preferences
- topic interest scores
- project capacity rules
- manual review flags

## Files

- `app.py` — main Streamlit application
- `requirements.txt` — Python dependencies for deployment
- `.gitignore` — ignores local and temporary files

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Community Cloud

1. Create a new GitHub repository.
2. Upload these files to the root of the repository.
3. Push to GitHub.
4. Go to Streamlit Community Cloud.
5. Create a new app and connect your GitHub repository.
6. Set the main file path to `app.py`.
7. Deploy.

## Input file expected

Upload a Google Form export in either:
- `.xlsx`
- `.csv`

The app tries to detect the response sheet automatically for Excel files and normalize common column names.

## Notes

- The app keeps project definitions directly in code.
- Output can be downloaded as Excel and CSV.
- Human review is still recommended for edge cases.
