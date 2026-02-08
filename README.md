# ULavalMailer_v2

## Setup (local)

1. Create a virtual environment:
   - python -m venv .venv
   - .\.venv\Scripts\Activate.ps1

2. Install dependencies:
   - pip install -r requirements.txt

3. Gmail OAuth files:
   - Put credentials.json in secrets/credentials.json
   - The app will create secrets/token.json after the first login

## Run
- python app.py

## Notes
- Real logs and secrets are ignored by Git.
- Sample CSVs are provided under 	est_data/*.SAMPLE.csv.