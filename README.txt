DAILY REPORT AUTOMATION - COMPLETE PROJECT

Files:
- app.py                         Streamlit application
- Auto_Pivot_Live_Template_v10.xlsx   Native Excel Pivot template
- requirements.txt               Python dependencies
- .streamlit/secrets.toml.example Password configuration example
- run.bat                        Windows local start command

Password:
Set APP_PASSWORD in Streamlit Secrets or environment variable.
Do not hard-code the password in app.py.

Run on Windows:
1. py -3.12 -m pip install -r requirements.txt
2. Copy .streamlit\\secrets.toml.example to .streamlit\\secrets.toml
3. Change APP_PASSWORD
4. py -3.12 -m streamlit run app.py

UI:
Dark Navy/Black Navy header; White/Light Corporate body.

Pivot:
The app uses the supplied native Pivot template and does not call LibreOffice to rewrite the final workbook.
Pivot sheets: R1 V, R1S V, R2 V, R6 V, R1 OCV, R6 OCV.
