@echo off
cd /d "C:\Users\willi\OneDrive\Desktop\COA_Vault_Project"
start "" streamlit run app.py --server.headless true
timeout /t 3
start chrome --app=http://localhost:8501