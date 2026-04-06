@echo off
title LLM_Assistant - Streamlit

REM === Adaptation du chemin du projet ===
cd /d "C:\LLM_Assistant"

REM === (Optionnel) activer un environnement virtuel ===
REM call venv\Scripts\activate.bat

REM === Lancer Streamlit ===
streamlit run app.py

pause
