@echo off
title LLM_Assistant - Streamlit

REM === Se placer dans le dossier du projet ===
cd /d "%~dp0"

REM === Vérifier l'environnement virtuel ===
if not exist ".venv\Scripts\python.exe" (
  echo ERREUR: environnement virtuel introuvable : %cd%\.venv\Scripts\python.exe
  pause
  exit /b 1
)

REM === Lancer Streamlit ===
.venv\Scripts\python.exe -m streamlit run app.py

pause