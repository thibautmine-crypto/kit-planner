@echo off
chcp 65001 >nul
title Kit Planner

:: Se placer dans le dossier du script (fonctionne peu importe l'emplacement)
cd /d "%~dp0"

:: Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERREUR] Python n'est pas installe.
    echo  Lance d'abord "INSTALLER (à faire une fois).bat"
    echo.
    pause
    exit /b 1
)

echo.
echo  Démarrage de Kit Planner...
echo  L'application va s'ouvrir dans votre navigateur.
echo.
echo  ─────────────────────────────────────────────
echo   Pour fermer l'app : ferme cette fenêtre
echo  ─────────────────────────────────────────────
echo.

:: Ouvrir le navigateur après 2 secondes
start "" timeout /t 2 >nul
start "" "http://localhost:8501"

:: Lancer Streamlit
streamlit run app.py --server.headless true --browser.gatherUsageStats false
