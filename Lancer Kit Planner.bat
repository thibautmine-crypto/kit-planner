@echo off
chcp 65001 >nul
title Kit Planner

:: Se placer dans le dossier du script
cd /d "%~dp0"

:: Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  [ERREUR] Python n'est pas installe.
    echo  Lance d'abord "INSTALLER (a faire une fois).bat"
    echo.
    pause
    exit /b 1
)

echo.
echo  Démarrage de Kit Planner...
echo  L'application va s'ouvrir dans votre navigateur dans quelques secondes.
echo.
echo  ─────────────────────────────────────────────
echo   Pour fermer l'app : ferme cette fenetre
echo  ─────────────────────────────────────────────
echo.

:: Ouvrir le navigateur en arrière-plan après 5 secondes (le temps que Streamlit démarre)
start "" cmd /c "timeout /t 5 /nobreak >nul && start http://localhost:8501"

:: Lancer Streamlit (garde la fenêtre ouverte)
streamlit run app.py --server.headless true --browser.gatherUsageStats false
