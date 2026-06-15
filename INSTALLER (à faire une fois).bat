@echo off
chcp 65001 >nul
title Kit Planner — Installation

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║        KIT PLANNER — INSTALLATION        ║
echo  ╚══════════════════════════════════════════╝
echo.

:: Vérifier Python
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] Python n'est pas installe sur ce PC.
    echo.
    echo  Télécharge Python ici : https://www.python.org/downloads/
    echo  Coche bien "Add Python to PATH" lors de l'installation.
    echo.
    pause
    exit /b 1
)

echo  [OK] Python détecté.
echo.
echo  Installation des modules requis...
echo.

pip install streamlit pandas openpyxl --quiet

if errorlevel 1 (
    echo.
    echo  [ERREUR] L'installation a échoué.
    pause
    exit /b 1
)

echo.
echo  ╔══════════════════════════════════════════╗
echo  ║   Installation terminée avec succès !    ║
echo  ║                                          ║
echo  ║   Tu peux maintenant lancer l'app avec   ║
echo  ║   "Lancer Kit Planner.bat"               ║
echo  ╚══════════════════════════════════════════╝
echo.
pause
