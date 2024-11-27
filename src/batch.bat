@echo off

:: Lire le chemin de python.exe depuis le fichier python_path.txt
set /p PYTHON_PATH=<..\python_path.txt

:: Exécuter le script Python avec le chemin récupéré
start "" "%PYTHON_PATH%" ".\PDF2EXCEL_ceta.py"


start "" "C:\Python312\python.exe" ".\PDF2EXCEL_ceta.py"
