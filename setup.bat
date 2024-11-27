@echo off
setlocal

:: Téléchargez et installez Python
echo Téléchargement de Python...
curl -o python-installer.exe https://www.python.org/ftp/python/3.9.7/python-3.9.7-amd64.exe

echo Installation de Python...
start /wait python-installer.exe /quiet InstallAllUsers=1 PrependPath=1

:: Vérifiez que Python est installé
python --version
if %errorlevel% neq 0 (
    echo Erreur: Python n'a pas été installé correctement.
    exit /b 1
)

:: Récupérez le chemin de python.exe
for /f "delims=" %%i in ('where python') do set PYTHON_PATH=%%i

:: Écrivez le chemin dans un fichier texte
echo %PYTHON_PATH% > python_path.txt

:: Installez pip
echo Installation de pip...
python -m ensurepip
python -m pip install --upgrade pip

:: Installez les dépendances
echo Installation des dépendances...
python -m pip install tabula-py pandas openpyxl xlsxwriter

echo Installation terminée. Python, pip et les dépendances sont installés.
echo Le chemin de python.exe a été enregistré dans python_path.txt.

endlocal
pause
