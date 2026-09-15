@echo off
REM Compileaza aplicatia in .exe, folosind dozare_titan.spec (care include
REM corect folderul assets/logo.png — vezi comentariul din .spec pentru
REM detalii despre de ce e nevoie de el).
REM
REM Ruleaza acest fisier din radacina proiectului (langa main.py), cu
REM dublu-click sau din linia de comanda.

python -m pip install --upgrade pyinstaller
if errorlevel 1 goto eroare

python -m PyInstaller --noconfirm dozare_titan.spec
if errorlevel 1 goto eroare

echo.
echo ============================================================
echo  Gata. Aplicatia compilata e in:
echo    dist\dozare_titan\dozare_titan.exe
echo.
echo  IMPORTANT: muta/copiaza TOT folderul "dist\dozare_titan"
echo  acolo unde vrei sa folosesti aplicatia, nu doar exe-ul —
echo  langa el trebuie sa ramana fisierele lui interne (inclusiv
echo  logo.png, altfel documentele revin la textul "ZIROM TITANIUM").
echo ============================================================
pause
goto sfarsit

:eroare
echo.
echo A aparut o eroare la compilare — vezi mesajele de mai sus.
pause

:sfarsit
