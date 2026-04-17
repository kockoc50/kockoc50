@echo off
chcp 65001 >nul
echo.
echo ================================
echo   Claude for Word - התקנה
echo ================================
echo.

:: Check Node.js
node --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [שגיאה] Node.js לא מותקן.
    echo פתח את הכתובת הזו והורד את גרסת LTS:
    echo https://nodejs.org
    echo.
    pause
    exit /b 1
)

echo [1/3] Node.js נמצא - מתקין חבילות...
call npm install
if %ERRORLEVEL% NEQ 0 (
    echo [שגיאה] ההתקנה נכשלה.
    pause
    exit /b 1
)

echo.
echo [2/3] יוצר תעודות SSL...
call npx office-addin-dev-certs install
if %ERRORLEVEL% NEQ 0 (
    echo [אזהרה] לא הצלחנו ליצור תעודות, ממשיכים...
)

echo.
echo [3/3] מפעיל שרת...
echo.
echo ============================================
echo  השרת פועל! עכשיו פתחי את Word ובצעי:
echo.
echo  Insert ^> Add-ins ^> Upload My Add-in
echo  ובחרי את הקובץ: manifest.xml
echo ============================================
echo.
echo כדי לעצור את השרת - לחצי Ctrl+C
echo.

npx http-server . -p 3000 --ssl --cert "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" --key "%USERPROFILE%\.office-addin-dev-certs\localhost.key" --cors
