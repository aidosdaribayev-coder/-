@echo off
chcp 65001 > nul
echo.
echo ══════════════════════════════════════════════════
echo    📊 DATA AGENT - ИИ Аналитик Энергетики
echo ══════════════════════════════════════════════════
echo.

:: Переходим в папку со скриптом
cd /d "%~dp0"

:: Увеличиваем лимит памяти Node.js
set NODE_OPTIONS=--max-old-space-size=8192

:: Установи сюда свой API ключ
set ANTHROPIC_API_KEY=YOUR_API_KEY_HERE


echo 📁 Папка: %cd%
echo 🔧 Память: 8 ГБ
echo.

:: Открываем браузер через 3 секунды
start "" cmd /c "timeout /t 3 /nobreak >nul && start http://localhost:3000"

echo 🚀 Запуск сервера...
echo.
echo    После запуска откроется браузер автоматически
echo    Для остановки нажмите Ctrl+C
echo.

node server.js

pause
