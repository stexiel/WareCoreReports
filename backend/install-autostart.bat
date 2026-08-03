@echo off
echo Регистрация задачи автозапуска WareCore Backend Server...
schtasks /create /tn "WareCore Backend Server" /tr "\"%~dp0start-server.bat\"" /sc onstart /ru "SYSTEM" /rl highest /f
if %errorlevel% equ 0 (
    echo Успешно! Сервер будет запускаться автоматически при старте Windows.
) else (
    echo Ошибка регистрации. Попробуйте запустить от имени администратора.
)
pause
