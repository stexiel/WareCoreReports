@echo off
echo Удаление задачи автозапуска WareCore Backend Server...
schtasks /delete /tn "WareCore Backend Server" /f
if %errorlevel% equ 0 (
    echo Успешно! Автозапуск отключен.
) else (
    echo Ошибка удаления. Возможно задача не была создана.
)
pause
