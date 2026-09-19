@echo off
chcp 936 >nul
echo 正在放行 Windows 防火墙 TCP 8000 端口（允许手机/局域网访问）...
netsh advfirewall firewall add rule name="StatsServer8000" dir=in action=allow protocol=TCP localport=8000 profile=private,domain
if %errorlevel%==0 (echo [成功] 已放行！手机连同一Wi-Fi后可访问本机) else (echo [失败] 请确认是以管理员身份运行的)
pause
