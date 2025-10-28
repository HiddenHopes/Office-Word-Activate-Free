Go and grab your pc or laptop and open cmd with administrator privileges.

Firstly run this code (Just copy and paste it into the command prompt, not the terminal)

cd /d %ProgramFiles%\Microsoft Office\Office16

Then have to convert your license to a volume one,

for /f %x in ('dir /b ..\root\Licenses16\ProPlus2019VL*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x" 
Then run these codes one by one,

cscript ospp.vbs /setprt:1688 
cscript ospp.vbs /unpkey:6MWKP >nul 
cscript ospp.vbs /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP 
cscript ospp.vbs /sethst:e8.us.to 
cscript ospp.vbs /act 
— Note —

If you see the error 0xC004F074, it means that your internet connection is unstable or the server is busy. Please make sure your device is online and try the command “act” again until you succeed.

Please help me by upvoting this answer

Thank you and happy typing!!!!^__^

Reference: https://www.quora.com/How-do-I-activate-Microsoft-Office-2019-Pro-Plus-for-free-on-Windows-11
