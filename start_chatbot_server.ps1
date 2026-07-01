$log    = "$env:USERPROFILE\sea-interpreter\chatbot_log.txt"
$python = "$env:USERPROFILE\AppData\Local\Programs\Python\Python314\python.exe"

$env:PYTHONIOENCODING = "utf-8"
$env:CHATBOT_PORT      = "8001"

$secrets = "$env:USERPROFILE\sea-interpreter\secrets.local.ps1"
if (Test-Path $secrets) { . $secrets }

Set-Location "$env:USERPROFILE\sea-interpreter"

while ($true) {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] CHATBOT SERVER START" | Add-Content $log -Encoding UTF8
    & $python chatbot_server.py *>> $log
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] CHATBOT SERVER EXIT -- restart in 5s" | Add-Content $log -Encoding UTF8
    Start-Sleep -Seconds 5
}
