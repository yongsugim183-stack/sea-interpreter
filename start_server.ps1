$log    = "$env:USERPROFILE\sea-interpreter\server_log.txt"
$python = "$env:USERPROFILE\AppData\Local\Programs\Python\Python314\python.exe"

$env:DISABLE_SSL_VERIFY              = "1"
$env:HF_HUB_DISABLE_SSL_VERIFICATION = "1"
$env:CURL_CA_BUNDLE                  = ""
$env:REQUESTS_CA_BUNDLE              = ""
$env:PYTHONIOENCODING                = "utf-8"

Set-Location "$env:USERPROFILE\sea-interpreter"

while ($true) {
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SERVER START" | Add-Content $log -Encoding UTF8
    & $python server.py *>> $log
    "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] SERVER EXIT -- restart in 5s" | Add-Content $log -Encoding UTF8
    Start-Sleep -Seconds 5
}