# Local part ################################

# Credentials (optional)
#$User = "john@domain.net"
#$Pwd = Read-Host -AsSecureString -Prompt "Password for user $User"
#$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $User, $Pwd

$ProjectPath = "C:\Users\John\Projects\w-tools"

# (Re)Import modules
Get-Module "PnpSharepoint" | ForEach-Object {Remove-Module $_}
Import-Module "$ProjectPath\Source Code\PnpSharepoint.ps1"

# Configuration files
$jsonfile = "$ProjectPath\Configuration Files\Sample.json"
$jsoncontent = Get-Content $jsonfile -encoding utf8 -raw

# Variables
$uri = "https://ddfdza10-5ec4-dZ39-819b-e21f28e175fe.webhook.fc.azure-automation.net/webhooks?token=Jn%2fe4m8vTAZDazd32R3dzdoF%2biR5nv%2b9oGTcIxcSAo%3d" # FullAutomation webhook

# Prepare content
$headers = @{ header01="header 01 content" ; header02="header 02 content"}
$body = $jsoncontent

# Send content
$response = Invoke-WebRequest -Method Post -Uri $uri -Body $body -Headers $headers -ContentType "text/plain; charset=utf-8"

# Get Jobid
$jobid = (ConvertFrom-Json ($response.Content)).jobids[0]
Write-Host "Job Id : $jobid"