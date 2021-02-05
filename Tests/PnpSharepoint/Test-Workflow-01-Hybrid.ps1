$ProjectPath = "C:\Users\John\Projects\w-tools"
$ExportPath = "C:\Users\John\Documents"

# Local part ################################

# Credentials (optional)
#$User = "john@domain.net"
#$Pwd = Read-Host -AsSecureString -Prompt "Password for user $User"
#$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $User, $Pwd

# (Re)Import modules
Get-Module "PnpSharepoint" | ForEach-Object {Remove-Module $_}
Import-Module "$ProjectPath\Source Code\PnpSharepoint.ps1"

# Configuration files
$jsonfile = "$ProjectPath\Configuration Files\Workflow-01.json"
$jsoncontent = Get-Content $jsonfile -encoding utf8 -raw

# Create Template from ref site
$test01a = $jsoncontent | New-SharepointSiteTemplate #-Credentials $Cred
$outfile = "$ExportPath\test01a.json"
$test01a | Out-File -Encoding "utf8" $outfile

# Get structure from ref site
$test01b = $test01a | Export-SharepointSiteStructure #-Credentials $Cred
$outfile = "$ExportPath\test01b.json"
$test01b | Out-File -Encoding "utf8" $outfile

# Automation part ################################

# Variables
$uri = "https://dfzf510-5ec4-49a9-819b-edrfff33e175fe.webhook.fc.azure-automation.net/webhooks?token=iG3dzDZ6DAdadAZLUzTEQotP3buUdazDDZyoZs%3d" # HybridAutomation Webhook

# Prepare content
$headers = @{ header01="header 01 content" ; header02="header 02 content"}
$body = $test01b

# ESend content
$response = Invoke-WebRequest -Method Post -Uri $uri -Body $body -Headers $headers -ContentType "text/plain; charset=utf-8"

# Get job id
$jobid = (ConvertFrom-Json ($response.Content)).jobids[0]
Write-Host "Job Id : $jobid"
