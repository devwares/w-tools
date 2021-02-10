$ProjectPath = "C:\Users\John\Projects\w-tools"
$ExportPath = "C:\Users\John\Documents"

# Credentials (optional)
#$User = "john@domain.net"
#$Pwd = Read-Host -AsSecureString -Prompt "Password for user $User"
#$Cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $User, $Pwd

# (Re)Import modules
Get-Module "PnpSharepoint" | ForEach-Object {Remove-Module $_}
Import-Module "$ProjectPath\Source Code\PnpSharepoint.ps1"

# Configuration files
$jsonfile = "$ProjectPath\Configuration Files\sample.json"
$jsoncontent = Get-Content $jsonfile -encoding utf8 -raw

# Create Template from ref site
$test01a = $jsoncontent | New-SharepointSiteTemplate #-Credentials $Cred
$outfile = "$ExportPath\test01a.json"
$test01a | Out-File -Encoding "utf8" $outfile

# Get structure from ref site
$test01b = $test01a | Export-SharepointSiteStructure #-Credentials $Cred
$outfile = "$ExportPath\test01b.json"
$test01b | Out-File -Encoding "utf8" $outfile

# Change some parameters
$Config = (ConvertFrom-Json -InputObject $test01b)
$version = "09"
$Config.SHAREPOINT.TemplateDestinationSiteURL = "https://domain.sharepoint.com/sites/newlocalsite$version"
$Config.SHAREPOINT.GroupsDestinationSiteURL = "https://domain.sharepoint.com/sites/newlocalsite$version"
$Config.SHAREPOINT.StructureDestinationSiteURL = "https://domain.sharepoint.com/sites/newlocalsite$version"
$Config.SHAREPOINT.COMMON.Name = "newlocalsite$version"
$Config.SHAREPOINT.COMMON.Title = "New Local Site $version"
$JsonContent = $Config | ConvertTo-Json -Depth 100

# Create new site
$test02a = $JsonContent | New-SharepointSite #-Credentials $Cred
$outfile = "$ExportPath\test02a.json"
$test02a | Out-File -Encoding "utf8" $outfile

# Apply Template to new site
$test03a = $test02a | Set-SharepointSiteFromTemplate #-Credentials $Cred
$outfile = "$ExportPath\test03a.json"
$test03a | Out-File -Encoding "utf8" $outfile

# Create Groups for new site
$test04a = $test03a | Set-SharepointSiteGroups #-Credentials $Cred
$outfile = "$ExportPath\test04a.json"
$test04a | Out-File -Encoding "utf8" $outfile

# Create Structure for new site
$test05a = $test04a | Set-SharepointSiteStructure #-Credentials $Cred
$outfile = "$ExportPath\test05a.json"
$test05a | Out-File -Encoding "utf8" $outfile