param
(
    [Parameter (Mandatory = $true)]
    [object] $WebhookData
)

# Credentials
write-output "Credentials"
$Cred = Get-AutomationPSCredential -Name "SPOAdmin"

# (Re)Import modules
write-output "(Re)Import modules"
Import-Module PnpSharepoint

# Json content
$jsoncontent = $WebhookData.RequestBody

# Change some parameters
$Config = (ConvertFrom-Json -InputObject $jsoncontent)
$version = "12"
$Config.SHAREPOINT.TemplateDestinationSiteURL = "https://domain.sharepoint.com/sites/fullautomationsite$version"
$Config.SHAREPOINT.GroupsDestinationSiteURL = "https://domain.sharepoint.com/sites/fullautomationsite$version"
$Config.SHAREPOINT.StructureDestinationSiteURL = "https://domain.sharepoint.com/sites/fullautomationsite$version"
$Config.SHAREPOINT.COMMON.Name = "fullautomationsite$version"
$Config.SHAREPOINT.COMMON.Title = "New Full Automation Site $version"
$JsonContent = $Config | ConvertTo-Json -Depth 100

# Create Template from ref site
write-output "Create Template from ref site"
$test01a = $JsonContent | New-SharepointSiteTemplate -Credentials $Cred

# Get structure from ref site
write-output "Get structure from ref site"
$test01b = $test01a | Export-SharepointSiteStructure -Credentials $Cred

# Create new site
write-output "Create new site"
$test02a = $test01b | New-SharepointSite -Credentials $Cred

# Apply Template to new site
write-output "Apply Template to new site"
$test03a = $test02a | Set-SharepointSiteFromTemplate -Credentials $Cred

# Create Groups for new site
write-output "Create Groups for new site"
$test04a = $test03a | Set-SharepointSiteGroups -Credentials $Cred

# Create Structure for new site
write-output "Create Structure for new site"
$test05a = $test04a | Set-SharepointSiteStructure -Credentials $Cred