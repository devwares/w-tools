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

# Get configuration object from raw Json content
write-output "Get configuration from raw Json content"
$Config = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)

# Change some parameters
$version = "10"
$Config.SHAREPOINT.TemplateDestinationSiteURL = "https://domain.sharepoint.com/sites/newhybridsite$version"
$Config.SHAREPOINT.GroupsDestinationSiteURL = "https://domain.sharepoint.com/sites/newhybridsite$version"
$Config.SHAREPOINT.StructureDestinationSiteURL = "https://domain.sharepoint.com/sites/newhybridsite$version"
$Config.SHAREPOINT.COMMON.Name = "newhybridsite$version"
$Config.SHAREPOINT.COMMON.Title = "New Hybrid Automation Site $version"

# Export Config to Json format
write-output "Export Config to Json format"
$JsonContent = $Config | ConvertTo-Json -Depth 100

# Create new site
write-output "Create new site"
$test02a = $JsonContent | New-SharepointSite -Credentials $Cred

# Apply Template to new site
write-output "Apply Template to new site"
$test03a = $test02a | Set-SharepointSiteFromTemplate -Credentials $Cred

# Create Groups for new site
write-output "Create Groups for new site"
$test04a = $test03a | Set-SharepointSiteGroups -Credentials $Cred

# Create Structure for new site
write-output "Create Structure for new site"
$test05a = $test04a | Set-SharepointSiteStructure -Credentials $Cred