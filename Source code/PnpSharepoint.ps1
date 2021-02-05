function New-SharepointSiteTemplate()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    Try {

        # Import Modules
		Import-Module "Azure"
        Import-module "sharepointpnppowershellonline"

        # Register Pnp Management Shell (run once on clients)
        #Register-PnPManagementShellAccess

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }

        # Sharepoint source site settings
        $TemplateReferenceSiteURL = $Config.SHAREPOINT.TemplateReferenceSiteURL

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        # Temporary file
        $timestamp = get-date -Format 'yyyyMMddHHmmssffff'
        $TemporaryFile = "spstmpfile" + $timestamp + ".tmp"

        # Connect to site
        Connect-PnPOnline -Url $TemplateReferenceSiteURL -Credentials $Credentials

        # Generate and save template to temporary file
        $encoding = [System.Text.Encoding]::$Config.XML.Encoding
        Get-PnPProvisioningTemplate -Out $TemporaryFile -Encoding $encoding

        # Look for destination media(s) in configuration ############

        # Files marked as active
        foreach ($file in $Config.XML.MEDIAS.FILESYSTEMS.FILES) {

            if ($file.Active -eq 1) {
                
                # Write XML template to $file.FullPath
                Copy-Item $TemporaryFile -destination $file.FullPath

            }

        }

        # Blobs marked as active
        $storageaccounts=$Config.XML.MEDIAS.STORAGEACCOUNTS
        foreach ($storageaccount in $storageaccounts)  { 
            
            foreach ($container in $storageaccount.CONTAINERS) {
        
                foreach ($blob in $container.BLOBS) {
        
                    if ($blob.Active -eq 1) {
        
                        # Get Blob parameters
                        $StorageAccountName = $storageaccount.Name
                        $StorageAccountKey = $storageaccount.Key
                        $ContainerName = $container.Name
                        $BlobName = $blob.Name

                        # Upload to Blob
                        $AzStorageBlobReturn = Set-BlobContent -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey `
                                                            -ContainerName $ContainerName -SourceFile $TemporaryFile -BlobName $BlobName
        
                    }
        
                }
        
            }
        
        }

        # Fill contents marked as active
        $Config.XML.CONTENTS | ForEach-Object {
            If ($_.Active -eq 1) {
                $EmbeddedXmlData = get-content $TemporaryFile -Encoding $Config.XML.Encoding | out-string | ConvertTo-Json -Depth 100
                $_.Data = $EmbeddedXmlData
            }
        }

        # Delete temporary file
        Remove-Item $TemporaryFile

        # Disconnection
        Disconnect-PnPOnline

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        Return $False
    }

    # Return Json configuration
    return $NewJsonContent

}

function Export-SharepointSiteStructure()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    function GetSiteStructure {
        param (
            [switch] $incall,
            [parameter(Mandatory=$True)][string] $path
        )
        # Initialize object
        $return_object = @{}
        $folders_object = @()
        # Get library title
        $DocLibTitle = (Get-PnPList $path).Title
        # Get subfolders list
        $Subfolders = Get-PnPFolderItem -FolderSiteRelativeUrl $path -ItemType Folder
        if (!$incall) {
            # Loop on found folders
            Foreach ($subfolder in $Subfolders) {
                $subfolderpath = $path + "/" + $subfolder.name
                $folders_object += $(GetSiteStructure -incall -path $subfolderpath)
            }
            $return_object = @{
                "DocLibName" = $path
                "DocLibTitle" = $DocLibTitle
                "folders" = $folders_object
            }
            return $return_object
        }
        else
        {
            $folders_object = @()
            $subfolderscount = 0
            # Loop on found folders
            Foreach ($subfolder in $Subfolders) {
                $subfolderscount++
                $subfolderpath = $path + "/" + $subfolder.name
                $folders_object += $(GetSiteStructure -incall -path $subfolderpath)
            }
            $return_object = @{
                "name" = $Path.Split("/")[-1]
            }
            if ($subfolderscount -gt 0) {
                $return_object += @{"folders" = $folders_object}
            }
            return $return_object
        }
    }

    Try {

        # Import Modules
        Import-module "sharepointpnppowershellonline"

        # Register Pnp Management Shell (run once on clients)
        #Register-PnPManagementShellAccess

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }

        # Sharepoint reference site settings
        $StructureDefinitionSiteUrl=$Config.SHAREPOINT.StructureDefinitionSiteUrl

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        # Initialize structure object
        $NewStructureDefinition = @()

        # Connect to site
        Connect-PnPOnline -Url $StructureDefinitionSiteUrl -Credentials $Credentials

        # Get structure for libraries marked as active
        foreach ($library in $Config.DEFINITION.options.structureDefinitionLibraries) {

            if ($library.Active -eq 1) {
                
                # Add library structure to new object
                #$NewStructure += BuildLibraryStructure -Path $library.Name | ConvertFrom-Json
                $NewStructureDefinition += $(GetSiteStructure -Path $library.Name)

            }

        }

        # Integrate new structure to Json content
        if ($Config.DEFINITION.structure) {
            $Config.DEFINITION.structure = $NewStructureDefinition
        }
        else {
            $Config.DEFINITION | Add-Member -MemberType "NoteProperty" -Name "structure" -Value $NewStructureDefinition
        }

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

Return $NewJsonContent

}

function New-SharepointSite()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    Try {

        # Import Modules
        Import-module "sharepointpnppowershellonline"
        
        # Register Pnp Management Shell (run once on clients)
        #Register-PnPManagementShellAccess

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }

        # Sharepoint root site settings
        $RootSiteURL = $Config.SHAREPOINT.RootSiteURL

        # Site common settings
        $NewSiteUI = $Config.SHAREPOINT.COMMON.UI
        $NewSiteName = $Config.SHAREPOINT.COMMON.Name
        $NewSiteTitle = $Config.SHAREPOINT.COMMON.Title
        $NewSiteLCID = $Config.SHAREPOINT.COMMON.LCID
        $NewSiteUrl = "$RootSiteURL/sites/$NewSiteName"
        $IsHub = $Config.SHAREPOINT.COMMON.isHub
        $HubSite = $Config.SHAREPOINT.COMMON.hubSite
        $Sharing = $Config.SHAREPOINT.COMMON.Sharing
        switch (($Sharing).ToUpper()) {
            "DISABLED" { }
            "EXISTINGEXTERNALUSERSHARINGONLY" { }
            "EXTERNALUSERSHARINGONLY" { }
            "EXTERNALUSERANDGUESTSHARING" { }
            Default {Throw "Unknown Sharing mode : $Sharing" }
        }

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){          
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        switch (($NewSiteUI).ToUpper()) {

            "CLASSIC" { 

                # Classic site specific settings
                $NewSiteTemplate = $Config.SHAREPOINT.CLASSIC.Template
                $NewSiteTimezone = $Config.SHAREPOINT.CLASSIC.Timezone
                $NewSiteOwner = $Config.SHAREPOINT.CLASSIC.Owner
                
                # Connection
                $Connection = Connect-PnPOnline -Url $RootSiteURL -Credentials $Credentials

                # Prepare params
                $params = @{
                    "Connection"        = $Connection;
                    "Title"             = $NewSiteTitle;
                    "Template"          = $NewSiteTemplate;
                    "LCID"              = $NewSiteLCID;
                    "Owner"             = $NewSiteOwner;
                    "Url"               = $NewSiteUrl;
                    "TimeZone"          = $NewSiteTimezone;
                    "RemoveDeletedSite" = $true;
                    "Force"             = $true;
                    "Wait"              = $true;
                }

                #Create new Classic site
                New-PnPTenantSite @params

                # Disconnection
                Disconnect-PnPOnline

             }

            "MODERN" {

                # Modern site specific settings
                $NewSiteType = $Config.SHAREPOINT.MODERN.Type
                $NewSiteDescription = $Config.SHAREPOINT.MODERN.Description
                $NewSiteClassification = $Config.SHAREPOINT.MODERN.Classification
                $NewSiteDesign = $Config.SHAREPOINT.MODERN.SiteDesign

                # Connection
                $Connection = Connect-PnPOnline -Url $RootSiteURL -Credentials $Credentials

                # Prepare params
                $params = @{
                    "Type"              = $NewSiteType;
                    "Title"             = $NewSiteTitle;
                    "Url"               = $NewSiteUrl;
                    "Description"       = $NewSiteDescription;
                    "Classification"    = $NewSiteClassification;
                    "SiteDesign"        = $NewSiteDesign;
                    "LCID"              = $NewSiteLCID;
                }

                # Create new Modern site
                $ReturnedUrl = New-PnPSite @params

                # Add owners marked as active
                $Connection = Connect-PnPOnline -Url $NewSiteUrl -Credentials $Credentials
                foreach ($owner in $Config.SHAREPOINT.MODERN.Owners)  {
                    if ($owner.active -eq 1) {
                        Set-PnPSite -Owners $owner.Name
                    }
                }

                # Disconnection
                Disconnect-PnPOnline

            }

            Default {Throw "Unknown user interface type : $NewSiteUI" }

        }

        # Connect to new site to apply other settings
        $Connection = Connect-PnPOnline -Url $NewSiteUrl -Credentials $Credentials

        # Hub site management
        if ($IsHub) {
            if (![string]::IsNullOrEmpty($HubSite)) {Throw "Configuration error : a hubsite cannot be associated with another hubsite"}
            else {Register-PnPHubSite -Site $NewSiteUrl}
        }
        elseif (![string]::IsNullOrEmpty($HubSite)) {
            Add-PnPHubSiteAssociation -Site $NewSiteUrl -HubSite $HubSite
        }

        # Sharing
        Set-PnPSite -Sharing $Sharing

        # Disconnection
        Disconnect-PnPOnline

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $NewJsonContent

}

function Set-SharepointSiteFromTemplate()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    Try {

        # Import Modules
		Import-Module "Azure"
        Import-module "sharepointpnppowershellonline"

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }
        
        # Sharepoint site settings
        $TemplateDestinationSiteURL = $Config.SHAREPOINT.TemplateDestinationSiteURL

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){          
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        # Connect to Sharepoint site
        Connect-PnPOnline -Url $TemplateDestinationSiteURL -Credentials $Credentials

        # Files marked as active
        foreach ($file in $config.XML.MEDIAS.FILESYSTEMS.FILES) {

            if ($file.Active -eq 1) {
                
                # Apply template to Sharepoint site
                Apply-PnPProvisioningTemplate -Path $file.FullPath -ClearNavigation

            }

        }

        # Blobs marked as active
        $storageaccounts=$Config.XML.MEDIAS.STORAGEACCOUNTS
        foreach ($storageaccount in $storageaccounts)  { 
            
            foreach ($container in $storageaccount.CONTAINERS) {
        
                foreach ($blob in $container.BLOBS) {
        
                    if ($blob.Active -eq 1) {
        
                        # Get Blob parameters
                        $StorageAccountName = $storageaccount.Name
                        $StorageAccountKey = $storageaccount.Key
                        $ContainerName = $container.Name
                        $BlobName = $blob.Name

                        # Temporary file
                        $timestamp = get-date -Format 'yyyyMMddHHmmssffff'
                        $TemporaryFile = "spstmpfile" + $timestamp + ".tmp"

                        # Get template file
                        Get-BlobContent -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey `
                                            -ContainerName $ContainerName -DestinationFile $TemporaryFile -BlobName $BlobName

                        # Apply template to Sharepoint site
                        Apply-PnPProvisioningTemplate -Path $TemporaryFile -ClearNavigation

                        # Delete temporary file
                        Remove-Item $TemporaryFile

                    }
        
                }
        
            }
        
        }

        # Xml contents marked as active
        $contents=$Config.XML.CONTENTS
        foreach ($content in $contents)  {

            if ($content.active -eq 1) {

                # Write temporary file
                $timestamp = get-date -Format 'yyyyMMddHHmmssffff'
                $TemporaryFile = "spstmpfile" + $timestamp + ".tmp"
                $content.Data | ConvertFrom-Json | Out-File -Force -Encoding $Config.XML.Encoding $TemporaryFile

                # Apply template to Sharepoint site
                Apply-PnPProvisioningTemplate -Path $TemporaryFile -ClearNavigation

                # Delete temporary file
                Remove-Item $TemporaryFile

            }

        }

        # Disconnection
        Disconnect-PnPOnline

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

Return $NewJsonContent

}

function Set-SharepointSiteGroups()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    function SPOCreateGroup {
        param (
            # URL of target site (not tenant)
            [Parameter(Mandatory = $true)]
            [string]
            $TargetUrl,
            # SPO Group Name
            [Parameter(Mandatory = $true)]
            [string]
            $GroupName,
            # Roles
            [Parameter(Mandatory = $true)]
            [string[]]
            $Roles,
            # Roles
            [Parameter(Mandatory = $true)]
            [string[]]
            $Members,
            # Post creation group checking timeout
            [Parameter(Mandatory = $true)]
            [string]
            $CheckGroupTimeout,
            # Permission checking timeout
            [Parameter(Mandatory = $true)]
            [string]
            $CheckPermissionTimeout,
            # SPO Group Description
            [Parameter(Mandatory = $false)]
            [string]
            $GroupDesc = ""
        )
        try {

            if (!(Get-PnPGroup -Identity $GroupName -ErrorAction SilentlyContinue)) {

                Write-Host -ForegroundColor Cyan "Adding group $GroupName to site $TargetUrl)"
                $newGroup = New-PnPGroup -Title $GroupName -Description $GroupDesc

                # Wait for group to be created before continue, throw error if timeout is reached
                $Timer = 0
                $Timeout = $CheckGroupTimeout
                Write-Host "Checking group creation (timeout set to $Timeout seconds)" -NoNewline
                Do{
                    Start-Sleep 1
                    $GroupExists = get-PnPGroup | where-object { $_.title -eq $GroupName }
                    $Timer++
                    Write-Host "." -NoNewline
                }Until (($Timer -eq $Timeout) -or ($GroupExists))
                If ($Timer -eq $Timeout) {Throw ("Group not found : $GroupName - try again later")}
                Write-Host "OK"

                # Wait for role definition to exist, throw error if timeout is reached
                $Timer = 0
                $Timeout = $CheckPermissionTimeout
                Write-Host "Checking role definition (timeout set to $Timeout seconds)" -NoNewline
                Do{
                    Start-Sleep 1
                    $RoleExists = Get-PnPRoleDefinition | where-object { $_.Name -eq $Roles }
                    $Timer++
                    Write-Host "." -NoNewline
                }Until (($Timer -eq $Timeout) -or ($RoleExists))
                If ($Timer -eq $Timeout) {Throw ("Role definition not found : $Roles - try again later")}
                Write-Host "OK"

                # Set permission level for newly created group
                Set-PnPGroupPermissions -Identity $newGroup -AddRole $Roles
                Write-Host -ForegroundColor Green "Added group $GroupName to site $TargetUrl)"
                Write-Host -ForegroundColor Cyan "Adding members to group $GroupName in site $TargetUrl)"
                foreach ($member in $Members) {
                    try {
                        Add-PnPUserToGroup -Identity $newGroup -LoginName $member
                    }
                    catch {
                        Write-Host -ForegroundColor Cyan "Failed to add member $member to group $GroupName in site $TargetUrl)"
                    }
                }
                Write-Host -ForegroundColor Green "Added members to group $GroupName in site $TargetUrl)"
                return $newGroup
            }
            else {
                Write-Host -ForegroundColor Yellow "Group $GroupName already present in site $TargetUrl)"
            }
        }
        catch {
            Throw $_.Exception.Message
        }
    }

    Try {
        
        # Import Modules
        Import-module "sharepointpnppowershellonline"

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }
        
        # Sharepoint site settings
        $GroupsDestinationSiteURL = $Config.SHAREPOINT.GroupsDestinationSiteURL

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){          
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        # Connect to Sharepoint site
        Connect-PnPOnline -Url $GroupsDestinationSiteURL -Credentials $Credentials

        # Create groups
        foreach ($groupDef in $Config.DEFINITION.groups) {
            SPOCreateGroup -TargetUrl $GroupsDestinationSiteURL -GroupName $groupDef.groupName -GroupDesc $groupDef.groupDescription -Roles $groupDef.groupPermissionLevel -Members $groupDef.groupMembers -CheckGroupTimeout $Config.DEFINITION.CheckGroupTimeout -CheckPermissionTimeout $Config.DEFINITION.CheckPermissionTimeout
        }

        # Disconnection
        Disconnect-PnPOnline

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $NewJsonContent

}

function Set-SharepointSiteStructure()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

    function SPOCreateLibrary {
        param (
            # Title of the library
            [Parameter(Mandatory = $true)]
            [string]
            $Title,
            # Name (url) of the library
            [Parameter(Mandatory = $true)]
            [string]
            $Name
        )
        try {
            Write-Host -ForegroundColor Cyan "Creating library : $Title"
            New-PnPList -Title $Title -Url $Name -Template DocumentLibrary -OnQuickLaunch
            return $Name
        }
        catch {
            Throw $_.Exception.Message
        }
        finally {
        }
    }

    function SPOPopulateLibrary {
        param (
            # folder
            [Parameter(Mandatory = $true)]
            $List,
            # Folders structure
            [Parameter(Mandatory = $true)]
            [PSCustomObject]
            $Folders,
            # Permissions at level library
            [Parameter(Mandatory = $false)]
            [PSCustomObject]
            $Permissions = $null,
            # Parent folder
            [Parameter(Mandatory = $false)]
            [string]
            $ParentFolder = ""
        )
        # apply permissions on library
        if ($Permissions) {
            Write-Host -ForegroundColor Cyan "Applying permissions to library $($List.Title)"
            $permBaseParams = @{
                "Identity"   = $List;
            }
            # reset role inheritance
            if ($Permissions.reset) {
                $resetParams = $permBaseParams.Clone()
                $resetParams.Add("ResetRoleInheritance", $true)
                Set-PnPList @resetParams
            }
            if ($Permissions.members) {
                $permBaseParams = @{
                    "Identity"   = $List;
                }
                # break role inheritance
                $breakParams = $permBaseParams.Clone()
                $breakParams.Add("BreakRoleInheritance", $true)
                Set-PnPList @breakParams
                foreach ($member in $Permissions.members) {
                    $oMember = Get-PnPGroup -Identity $member.memberName -ErrorAction SilentlyContinue
                    if (!$member.memberPermissions) {
                        Write-Host -ForegroundColor Yellow "No permissions defined for Group $($member.memberName) in structure template"
                        continue
                    }
                    if ($oMember) {
                        Write-Host -ForegroundColor Cyan "Applying permissions $($member.memberPermissions) to group $($member.memberName)"
                        $groupPermParams = $permBaseParams.Clone()
                        $groupPermParams.Add("Group", $oMember)
                        foreach ($role in $member.memberPermissions) {
                            $groupPermRoleParams = $groupPermParams.Clone()
                            $groupPermRoleParams.Add("AddRole", $role)
                            Set-PnPListPermission @groupPermRoleParams
                        }
                    }
                    else { Write-Host -ForegroundColor Yellow "Group $($member.memberName) doesn't exists" }
                }
            }
        }
    
        # create folders
        foreach ($folder in $Folders) {
            try {
                $cleared = $false
                $parentRelativeUrl = "$($list.RootFolder.Name)/$ParentFolder"
                Write-Host -ForegroundColor Cyan "Creating folder : $parentRelativeUrl$($folder.name)"
                # create folder
                $oFolder = Resolve-PnPFolder -Connection $Connection -SiteRelativePath "$parentRelativeUrl$($folder.name)"
                # apply permissions on folder
                if ($folder.permissions) {
                    Write-Host -ForegroundColor Cyan "Applying permissions to folder $($folder.name)"
                    $permBaseParams = @{
                        "List"         = $docLib;
                        "Identity"     = $oFolder;
                        "SystemUpdate" = $true;
                    }
                    if ($folder.permissions.reset) {
                        $resetParams = $permBaseParams.Clone()
                        $resetParams.Add("InheritPermissions", $true)
                        Set-PnPFolderPermission @resetParams
                    }
                }
                foreach ($member in $folder.permissions.members) {
                    $oMember = Get-PnPGroup -Identity $member.memberName -ErrorAction SilentlyContinue
                    if (!$member.memberPermissions) {
                        Write-Host -ForegroundColor Yellow "No permissions defined for Group $($member.memberName) in structure template"
                        continue
                    }
                    if ($oMember) {
                        Write-Host -ForegroundColor Cyan "Applying permissions $($member.memberPermissions) to group $($member.memberName)"
                        $groupPermParams = $permBaseParams.Clone()
                        $groupPermParams.Add("Group", $oMember)
                        if (!$cleared -and $folder.permissions.clearExisting) {
                            $groupPermClearParams = $groupPermParams.Clone()
                            $groupPermClearParams.Add("ClearExisting", $true)
                            Set-PnPFolderPermission @groupPermClearParams
                            $cleared = $true
                        }
                        foreach ($role in $member.memberPermissions) {
                            $groupPermRoleParams = $groupPermParams.Clone()
                            $groupPermRoleParams.Add("AddRole", $role)
                            Set-PnPFolderPermission @groupPermRoleParams
                        }
                    }
                    else { Write-Host -ForegroundColor Yellow "Group $($member.memberName) doesn't exists" }
                }
                # Create subfolders
                if ($null -ne $folder.folders) {
                    SPOPopulateLibrary -List $docLib -ParentFolder "$ParentFolder$($folder.name)/" -Folders $folder.folders
                }
            }
            catch {
                Throw $_.Exception.Message
            }
            finally {
            }
        }
    }

    Try {
        
        # Import Modules
        Import-module "sharepointpnppowershellonline"

        # Disable Telemetry (to avoid prompt)
        #Disable-PnPPowerShellTelemetry -force

        # Read config
        If ([string]::IsNullOrEmpty($JsonContent)){
            $Config = Get-Content $ConfigFile | ConvertFrom-Json
        }
        else {
            $Config = $JsonContent | ConvertFrom-Json
        }
        
        # Sharepoint site settings
        $StructureDestinationSiteURL = $Config.SHAREPOINT.StructureDestinationSiteURL

        # If no Credentials parameter, checks Json config for encrypted password
        If (!$Credentials){          
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.User)){
                Throw "Missing username : fix Json content or use -Credentials parameter"
            }
            Else{
                $UserName = $Config.SHAREPOINT.User
                If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
                }
                Else {
                    $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                    $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                    $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
                }
            }
            $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password
        }

        # Connect to Sharepoint site
        Connect-PnPOnline -Url $StructureDestinationSiteURL -Credentials $Credentials

        # Create structure
        foreach ($docLibDef in $Config.DEFINITION.structure) {

            $docLib = Get-PnPList -Identity $docLibDef.docLibName
            if ($null -eq $docLib) {
                if ($null -eq $docLibDef.docLibTitle) {$DocLibTitle = $docLibDef.docLibName}
                $doclibName = SPOCreateLibrary -Title $DocLibTitle -Name $docLibDef.docLibName
                if ($null -ne $doclibName) {
                    $docLib = Get-PnPList -Identity $docLibDef.docLibName
                }
            }
            if (($null -ne $docLib) -and ($null -ne $docLibDef.folders)) {
                SPOPopulateLibrary -List $docLib -Folders $docLibDef.folders -Permissions $docLibDef.permissions
            }

        }

        # Disconnection
        Disconnect-PnPOnline

        # Convert PSCustomObject to Json
        $NewJsonContent = $Config | ConvertTo-Json -Depth 100

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $NewJsonContent

}

function New-SharepointMembers()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True, ParameterSetName = "pipeline")] $JsonContent,
        [parameter(Mandatory=$True, ParameterSetName = "filepath")][string] $ConfigFile,
        [parameter(Mandatory=$False)][System.Management.Automation.PSCredential] $Credentials
    )

}