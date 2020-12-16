# TODO : add other destination types for template file
function New-SharepointSiteTemplateFile()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True)][string] $ConfigFile,
        [parameter(Mandatory=$False)][securestring] $Password
    )

    Try {

        # Import Modules
        Import-Module "Aezan"
        Import-module "sharepointpnppowershellonline"

        # Autoriser Pnp Management Shell (run once on clients)
        #Register-PnPManagementShellAccess

        # Read config
        $Config = Get-Content $ConfigFile | ConvertFrom-Json

        # Sharepoint source site settings
        $SiteURL = $Config.SHAREPOINT.SiteURL
        $UserName = $Config.SHAREPOINT.User

        If (!$Password){
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
            }
            else {
                $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
            }
        }

        # Storage account settings
        $StorageAccountName = $Config.STORAGEACCOUNT.Name
        $StorageAccountKey = $Config.STORAGEACCOUNT.Key
        $ContainerName = $Config.STORAGEACCOUNT.CONTAINER.Name
        $BlobName = $Config.STORAGEACCOUNT.CONTAINER.BLOB.Name

        # Temporary file
        $length = 6
        $characters = "1234567890"
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
        $private:ofs=""
        $TemporaryFile = "$env:temp\spstmpfile_" + [String]$characters[$random] + ".tmp"

        # Credentials
        $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password

        # Connect to site
        Connect-PnPOnline -Url $SiteURL -Credentials $Credentials

        # Generate and save template
        Get-PnPProvisioningTemplate -Out $TemporaryFile
        $AzStorageBlobReturn = Set-BlobContent -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey `
                            -ContainerName $ContainerName -SourceFile $TemporaryFile -BlobName $BlobName

        # Delete temporary file
        Remove-Item $TemporaryFile

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    #Return AzureStorageBlob object
    Return $AzStorageBlobReturn

}

function New-SharepointSite()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True)][string] $ConfigFile,
        [parameter(Mandatory=$False)][securestring] $Password
    )

    Try {

        # Import Modules
        Import-Module "Aezan"
        Import-module "sharepointpnppowershellonline"
        
        # Autoriser Pnp Management Shell (run once on clients)
        #Register-PnPManagementShellAccess

        # Read config
        #$Config = Get-Content "C:\Users\WilliamVilleger\OneDrive - Aezan\Clients\KPMG\Pnp-powershell\Configuration Files\New-SharepointSite-02.json" | ConvertFrom-Json
        $Config = Get-Content $ConfigFile | ConvertFrom-Json

        # Sharepoint source site settings
        $RootSiteURL = $Config.SHAREPOINT.RootSiteURL
        $NewSiteName = $Config.SHAREPOINT.NewSiteName
        $NewSiteType = $Config.SHAREPOINT.NewSiteType
        $NewSiteTitle = $Config.SHAREPOINT.NewSiteTitle
        $NewSiteDescription = $Config.SHAREPOINT.NewSiteDescription
        $NewSiteClassification = $Config.SHAREPOINT.NewSiteClassification
        $NewSiteDesign = $Config.SHAREPOINT.NewSiteDesign
        $UserName = $Config.SHAREPOINT.User
        If (!$Password){
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
            }
            else {
                $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
            }
        }

        # Credentials
        $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password

        # Connect to site
        Connect-PnPOnline -Url $RootSiteURL -Credentials $Credentials

        # Create new site
        $NewSiteUrl = New-PnPSite -Type $NewSiteType -Title $NewSiteTitle `
                            -Url "$RootSiteURL/sites/$NewSiteName" `
                            -Description $NewSiteDescription `
                            -Classification $NewSiteClassification `
                            -SiteDesign $NewSiteDesign

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $NewSiteUrl

}

# TODO : add other source types for template file
function Set-SharepointSiteFromTemplate()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True)][string] $ConfigFile,
        [parameter(Mandatory=$False)][securestring] $Password
    )

    Try {

        # Import Modules
        Import-Module "Aezan"
        Import-module "sharepointpnppowershellonline"

        # Read config
        # $ConfigFile = "C:\Users\WilliamVilleger\OneDrive - Aezan\Clients\KPMG\Pnp-powershell\Configuration Files\Set-SharepointSiteFromTemplate-03.json"
        $Config = Get-Content $ConfigFile | ConvertFrom-Json

        # Sharepoint source site settings
        $SiteURL = $Config.SHAREPOINT.SiteURL
        $UserName = $Config.SHAREPOINT.User

        If (!$Password){
            If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                    $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
            }
            else {
                $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
                $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
                $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
            }
        }

        # Storage account settings
        $StorageAccountName = $Config.STORAGEACCOUNT.Name
        $StorageAccountKey = $Config.STORAGEACCOUNT.Key
        $ContainerName = $Config.STORAGEACCOUNT.CONTAINER.Name
        $BlobName = $Config.STORAGEACCOUNT.CONTAINER.BLOB.Name

        # Temporary file
        $length = 6
        $characters = "1234567890"
        $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
        $private:ofs=""
        $TemporaryFile = "$env:temp\spstmpfile_" + [String]$characters[$random] + ".tmp"

        # Get template file
        Get-BlobContent -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey `
                            -ContainerName $ContainerName -DestinationFile $TemporaryFile -BlobName $BlobName

        # Credentials
        $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password

        # Connect to new site
        Connect-PnPOnline -Url $SiteURL -Credentials $Credentials

        # Apply template to new site
        $Toto = Apply-PnPProvisioningTemplate -Path $TemporaryFile

        # Delete temporary file
        Remove-Item $TemporaryFile

    }
    Catch{
        Write-Error "Error  --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

Return $Toto

}

# NOTE : REMOVE THIS FUNCTION ? MODIFY/DUPLICATE IT TO CALL THE THREE OTHER ONES ?
function New-SharepointSiteFromTemplate()
{
    Param(
        [parameter(Mandatory=$True, ValueFromPipeline = $True)][string] $ConfigFile,
        [parameter(Mandatory=$False)][securestring] $Password
    )

    # Import Modules
    Import-Module "Aezan"
    Import-module "sharepointpnppowershellonline"
    
    # Autoriser Pnp Management Shell (run once on clients)
    #Register-PnPManagementShellAccess

    # Read config
    #$Config = Get-Content "C:\Users\WilliamVilleger\OneDrive - Aezan\Clients\KPMG\Pnp-powershell\Configuration Files\New-SharepointSiteFromTemplate-02.json" | ConvertFrom-Json
    $Config = Get-Content $ConfigFile | ConvertFrom-Json

    # Sharepoint source site settings
    $RootSiteURL = $Config.SHAREPOINT.RootSiteURL
    $NewSiteName = $Config.SHAREPOINT.NewSiteName
    $NewSiteType = $Config.SHAREPOINT.NewSiteType
    $NewSiteTitle = $Config.SHAREPOINT.NewSiteTitle
    $NewSiteDescription = $Config.SHAREPOINT.NewSiteDescription
    $NewSiteClassification = $Config.SHAREPOINT.NewSiteClassification
    $NewSiteDesign = $Config.SHAREPOINT.NewSiteDesign
    $UserName = $Config.SHAREPOINT.User
    If (!$Password){
        If([string]::IsNullOrEmpty($Config.SHAREPOINT.EncryptedPassword)){
                $Password = Read-Host -AsSecureString -Prompt "Password for user $UserName"
        }
        else {
            $Key = Get-Content $Config.SHAREPOINT.EncryptionKeyFile
            $EncryptedPassword = $Config.SHAREPOINT.EncryptedPassword
            $Password = $EncryptedPassword | ConvertTo-SecureString -Key $Key
        }
    }

    # Storage account settings
    $StorageAccountName = $Config.STORAGEACCOUNT.Name
    $StorageAccountKey = $Config.STORAGEACCOUNT.Key
    $ContainerName = $Config.STORAGEACCOUNT.CONTAINER.Name
    $BlobName = $Config.STORAGEACCOUNT.CONTAINER.BLOB.Name

    # Temporary file
    $length = 6
    $characters = "1234567890"
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    $TemporaryFile = "$env:temp\spstmpfile_" + [String]$characters[$random] + ".tmp"

    # Get template file
    Get-BlobContent -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey `
                        -ContainerName $ContainerName -DestinationFile $TemporaryFile -BlobName $BlobName  

    # Credentials
    $Credentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $UserName, $Password

    # Connect to site
    Connect-PnPOnline -Url $RootSiteURL -Credentials $Credentials

    # Create new site
    $NewSiteUrl = New-PnPSite -Type $NewSiteType -Title $NewSiteTitle `
                        -Url "$RootSiteURL/sites/$NewSiteName" `
                        -Description $NewSiteDescription `
                        -Classification $NewSiteClassification `
                        -SiteDesign $NewSiteDesign
     


}

