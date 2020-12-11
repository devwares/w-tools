function Find-AzVms() {

    param
    (
        [Parameter(Mandatory=$true)] [array] $SubscriptionList,
        [Parameter(Mandatory=$false)] [string] $ExportFilePath,
        [Parameter(Mandatory=$false)] [string] $ExportFileFormat,
        [Parameter(Mandatory=$false)] [string] $Delimiter
    )

    Import-Module Az.Compute
    Import-Module Az.Accounts
    Import-Module Az.Network

    # Initialize array
    $report = @()

    # Initialize export if ExportFilePath specified
    If(-not [string]::IsNullOrEmpty($ExportFilePath)){
    
        # Default format is csv
        If([string]::IsNullOrEmpty($ExportFileFormat)){$ExportFileFormat = "csv"}
        
        # Check ExportFileFormat and format export
        Switch ($ExportFileFormat.ToLower()) {

            "csv"  {$export = @(); break}

            "json" {$export = @{}; break}

            default {Write-Host "ExportFileFormat : $ExportFileFormat not supported" ; return -1 ; break}

        }
    }

    ForEach ($subscriptionId in $subscriptionList)
    {
        Select-AzSubscription $subscriptionId
        $SubscriptionName = (Select-AzSubscription $subscriptionId).Name

        # Gets list of all Virtual Machines
        $vms = Get-AzVM

        # Gets list of all public IPs
        $publicIps = Get-AzPublicIpAddress

        # Gets list of network interfaces attached to virtual machines
        $nics = Get-AzNetworkInterface | Where-Object { $_.VirtualMachine -NE $null} 

        # Gets number of VMs
        $VmsCounter = 0

        foreach ($nic in $nics) {
            
            # Display progress
            $VmsCounter = $VmsCounter+1
            $PercentComplete = (100/$nics.Count)*$VmsCounter
            $ProgressMessage = "Getting informations for " + $vm.Name + " in " + $SubscriptionName
            Write-Progress -Activity $ProgressMessage -PercentComplete $PercentComplete

            # Get attached Virtual Machine
            $vm = $vms | Where-Object -Property Id -eq $nic.VirtualMachine.id

            # $info will store current VM info
            $info = "" | Select Subscription, VmName, VmSize, ResourceGroupName, Region, VirtualNetwork, Subnet, PrivateIpAddress, PublicIPAddress, OSVersion, OsType

            # Subscription
            $info.Subscription = (Select-AzSubscription $subscriptionId).Name

            # VmName
            $info.VmName = $vm.Name
            
            # VmSize
            $info.VmSize = $vm.HardwareProfile.VmSize

            # ResourceGroupName
            $info.ResourceGroupName = $vm.ResourceGroupName

            # Region
            $info.Region = $vm.Location

            # VirtualNetwork
            $info.VirtualNetwork = $nic.IpConfigurations.subnet.Id.Split("/")[-3]

            # Subnet
            $info.Subnet = $nic.IpConfigurations.subnet.Id.Split("/")[-1]
        
            # Private IP Address
            $info.PrivateIpAddress = $nic.IpConfigurations.PrivateIpAddress

            # NIC's Public IP Address, if exists
            foreach($publicIp in $publicIps) { 
            if($nic.IpConfigurations.id -eq $publicIp.ipconfiguration.Id) {
                $info.PublicIPAddress = $publicIp.ipaddress
                }
            }
        
            # OsVersion
            $info.OsVersion = $vm.StorageProfile.ImageReference.Offer + ' ' + $vm.StorageProfile.ImageReference.Sku

            # OsType
            $info.OsType = $vm.StorageProfile.OsDisk.OsType

            # Append
            $report+=$info

        }

    }

    # Output to a file if ExportFilePath specified
    If(-not [string]::IsNullOrEmpty($ExportFilePath)){

        Switch ($ExportFileFormat.ToLower()) {

        # Export to CSV
        "csv"  {

                    # Default delimiter
                    If([string]::IsNullOrEmpty($Delimiter)){
                        $report | Export-CSV -path $ExportFilePath
                    }
                    # Delimiter specified
                    Else {
                        $report | Export-CSV -path $ExportFilePath -Delimiter $Delimiter
                    }

                    ; break

                }

        # Export to JSON
        "json" {
        
                    # $export = @{}
                    $JsonSubscriptionList = $report | Select-Object -Unique -property subscription
        
                    ForEach ($JsonSubscription in $JsonSubscriptionList) {

                        $JsonVnetList = $report | Where-Object -Property Subscription -eq $JsonSubscription.Subscription | Select-Object -Unique -property VirtualNetwork
                        $export[$JsonSubscription.Subscription] = @{
                            Name = $JsonSubscription.Subscription
                            VirtualNetworks = @{}
                        }

                        ForEach ($JsonVnet in $JsonVnetList) {

                            $JsonSubnetList = $report | Where-Object -Property Subscription -eq $JsonSubscription.Subscription  | Where-Object -Property VirtualNetwork -eq $JsonVnet.VirtualNetwork | Select-Object -Unique -property Subnet
                            $export[$JsonSubscription.Subscription]["VirtualNetworks"][$JsonVnet.VirtualNetwork] = @{
                                Name = $JsonVnet.VirtualNetwork
                                Subnets = @{}
                            }

                            ForEach ($JsonSubnet in $JsonSubnetList) {

                                $JsonVmsList = $report | Where-Object -Property Subscription -eq $JsonSubscription.Subscription  | Where-Object -Property VirtualNetwork -eq $JsonVnet.VirtualNetwork | Where-Object -Property Subnet -eq $JsonSubnet.Subnet | Select-Object -Unique -property VmName,VmSize,ResourceGroupName,Region,PrivateIpAddress,PublicIPAddress,OSVersion,OsType
                                $export[$JsonSubscription.Subscription]["VirtualNetworks"][$JsonVnet.VirtualNetwork]["Subnets"][$JsonSubnet.Subnet] = @{
                                    Name = $JsonSubnet.Subnet
                                    Vms = @{}
                                }

                                ForEach ($JsonVm in $JsonVmsList) {

                                    $export[$JsonSubscription.Subscription]["VirtualNetworks"][$JsonVnet.VirtualNetwork]["Subnets"][$JsonSubnet.Subnet]["Vms"][$JsonVm.VmName] = @{
                                        Name = $JsonVm.VmName
                                        Size = $JsonVm.VmSize
                                        ResourceGroupName = $JsonVm.ResourceGroupName
                                        Region = $JsonVm.Region
                                        PrivateIpAddress = $JsonVm.PrivateIpAddress
                                        PublicIPAddress = $JsonVm.PublicIPAddress
                                        OSVersion = $JsonVm.OSVersion
                                        OsType = $JsonVm.OsType.ToString()
                                    }

                                }

                            }

                        }

                    }

                    $export | ConvertTo-Json -Depth 7| out-file $ExportFilePath

                    ; break

                }

        }

    }

    return $report

}

Function Start-AzureVm()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [string] $VmName,
        [Parameter(Mandatory=$true)] [string] $RegGroupName,
        [Parameter(Mandatory=$true)] [string] $Tenant,
        [Parameter(Mandatory=$true)] [string] $User,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password
    )
 
    Try {

        $Credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $Password
        Connect-AzAccount -Credential $Credential -Tenant $Tenant
        Start-AzVM -ResourceGroupName $RegGroupName -Name $VmName

  }
    Catch {
        write-host -f Red "Error starting Vm -->" $_.Exception.Message
    }
}

Function New-BlobContainer()
{

    param
    (
        [Parameter(Mandatory=$true)] [string] $Region,
        [Parameter(Mandatory=$true)] [string] $ResGroupName,
        [Parameter(Mandatory=$true)] [string] $StorageAccountName,
        [Parameter(Mandatory=$true)] [string] $ContainerName
    )

    Import-Module Az.Accounts

    Try{

        # Create ResourceGroup name if doesn't exist
        Get-AzResourceGroup -Name $ResGroupName -Location $Region -ErrorVariable notPresent -ErrorAction SilentlyContinue | Out-Null
        if ($notPresent)
        {
            Write-Warning "Resource group $ResGroupName does not exist : creating"
            New-AzResourceGroup -Name $ResGroupName -Location $Region
        }
        else
        {
            Write-Host "Found resource group : $ResGroupName"
        }

        # Create Storage Account if doesn't exist
        $StorageAccount = Get-AzStorageAccount -ResourceGroupName $ResGroupName -Name $StorageAccountName -ErrorVariable notPresent -ErrorAction SilentlyContinue
        if ($notPresent)
        {
            Write-Warning "Storage Account $StorageAccountName does not exist : creating"
            $storageAccount = New-AzStorageAccount -ResourceGroupName $ResGroupName `
            -Name $StorageAccountName `
            -SkuName Standard_LRS `
            -Location $Region
        }
        else
        {
            Write-Host "Found Storage Account : $StorageAccountName"
        }

        # Keep storage account Context for blob creation
        $ctx = $storageAccount.Context

        # Create Container if doesn't exist
        $Container = Get-AzStorageContainer -Context $ctx -Name $ContainerName -ErrorVariable notPresent -ErrorAction SilentlyContinue
        if ($notPresent)
        {
            $Container = New-AzStorageContainer -Context  $ctx -Name $ContainerName
        }
        else
        {
            Throw "Container already exists : $ContainerName"
        }
        
    }
    Catch{
        Write-Error "Container not created --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $Container

}
# SUGGESTION : add parameter to allow use of -UseConnectedAccount with New-AzStorageContext for OAuth (Azure AD)
# SUGGESTION : ValueFromPipeline = $true for $SourceFile
Function New-BlobContent()
{

    param
    (
        [Parameter(Mandatory=$true)] [string] $StorageAccountName,
        [Parameter(Mandatory=$true)] [string] $StorageAccountKey,
        [Parameter(Mandatory=$true)] [string] $ContainerName,
        [Parameter(Mandatory=$true)] [string] $SourceFile,
        [Parameter(Mandatory=$true)] [string] $BlobName
    )

    Import-Module Az.Accounts

    Try{

        # Get Context
        $ctx = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

        # Upload
        $Upload = Set-AzStorageBlobContent -Container $ContainerName -File $SourceFile -Blob $BlobName -Context $ctx -Force

    }
    Catch{
        Write-Error "Blob not created --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $Upload

}

Function Get-BlobContent()
{

    param
    (
        [Parameter(Mandatory=$true)] [string] $StorageAccountName,
        [Parameter(Mandatory=$true)] [string] $StorageAccountKey,
        [Parameter(Mandatory=$true)] [string] $ContainerName,
        [Parameter(Mandatory=$true)] [string] $BlobName,
        [Parameter(Mandatory=$true)] [string] $DestinationFile
    )

    Import-Module Az.Accounts

    Try{

        # Get Context
        $ctx = New-AzStorageContext -StorageAccountName $StorageAccountName -StorageAccountKey $StorageAccountKey

        # Download
        $Download = Get-AzStorageBlobContent -Container $ContainerName -Blob $BlobName -Destination $DestinationFile -Context $ctx -Force

    }
    Catch{
        Write-Error "Blob not downloaded --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    Return $Download

}