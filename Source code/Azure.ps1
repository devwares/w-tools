function Find-AzVms() {

    param
    (
        [Parameter(Mandatory=$true)] [array] $SubscriptionList,
        [Parameter(Mandatory=$false)] [string] $ExportFilePath,
        [Parameter(Mandatory=$false)] [string] $Delimiter
    )

    Import-Module Az.Compute
    Import-Module Az.Accounts
    Import-Module Az.Network

    # Initialize returned array
    $report = @()

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

    # Export in a file if specified
    If(-not [string]::IsNullOrEmpty($ExportFilePath)){
        If([string]::IsNullOrEmpty($Delimiter)){
            $report | Export-CSV -path $ExportFilePath
        }
        # Custom delimiter if specified
        Else {
            $report | Export-CSV -path $ExportFilePath -Delimiter $Delimiter
        }

    }

    return $report

}