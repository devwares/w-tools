# TODO : write documentation
# SUGGESTION : create a separate function to test tcp connections

function Enable-Proxy {

    param
    (
        [Parameter(Mandatory=$false)][String]$Scope    
    )

    # Set scope (User if not specified)
    If ([string]::IsNullOrEmpty($Scope)){
        $RegistryScope = "HKCU"
    }
    Else {
        $RegistryScope = Switch ($Scope.ToLower()) {
            "user" {"HKCU"; break}
            "machine" {"HKLM"; break}
            default {"UNKNOWN"; break}
            }
    }
    If ($RegistryScope -eq "UNKNOWN") {
        Write-Error "Unknown scope : $Scope" -ErrorAction:Continue
        return $False
    }
    $RegistryPath = $RegistryScope + ':\Software\Microsoft\Windows\CurrentVersion\Internet Settings'

    # Enable proxy
    Set-ItemProperty -Path $RegistryPath -name ProxyEnable -Value 1

}

function Disable-Proxy {

    param
    (
        [Parameter(Mandatory=$false)][String]$Scope
    )

    # Set scope (User if not specified)
    If ([string]::IsNullOrEmpty($Scope)){
        $RegistryScope = "HKCU"
    }
    Else {
        $RegistryScope = Switch ($Scope.ToLower()) {
            "user" {"HKCU"; break}
            "machine" {"HKLM"; break}
            default {"UNKNOWN"; break}
            }
    }
    If ($RegistryScope -eq "UNKNOWN") {
        Write-Error "Unknown scope : $Scope" -ErrorAction:Continue
        return $False            
    }
    $RegistryPath = $RegistryScope + ':\Software\Microsoft\Windows\CurrentVersion\Internet Settings'

    # Disable proxy
    Set-ItemProperty -Path $RegistryPath -name ProxyEnable -Value 0

}

function Connect-ToProxy {

    param
    (
        [Parameter(Mandatory=$true)][string]$ProxyString, # e.g "http://192.168.0.1:3128"
        [Parameter(Mandatory=$false)][string] $ProxyUser,
        [Parameter(Mandatory=$false)][Security.SecureString]$ProxyPassword
    )

    try {

        $proxyUri = new-object System.Uri($proxyString)

        # Create WebProxy
        [System.Net.WebRequest]::DefaultWebProxy = new-object System.Net.WebProxy ($proxyUri, $true)

        # Use credentials on Proxy if user specified
        if (![string]::IsNullOrEmpty($ProxyUser))
        {
            # Ask for password if not specified
            if (!$ProxyPassword){
                [System.Net.WebRequest]::DefaultWebProxy.Credentials = Get-Credential -UserName $ProxyUser -Message "Proxy Authentication"
            }
            else {
                [System.Net.WebRequest]::DefaultWebProxy.Credentials = New-Object System.Net.NetworkCredential($ProxyUser, $ProxyPassword)
            }
        
        }

    }
    catch
    {
        Write-Error "Connection to proxy failed --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

}

function Set-Proxy {

    param
    (
            [Parameter(Mandatory=$true,ParameterSetName='fill')][string]$ProxyServerName,
            [Parameter(Mandatory=$true,ParameterSetName='fill')][int32]$ProxyServerPort,
            [Parameter(Mandatory=$false,ParameterSetName='fill')][bool]$ProxyDisable,
            [Parameter(Mandatory=$false,ParameterSetName='reset')][bool]$Reset,
            [Parameter(Mandatory=$false,ParameterSetName='fill')][bool]$ProxyTestConnection,
            [Parameter(Mandatory=$false)][string]$Scope
    )
 
    Try{


        If ($Reset){
            $ProxyServerValue = ""
            $ProxyDisable = $true
        }
        else {
            $ProxyServerValue = "$($ProxyServerName):$($ProxyServerPort)"
            # Perform a connection test if specified
            If ($ProxyTestConnection){
                If (!(Test-NetConnection -ComputerName $ProxyServerName -Port $ProxyServerPort).TcpTestSucceeded) {
                    Write-Error -Message "Invalid proxy server address or port:  $($ProxyServerName):$($ProxyServerPort)"
                    return $False
                }
            }
        }
    
        # Set scope (User if not specified)
        If ([string]::IsNullOrEmpty($Scope)){
            $RegistryScope = "HKCU"
        }
        Else {
            $RegistryScope = Switch ($Scope.ToLower()) {
                "user" {"HKCU"; break}
                "machine" {"HKLM"; break}
                default {"UNKNOWN"; break}
                }
        }
        If ($RegistryScope -eq "UNKNOWN") {
            Write-Error "Unknown scope : $Scope" -ErrorAction:Continue
            return $False            
        }

        # Set proxy
        $RegistryPath = $RegistryScope + ':\Software\Microsoft\Windows\CurrentVersion\Internet Settings'
        Set-ItemProperty -Path $RegistryPath -name ProxyServer -Value $ProxyServerValue

        # Enable proxy unless Disabled specified
        If ($ProxyDisable) {Disable-Proxy -Scope $Scope} else {Enable-Proxy -Scope $Scope}

    }
    catch
    {
        Write-Error "Connection to proxy failed --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

}

function Invoke-PsCommandAs {

    param
    (
            [Parameter(Mandatory=$true)][string]$WindowsUserName,
            [Parameter(Mandatory=$true)][securestring]$WindowsUserPassword,
            [Parameter(Mandatory=$true, ValueFromPipeline = $true)][string]$PsCommand,
            [Parameter(Mandatory=$false)][string]$ImportModules
    )

    # Credentials
    $Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $WindowsUserName, $WindowsUserPassword

    # Modules to import if specified
    If (-not [string]::IsNullOrEmpty($ImportModules)){

        $PSImportModuleCommand = ""
        $ImportModulesList = $ImportModules.Split(";");

        ForEach ($Module in $ImportModulesList)
        {
            $PSImportModuleCommand = $PSImportModuleCommand + "Import-Module `'$Module`'" + ";"
        }

        $PsFinalCommand = $PSImportModuleCommand + $PsCommand

    }
    Else
    {
        $PsFinalCommand = $PsCommand
    }
    

    # Run Import-Module + Parameter command
    Start-Process Powershell -ArgumentList $PsFinalCommand -NoNewWindow -credential $Cred 

}

function Set-EncapsulationContextPolicy
{
    REG ADD HKLM\SYSTEM\CurrentControlSet\Services\PolicyAgent /v AssumeUDPEncapsulationContextOnSendRule /t REG_DWORD /d 0x2 /f
}

function New-L2tpPskVpn
{
    param (
        [Parameter(Mandatory=$true)][string]$VpnConName, 
        [Parameter(Mandatory=$true)][string]$VpnServerAddress,    
        [Parameter(Mandatory=$true)][string]$PreSharedKey,
        [Parameter(Mandatory=$false)][PSCustomObject]$DestinationNetworks # if specified, do not route all traffic through VPN
    )

    $VpnConExists = Get-VpnConnection -Name $VpnConName -ErrorAction Ignore
    if ($VpnConExists) {
        # Remove old connection if exists
        Remove-VpnConnection -Name $VpnConName -Force -PassThru
    }

    # Disable persistent command history
    Set-PSReadlineOption -HistorySaveStyle SaveNothing
    # Create VPN connection
    Add-VpnConnection -Name $VpnConName -ServerAddress $VpnServerAddress -L2tpPsk $PreSharedKey -TunnelType L2tp -EncryptionLevel Required -AuthenticationMethod Chap,MSChapv2 -Force -RememberCredential -PassThru
    # Ignore the data encryption warning (data is encrypted in the IPsec tunnel)

    if ($DestinationNetworks)
    {
        # Remove default gateway
        Set-VpnConnection -Name $VpnConName -SplitTunneling $True
        foreach ($DestinationNetwork in $DestinationNetworks)
        {
            # Add route after successul connection
            $RouteToAdd = $DestinationNetwork.Address + '/' + $DestinationNetwork.NetMask
            Add-VpnConnectionRoute -ConnectionName $VpnConName -DestinationPrefix $RouteToAdd
        }
    }

}

Function Get-EmptyFiles
{

    param
    (
        [Parameter(Mandatory=$true)] [string] $Path
    )

    # Initialize
    $report = @()

    # List directory
    Try {
        Get-Childitem -Path $Path -Recurse | foreach-object {
            if(!$_.PSIsContainer -and $_.length -eq 0) {
                # Get properties
                $file = "" | Select-Object Name,FullName
                $file.Name=$_.Name
                $file.FullName=$_.FullName
                # Append
                $report+=$file
            }
        }
    }
    Catch
    {
        write-host -f Red "Error listing directory $Path -->" $_.Exception.Message
        return $false
    }

    return $report

}

function Disable-UserAccessControl {
    Set-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "ConsentPromptBehaviorAdmin" -Value 00000000 -Force
    Write-Host "User Access Control (UAC) has been disabled." -ForegroundColor Green    
}

function Get-BroadcastAddress
{
   
    param
    (
        [Parameter (Mandatory=$true)] $IPAddress,
        [Parameter (Mandatory=$false)] $SubnetMask='255.255.255.0'
    )

    filter Convert-IP2Decimal
    {
        ([IPAddress][String]([IPAddress]$_)).Address
    }


    filter Convert-Decimal2IP
    {
    ([System.Net.IPAddress]$_).IPAddressToString 
    }


    [UInt32]$ip = $IPAddress | Convert-IP2Decimal
    [UInt32]$subnet = $SubnetMask | Convert-IP2Decimal
    [UInt32]$broadcast = $ip -band $subnet 
    $broadcast -bor -bnot $subnet | Convert-Decimal2IP

}

function Send-Wol
{
    param (
        [Parameter(Mandatory=$true)][string]$MacAddress,
        [Parameter(Mandatory=$true)][string]$IpAddress,
        [Parameter(Mandatory=$False)][string]$SubnetMask = "255.255.255.0",
        [Parameter(Mandatory=$false)][string]$Port = 9
    )

    $MacAddress = ($MacAddress.Replace("-","")).Replace(":","")
    $BroadcastAddress = Get-BroadcastAddress -IPAddress $IpAddress -SubnetMask $SubnetMask
    $target = 0,2,4,6,8,10 | % {[convert]::ToByte($MacAddress.Substring($_,2),16)}
    $packet = (,[byte]255 * 6) + ($target * 16)
    $udpclient = New-Object System.Net.Sockets.UdpClient
    $udpclient.Connect($BroadcastAddress,$Port)
    [void]$udpclient.Send($packet, 102)
    Write-Host "$IpAddress $MacAddress $BroadcastAddress"

}