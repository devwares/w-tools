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


