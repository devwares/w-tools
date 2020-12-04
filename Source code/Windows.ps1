# TODO : write documentation
function Enable-Proxy {
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -Value 1
}

function Disable-Proxy {
    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyEnable -Value 0
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

# SUGGESTION : add "-scope" option to set either system-wide proxy or user proxy
function Set-Proxy {

    param
    (
            [Parameter(Mandatory=$true)][string]$ProxyServerName,
            [Parameter(Mandatory=$true)][int32]$ProxyServerPort,
            [Parameter(Mandatory=$false)][bool]$ProxyDisable,
            [Parameter(Mandatory=$false)][bool]$ProxyTestConnection
    )

    Try{

        # Perform a connection test if specified
        If ($ProxyTestConnection){
            If (!(Test-NetConnection -ComputerName $ProxyServerName -Port $ProxyServerPort).TcpTestSucceeded) {
                Write-Error -Message "Invalid proxy server address or port:  $($ProxyServerName):$($ProxyServerPort)"
                return $False
            }
        }

        # Set proxy
        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings' -name ProxyServer -Value "$($ProxyServerName):$($ProxyServerPort)"

        # Enable proxy unless Disabled specified
        if ($ProxyDisable) {Disable-Proxy} else {Enable-Proxy}

    }
    catch
    {
        Write-Error "Connection to proxy failed --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

}


