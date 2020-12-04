function Test-Connect-ToProxy {

    $url = "http://192.168.0.1:8080"
    $username = "john"
    $password = Read-Host "Enter password for $username" -AsSecureString

    Connect-ToProxy -ProxyString $url
    Connect-ToProxy -ProxyString $url -ProxyUser $username
    Connect-ToProxy -ProxyString $url -ProxyUser $username -ProxyPassword $password

}

function Test-Set-Proxy{

    $test1 = Set-Proxy -ProxyServerName "192.168.0.1" -ProxyServerPort 3128 -ProxyTestConnection $true
    $test2 = Set-Proxy -ProxyServerName "proxy.domain.com" -ProxyServerPort 8080 -ProxyDisable $true
    $test3 = Set-Proxy -ProxyServerName "proxy.domain.net" -ProxyServerPort 8888 -ProxyDisable $false
    
}