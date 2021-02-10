param
(
    [Parameter(Mandatory=$true)] [string] $name,
    [Parameter(Mandatory=$true)] [string] $address
)

$ProjectPath = "C:\Users\John\Projects\Pnp-Powershell"
$uri = "https://s17events.azure-automation.net/webhooks?token=8J7Gg9%3DFffeEEd34dsSSJyhy66tRffd4tg7OjquyPQA%3t"

#$jsoncontent = get-content "$ProjectPath\Configuration Files\sample.json" -encoding utf8 -raw
$jsoncontent  = @(
            @{ Name="$name";Address="$address"}
        )
$body = ConvertTo-Json -InputObject $jsoncontent
$headers = @{ header01="header 01 content" ; header02="header 02 content"}
$response = Invoke-WebRequest -Method Post -Uri $uri -Body $body -Headers $header -ContentType "text/plain; charset=utf-8"
$jobid = (ConvertFrom-Json ($response.Content)).jobids[0]