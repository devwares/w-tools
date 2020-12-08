param
(
    [Parameter(Mandatory=$true)] [string] $name,
    [Parameter(Mandatory=$true)] [string] $address
)

$uri = "https://s17events.azure-automation.net/webhooks?token=8J7Gg9%3DFffeEEd34dsSSJyhy66tRffd4tg7OjquyPQA%3t"

$jsoncontent  = @(
            @{ Name="$name";Address="$address"}
        )
$body = ConvertTo-Json -InputObject $jsoncontent
$header = @{ message="Test Header"}
$response = Invoke-WebRequest -Method Post -Uri $uri -Body $body -Headers $header
$jobid = (ConvertFrom-Json ($response.Content)).jobids[0]