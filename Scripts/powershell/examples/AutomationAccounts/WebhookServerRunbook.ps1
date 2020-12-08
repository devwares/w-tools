param
(
    [Parameter (Mandatory = $true)]
    [object] $WebhookData
)

$HeaderMessageContent = $WebhookData.RequestHeader.message
$JsonContent = (ConvertFrom-Json -InputObject $WebhookData.RequestBody)

$Name = $JsonContent | Select-object -expand "Name"
$Address = $JsonContent | Select-object -expand "Address"

Write-Output "Header message = $HeaderMessageContent"
Write-Output "Name = $Name"
Write-Output "Address = $Address"