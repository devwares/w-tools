function IsNumeric {
    param (
        [Parameter(Mandatory=$true)]$Value
    )
    
    $typeValue = $Value.getTypeCode().value__
    if ($typeValue -ge 5 -and $typeValue -le 15) {return $True}
    else {return $False}

}

function Get-UnixTimeStamp {
    return (get-date -date (get-date).ToUniversalTime() -UFormat %s).split(',')[0]
}