function Send-ExchangeMail
{
    
    Param(
        [parameter(Mandatory=$False)][string] $ExchangeServerName,
        [parameter(Mandatory=$False)][int32] $ExchangeServerPort,
        [parameter(Mandatory=$False)][bool] $ExchangeServerUseSsl,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMailTo,
        [parameter(Mandatory=$True)][string] $ExchangeMailTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMailBody,
        [parameter(Mandatory=$False)][bool] $ExchangeMailBodyAsHtml,
        [parameter(Mandatory=$False)][array] $ExchangeAttachementsList
    )

    # Set default values for Exchange server if not specified
    if ([string]::IsNullOrEmpty($ExchangeServerName)){$ExchangeServerName = 'smtp.office365.com'}
    if (!$ExchangeServerPort){$ExchangeServerPort = 25} # alternate value for Exchange = 587
    if (!$ExchangeServerUseSsl){$ExchangeServerUseSsl = $true}

    # Credentials
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeUserName, $ExchangePassword
    
    # Set mail default mail parameters if not specified
    if (!$ExchangeMailBodyAsHtml){$ExchangeMailBodyAsHtml = $False}

    # Prepare Hash content
    $ExchangeMailParameters = @{

    To = $ExchangeMailTo
    From = $ExchangeUserName
    Subject = $ExchangeMailTitle
    Body = $ExchangeMailBody
    BodyAsHtml = $ExchangeMailBodyAsHtml
    SmtpServer = $ExchangeServerName
    UseSSL = $ExchangeServerUseSsl
    Credential = $cred
    Port = $ExchangeServerPort

    }
 
    # Send mail using hash content
    try{
        if (!$ExchangeAttachementsList){Send-MailMessage @ExchangeMailParameters -ErrorAction Stop}
        else {Send-MailMessage @ExchangeMailParameters -Attachments $ExchangeAttachementsList -ErrorAction Stop}
    }
    catch {
        Write-Error "Mail was not sent --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    return $True

}

function New-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
        [parameter(Mandatory=$True)][string] $ExchangeServerPort,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingBody,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingTime,
        [parameter(Mandatory=$True)][bool] $ExchangeMeetingIsTeams,
        [parameter(Mandatory=$False)][array] $ExchangeAttachementsList
    )

    # MeetingId if created, else null
    $ExchangeMeetingState = $null
    return $ExchangeMeetingState

}

function Edit-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
        [parameter(Mandatory=$True)][string] $ExchangeServerPort,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingBody,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingTime,
        [parameter(Mandatory=$True)][bool] $ExchangeMeetingIsTeams,
        [parameter(Mandatory=$False)][array] $ExchangeAttachementsList
    )

    # MeetingId if modified, else null
    $ExchangeMeetingState = $null
    return $ExchangeMeetingState
}

function Remove-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
        [parameter(Mandatory=$True)][string] $ExchangeServerPort,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId
    )

    # MeetingId if modified, else null
    $ExchangeMeetingState = $null
    return $ExchangeMeetingState
}