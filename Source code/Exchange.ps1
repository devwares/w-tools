Import-Module ExchangeOnlineManagement

function ExchangeSendMail
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMailTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMailBody,
        [parameter(Mandatory=$False)][array] $ExchangeAttachementsList
    )

    # True if sent
    $ExchangeMailSent = $False
    return $ExchangeMailSent

}

function ExchangeCreateMeeting
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
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

function ExchangeModifyMeeting
{
    Param(
        [parameter(Mandatory=$True)][string] $ExchangeServerName,
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

