Function Test-Send-SimpleMail()
{
    $username = 'sender@domain.com'
    $password = Read-Host "Enter password of $username" -AsSecureString
    $to = 'destination@domain.com'
    $subject = 'Test Mail'
    $body = 'This is for testing purposes'
    $server = 'smtp.office365.com'
    $port = 587
    $filelist =@("C:\temp\file1.txt","C:\temp\file2.txt")

    # Send mail with default server parameters, without attachment
    $return = Send-SimpleMail -UserName $username -Password $password -MailTo $to -MailTitle $subject -MailBody $body
    
    # Send mail with default server parameters, with files attached
    $return = Send-SimpleMail -UserName $username -Password $password -MailTo $to -MailTitle $subject -MailBody $body -AttachementsList $filelist

    # Send mail with custom server parameters
    $return = Send-SimpleMail -ServerName $server -ServerPort $port -UserName $username -Password $password -MailTo $to -MailTitle $subject -MailBody $body

}

Function Test-New-ExchangeMeeting()
{

    $username = "creator@domain.com"
    $password = Read-Host "Enter password of $username" -AsSecureString
    $ewsurl = "https://outlook.office365.com/EWS/Exchange.asmx"
    $title = "Test Meeting" 
    $body = "Body of test Meeting"
    $start = '202012031605'
    $end = '202012031725'
    $teams = $false
    $filelist =@("C:\temp\file1.txt","C:\temp\file2.txt")

    # Create simple Office 365 meeting, no Teams and no attachement
    $MeetingId = New-ExchangeMeeting -ExchangeUserName $username -ExchangePassword $password -ExchangeMeetingTitle $title -ExchangeMeetingBody $body -ExchangeMeetingStartDate $start -ExchangeMeetingEndDate $end -ExchangeMeetingIsTeams $teams

    # Create meeting for custom Exchange server, no Teams and no attachement

    # Create Teams Office 365 meeting

    # Create Office 365 meeting with attached files
    $MeetingId = New-ExchangeMeeting -ExchangeUserName $username -ExchangePassword $password -ExchangeMeetingTitle $title -ExchangeMeetingBody $body -ExchangeMeetingStartDate $start -ExchangeMeetingEndDate $end -ExchangeMeetingIsTeams $teams -ExchangeAttachementsList $filelist

}