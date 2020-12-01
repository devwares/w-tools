Function Test-Send-EchangeMail()
{
    $username = 'user@officedomain.com'
    $password = Read-Host "Enter password of $username" -AsSecureString
    $to = 'destination@domain.com'
    $subject = 'Test Mail'
    $body = 'This is for testing purposes'
    $server = 'smtp.office365.com'
    $port = 587
    $filelist =@("C:\temp\file1.txt","C:\temp\file2.txt")

    # Send mail with default server parameters, without attachment
    $return = Send-ExchangeMail -ExchangeUserName $username -ExchangePassword $password -ExchangeMailTo $to -ExchangeMailTitle $subject -ExchangeMailBody $body
    
    # Send mail with default server parameters, with files attached
    $return = Send-ExchangeMail -ExchangeUserName $username -ExchangePassword $password -ExchangeMailTo $to -ExchangeMailTitle $subject -ExchangeMailBody $body -ExchangeAttachementsList $filelist

    # Send mail with custom server parameters
    $return = Send-ExchangeMail -ExchangeServerName $server -ExchangeServerPort $port -ExchangeUserName $username -ExchangePassword $password -ExchangeMailTo $to -ExchangeMailTitle $subject -ExchangeMailBody $body

}