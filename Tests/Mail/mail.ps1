  
Function Test-Send-SimpleMail()
{
    $username = 'sender@domain.com'
    $password = Read-Host "Enter password of $username" -AsSecureString
    $to = 'joe@domain.com', "john@domain.net"
    $subject = 'Test Mail'
    $body = 'This is for testing purposes'
    $server = 'smtp.office365.com'
    $port = 587
    $filelist =@("C:\temp\file1.txt","C:\temp\file2.txt")

    # Send mail to single user, with default server parameters, without attachment
    $return = Send-SimpleMail -UserName $username -Password $password -MailTo "john@domain.org" -MailTitle $subject -MailBody $body
    
    # Send mail with default server parameters, with files attached
    $return = Send-SimpleMail -UserName $username -Password $password -MailTo $to -MailTitle $subject -MailBody $body -AttachementsList $filelist

    # Send mail with custom server parameters
    $return = Send-SimpleMail -ServerName $server -ServerPort $port -UserName $username -Password $password -MailTo $to -MailTitle $subject -MailBody $body

}