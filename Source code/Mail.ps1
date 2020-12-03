function Send-SimpleMail
{
    
    Param(
        [parameter(Mandatory=$False)][string] $ServerName,
        [parameter(Mandatory=$False)][int32] $ServerPort,
        [parameter(Mandatory=$False)][bool] $ServerUseSsl,
        [parameter(Mandatory=$True)][string] $UserName,
        [parameter(Mandatory=$True)][SecureString] $Password,
        [parameter(Mandatory=$True)][string] $MailTo,
        [parameter(Mandatory=$True)][string] $MailTitle,
        [parameter(Mandatory=$True)][string] $MailBody,
        [parameter(Mandatory=$False)][bool] $BodyAsHtml,
        [parameter(Mandatory=$False)][array] $AttachementsList
    )

    # Set Office 365 default values for SMTP server if not specified
    if ([string]::IsNullOrEmpty($ServerName)){$ServerName = 'smtp.office365.com'}
    if (!$ServerPort){$ServerPort = 25} # alternate value for Simple = 587
    if (!$ServerUseSsl){$ServerUseSsl = $true}

    # Credentials
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, $Password
    
    # Set mail default mail parameters if not specified
    if (!$BodyAsHtml){$BodyAsHtml = $False}

    # Prepare Hash content
    $MailParameters = @{

    To = $MailTo
    From = $UserName
    Subject = $MailTitle
    Body = $MailBody
    BodyAsHtml = $BodyAsHtml
    SmtpServer = $ServerName
    UseSSL = $ServerUseSsl
    Credential = $cred
    Port = $ServerPort

    }
 
    # Send mail using hash content
    try{
        if (!$AttachementsList){Send-MailMessage @MailParameters -ErrorAction Stop}
        else {Send-MailMessage @MailParameters -Attachments $AttachementsList -ErrorAction Stop}
    }
    catch {
        Write-Error "Mail was not sent --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    return $True

}