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

function Send-ExchangeMail
{

    # Param(
    #     [parameter(Mandatory=$False)][string] $ServerName,
    #     [parameter(Mandatory=$False)][int32] $ServerPort,
    #     [parameter(Mandatory=$False)][bool] $ServerUseSsl,
    #     [parameter(Mandatory=$True)][string] $UserName,
    #     [parameter(Mandatory=$True)][SecureString] $Password,
    #     [parameter(Mandatory=$True)][string] $MailTo,
    #     [parameter(Mandatory=$True)][string] $MailTitle,
    #     [parameter(Mandatory=$True)][string] $MailBody,
    #     [parameter(Mandatory=$False)][bool] $BodyAsHtml,
    #     [parameter(Mandatory=$False)][array] $AttachementsList
    # )

    # Load Exchange Web Services Managed API
    $EWSServicePath = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'
    Import-Module $EWSServicePath

}


# TODO : fix return string
# TODO : isteams
function New-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingBody,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingStartDate,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingEndDate,
        [parameter(Mandatory=$True)][bool] $ExchangeMeetingIsTeams,
        [parameter(Mandatory=$False)][array] $ExchangeAttachementsList
    )

    Try {

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

        # Credentials
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeUserName, $ExchangePassword
        $exchService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($cred)

        # Set default Office 365 Url for Exchange Web Service if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceUrl)){$ExchangeWebServiceUrl = "https://outlook.office365.com/EWS/Exchange.asmx"}
        $exchService.Url= new-object Uri($ExchangeWebServiceUrl)

        # setup extended property set
        $CleanGlobalObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting, 0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropSet.Add($CleanGlobalObjectId);

        # Bind to the Calendar folder  
        # $folderid generates a true
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)     
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService,$folderid)

        # Convert date strings to system.datetime
        $MeetingStartDatetime=[System.DateTime]::ParseExact($ExchangeMeetingStartDate,'yyyyMMddHHmm',$null)
        $MeetingEndDatetime=[System.DateTime]::ParseExact($ExchangeMeetingEndDate,'yyyyMMddHHmm',$null)

        # Create Appointment object
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $exchService
            $appointment.Subject = $MeetingSubject
            $appointment.Body = $MeetingBody
            $appointment.Start = $MeetingStartDatetime;
            $appointment.End = $MeetingEndDatetime;

        # Add attachment(s) if specified
        If ($ExchangeAttachementsList){
            ForEach ($file in $ExchangeAttachementsList){
                $appointment.Attachments.AddFileAttachment($file);
            }
        }

        #$appointment.conferencetype = 2
        # retrouver la valeur de conferencetype pour un rendez-vous teams créé, avec son ID 
        # ex : 040000008200E00074C5B7101A82E0080000000064701EB606C8D601000000000000000010000000B27DED920EA4D34D983AAA3EA6A103A8


        $RequiredAttendees = $row.advisorEmail;
        if($RequiredAttendees) {$RequiredAttendees | %{[void]$appointment.RequiredAttendees.Add($_)}}

        $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
                 
        # Set the unique id for the appointment and convert to text
            $appointment.Load($psPropSet);
            $CalIdVal = $null;
            $appointment.TryGetProperty($CleanGlobalObjectId, [ref]$CalIdVal);
            $CalIdVal64 = [Convert]::ToBase64String($CalIdVal)

            # debug           
            $appointment | Export-Csv -Path 'c:\0\tmp\test.csv' -Delimiter ";"

        }

    Catch{
        Write-Error "Meeting was not created --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    # Return Unique id for created meeting
    return [string]$CalIdVal64

}

function Edit-ExchangeMeeting
{

    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId
    )

    # MeetingId if modified, else null
    $ExchangeMeetingState = $null
    return $ExchangeMeetingState
}

function Remove-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId
    )

    Try{

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010)

        # Credentials
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeUserName, $ExchangePassword
        $exchService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($cred)

        # Set default Office 365 Url for Exchange Web Service if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceUrl)){$ExchangeWebServiceUrl = "https://outlook.office365.com/EWS/Exchange.asmx"}
        $exchService.Url= new-object Uri($ExchangeWebServiceUrl)

        # setup extended property set
        $CleanGlobalObjectId = new-object Microsoft.Exchange.WebServices.Data.ExtendedPropertyDefinition([Microsoft.Exchange.WebServices.Data.DefaultExtendedPropertySet]::Meeting, 0x23, [Microsoft.Exchange.WebServices.Data.MapiPropertyType]::Binary);
        $psPropSet = new-object Microsoft.Exchange.WebServices.Data.PropertySet([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties);
        $psPropSet.Add($CleanGlobalObjectId);

        # Bind to the Calendar folder  
        # $folderid generates a true
        $folderid = new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Calendar,$MailboxName)     
        $Calendar = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($exchService,$folderid)

        # Find Item
        $ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1) 
        $sfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo($CleanGlobalObjectId, $ExchangeMeetingId);
        $fiResult = $Calendar.FindItems($sfSearchFilter, $ivItemView) 

        foreach ($a in $fiResult) { 
            $a.Delete(0);
            # record cancellation in the data source            
        }

    }

    Catch{
        Write-Error "Error during meeting deletion --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    return $ExchangeMeetingState
}