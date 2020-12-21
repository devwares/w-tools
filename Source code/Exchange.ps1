# SUGGESTION : function that creates Appointment object from existing Id

function Send-ExchangeMail
{

    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$True)][string] $ExchangeMailTo,
        [parameter(Mandatory=$False)][string] $ExchangeMailCc,
        [parameter(Mandatory=$False)][string] $ExchangeMailBcc,
        [parameter(Mandatory=$True)][string] $ExchangeMailTitle,
        [parameter(Mandatory=$True)][string] $ExchangeMailBody,
        [parameter(Mandatory=$False)][string] $ExchangeMailBodyType,
        [parameter(Mandatory=$False)][string] $ExchangeAttachments
    )

    try {

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

        # Credentials
        $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $ExchangeUserName, $ExchangePassword
        $exchService.Credentials = new-object Microsoft.Exchange.WebServices.Data.WebCredentials($cred)

        # Set default Office 365 Url for Exchange Web Service if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceUrl)){$ExchangeWebServiceUrl = "https://outlook.office365.com/EWS/Exchange.asmx"}
        $exchService.Url= new-object Uri($ExchangeWebServiceUrl)

        # Create the email message and set the Subject and Body
        $message = New-Object Microsoft.Exchange.WebServices.Data.EmailMessage -ArgumentList $exchService
        $message.Subject = $ExchangeMailTitle
        $message.Body = $ExchangeMailBody + "`r`n"

        # Set Body type (Default = HTML)
        If ($ExchangeMailBodyType.ToLower() -eq "text"){
            $message.Body.BodyType = 'Text'
        }
        Else{
            $message.Body.BodyType = 'HTML'
        }

        # Add attachments if specified
        if ($ExchangeAttachments){
            # Split attachment string into array
            $ExchangeAttachmentsList = $ExchangeAttachments.Split(";");
            ForEach ($file in $ExchangeAttachmentsList){
                # Check file path before attaching
                if (Test-Path $file){
                    $message.Attachments.AddFileAttachment($file) | Out-Null; # Out-Null used here not to go into pipeline
                }
                else {
                    Write-Error "File not found $file --> $($_.Exception.Message)" -ErrorAction:Continue
                    return $False                    
                }
            }
        }

        # Add each specified "To" recipient
        # Split attachment string into array
        $ExchangeMailToList = $ExchangeMailTo.Split(";");
        ForEach ($Recipient in $ExchangeMailToList)
        {
            $message.ToRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
        }

        # Add each specified "Cc" recipient
        # Split attachment string into array
        $ExchangeMailCcList = $ExchangeMailCc.Split(";");
        ForEach ($Recipient in $ExchangeMailCcList)
        {
            $message.CcRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
        }

        # Add each specified "Bcc" recipient
        # Split attachment string into array
        $ExchangeMailBccList = $ExchangeMailBcc.Split(";");
        ForEach ($Recipient in $ExchangeMailBccList)
        {
            $message.BccRecipients.Add($Recipient) | Out-Null # Out-Null used here not to go into pipeline
        }

        # Send the message (copy gets saved in sent items of the user)
        $message.SendAndSaveCopy() | Out-Null # Out-Null used here not to go into pipeline

    }

    catch [exception] {
        Write-Error "Mail not sent --> $($_.Exception.Message)" -ErrorAction:Continue
        return $False
    }

    return $True

}

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
        [parameter(Mandatory=$True)][string] $ExchangeRequiredAttendees,
        [parameter(Mandatory=$False)][string] $ExchangeOptionalAttendees,
        [parameter(Mandatory=$False)][string] $ExchangeAttachments
    )

    Try {

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

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
        $MeetingStartDatetime=[System.DateTime]::ParseExact($ExchangeMeetingStartDate,'yyyy-MM-ddTHH:mm:ss',$null)
        $MeetingEndDatetime=[System.DateTime]::ParseExact($ExchangeMeetingEndDate,'yyyy-MM-ddTHH:mm:ss',$null)

        # Split attendees strings
        if (-not [string]::IsNullOrEmpty($ExchangeRequiredAttendees)) {$ExchangeRequiredAttendeesList = $ExchangeRequiredAttendees.Split(";") }
        if (-not [string]::IsNullOrEmpty($ExchangeOptionalAttendees)) {$ExchangeOptionalAttendeesList = $ExchangeOptionalAttendees.Split(";") }

        # Create Appointment object
        $appointment = New-Object Microsoft.Exchange.WebServices.Data.Appointment -ArgumentList $exchService
            $appointment.Subject = $ExchangeMeetingTitle
            $appointment.Body = $ExchangeMeetingBody
            $appointment.Start = $MeetingStartDatetime;
            $appointment.End = $MeetingEndDatetime;
            foreach ($attendee in $ExchangeRequiredAttendeesList) {
                $null = $appointment.RequiredAttendees.Add($attendee)
            }
            foreach ($attendee in $ExchangeOptionalAttendeesList) {
                $null = $appointment.OptionalAttendees.Add($attendee)
            }

        # Add attachment(s) if specified
        If ($ExchangeAttachments){
            $ExchangeAttachmentsList = $ExchangeAttachments.Split(";")
            ForEach ($file in $ExchangeAttachmentsList){
                $appointment.Attachments.AddFileAttachment($file) | Out-Null ; # Out-Null used here not to go into pipeline
            }
        }

        #$RequiredAttendees = $row.advisorEmail;
        #if($RequiredAttendees) {$RequiredAttendees | %{[void]$appointment.RequiredAttendees.Add($_)}}

        # Save new meeting
        $appointment.Save([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)
                 
        # Set the unique id for the appointment and convert to text
            $appointment.Load($psPropSet);
            $CalIdVal = $null;
            $appointment.TryGetProperty($CleanGlobalObjectId, [ref]$CalIdVal) | Out-Null ; # Out-Null used here not to go into pipeline
            $CalIdVal64 = [Convert]::ToBase64String($CalIdVal)

        }

    Catch{
        Write-Error "Meeting was not created --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    # Return Unique id for created meeting
    return [string]$CalIdVal64

}

# SUGGESTION : "add", "remove" and "update" options for attendees and attachments
# SUGGESTION : add "no notification" option
function Edit-ExchangeMeeting
{

    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$False)][string] $ExchangeMeetingTitle,
        [parameter(Mandatory=$False)][string] $ExchangeMeetingBody,
        [parameter(Mandatory=$False)][string] $ExchangeMeetingStartDate,
        [parameter(Mandatory=$False)][string] $ExchangeMeetingEndDate,
        [parameter(Mandatory=$False)][string] $ExchangeRequiredAttendees,
        [parameter(Mandatory=$False)][string] $ExchangeOptionalAttendees,
        [parameter(Mandatory=$False)][array] $ExchangeAttachments,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId
    )

    Try{

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

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

        # Returns Null if no meeting found with specified Id
        if ($fiResult.TotalCount -eq 0){return $null}

        ForEach ($appointment in $fiResult) { 

            # Update Attachments if specified
            If ($ExchangeAttachments){
                $ExchangeAttachmentsList = $ExchangeAttachments.Split(";")
                # Clear attachments collection
                $appointment.Attachments.Clear()
                # Save updated meeting without sending updates to attendees, to clear old attachments
                $appointment.Update([Microsoft.Exchange.WebServices.Data.ConflictResolutionMode]::AutoResolve, $True)
                # Add new attachment(s)
                ForEach ($file in $ExchangeAttachmentsList){
                    $appointment.Attachments.AddFileAttachment($file);
                }
            }

            # Update required attendees if specified
            If (-not [string]::IsNullOrEmpty($ExchangeRequiredAttendees)){
                # Split attendees strings
                $ExchangeRequiredAttendeesList = $ExchangeRequiredAttendees.Split(";")
                # Clear attendees collection
                $appointment.RequiredAttendees.Clear()
                # Add new attendee(s)
                foreach ($attendee in $ExchangeRequiredAttendeesList) {
                    $null = $appointment.RequiredAttendees.Add($attendee)
                }
            }

            # Update optional attendees if specified
            If (-not [string]::IsNullOrEmpty($ExchangeOptionalAttendees)){
                # Split attendees strings
                $ExchangeOptionalAttendeesList = $ExchangeOptionalAttendees.Split(";")
                # Clear attendees collection
                $appointment.OptionalAttendees.Clear()
                # Add new attendee(s)
                foreach ($attendee in $ExchangeOptionalAttendeesList) {
                    $null = $appointment.OptionalAttendees.Add($attendee)
                }
            }

            # Update subject if specified
            If (-not [string]::IsNullOrEmpty($ExchangeMeetingTitle)){
                $appointment.Subject = $ExchangeMeetingTitle
            }

            # Update body if specified
            If (-not [string]::IsNullOrEmpty($ExchangeMeetingBody)){
                $appointment.Body = $ExchangeMeetingBody
            }

            # Update start date if specified
            If (-not [string]::IsNullOrEmpty($ExchangeMeetingStartDate)){
                # Convert date strings to system.datetime
                $MeetingStartDatetime=[System.DateTime]::ParseExact($ExchangeMeetingStartDate,'yyyy-MM-ddTHH:mm:ss',$null)
                # Set meeting start date
                $appointment.Start = $MeetingStartDatetime;
            }

            # Update end date if specified
            If (-not [string]::IsNullOrEmpty($ExchangeMeetingEndDate)){
                # Convert date strings to system.datetime
                $MeetingEndDatetime=[System.DateTime]::ParseExact($ExchangeMeetingEndDate,'yyyy-MM-ddTHH:mm:ss',$null)
                # Set meeting end date
                $appointment.End = $MeetingEndDatetime;
            }

            # Save updated meeting and send notification to all attendees
            $appointment.Update([Microsoft.Exchange.WebServices.Data.SendInvitationsMode]::SendToAllAndSaveCopy)

        }

    }

    Catch{
        Write-Error "Error during meeting update --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    return $ExchangeMeetingId

}

function Stop-ExchangeMeeting
{
    Param(
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceUrl,
        [parameter(Mandatory=$False)][string] $ExchangeWebServiceDll,
        [parameter(Mandatory=$True)][string] $ExchangeUserName,
        [parameter(Mandatory=$True)][SecureString] $ExchangePassword,
        [parameter(Mandatory=$False)][switch] $Delete,
        [parameter(Mandatory=$True)][string] $ExchangeMeetingId
    )

    Try{

        # Set default path to Dll if not specified
        if ([string]::IsNullOrEmpty($ExchangeWebServiceDll)){$ExchangeWebServiceDll = 'C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll'}
        
        # Load Exchange Web Services API
        Import-Module $ExchangeWebServiceDll

        # Create EWS object
        $exchService = new-object Microsoft.Exchange.WebServices.Data.ExchangeService([Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2013)

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

        # Returns Null if no meeting found with specified Id
        if ($fiResult.TotalCount -eq 0){return $null}

        foreach ($appointment in $fiResult) { 
            
            # Deletes meeting or cancel it depending of "-Delete" argument
            If ($Delete){
                $appointment.Delete(0);
            }
            else {
                $appointment.CancelMeeting() | Out-Null # Out-Null used here not to go into pipeline
            }
            
        }

    }

    Catch{
        Write-Error "Error during meeting cancelation --> $($_.Exception.Message)" -ErrorAction:Continue
        return $null
    }

    return $ExchangeMeetingId

}