#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Function Download-FileFromLibrary()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $User,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password,
        [Parameter(Mandatory=$true)] [string] $SourceFile,
        [Parameter(Mandatory=$true)] [string] $TargetFile
    )
 
    Try {

        # Credentials
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $Password)
 
        #Setup the context
        $Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Ctx.Credentials = $Credentials
     
        #sharepoint online powershell download file from library
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($Ctx,$SourceFile)
        $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
        $FileInfo.Stream.CopyTo($WriteStream)
        $WriteStream.Close()
 
        Write-host -f Green "File '$SourceFile' Downloaded to '$TargetFile' Successfully!" $_.Exception.Message
  }
    Catch {
        write-host -f Red "Error Downloading File!" $_.Exception.Message
    }
}
Function Upload-FileToLibrary()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $DocLibName,
        [Parameter(Mandatory=$true)] [string] $User,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password,
        [Parameter(Mandatory=$true)] [String] $SourceFile,
        [Parameter(Mandatory=$false)] [string] $TargetDirectory
    )
 
    Try {

        # Credentials
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $Password)

        #Setup the context
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Context.Credentials = $Credentials

        #Retrieve list
        $List = $Context.Web.Lists.GetByTitle($DocLibName)
        $Context.Load($List)
        $Context.Load($List.RootFolder)
        $Context.ExecuteQuery()
        $ServerRelativeUrlOfRootFolder = $List.RootFolder.ServerRelativeUrl
        $UploadFolderUrl=  $ServerRelativeUrlOfRootFolder + "/" + $TargetDirectory

        #Get Object for File
        $FileName = Split-Path -Path $SourceFile -Leaf -Resolve
        $FilePath = Split-Path $SourceFile
        $File = (Get-ChildItem $FilePath -file | Where-Object {$_.Name -eq $FileName})

        #Upload file
        $FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $File

        If($TargetDirectory -eq $null)
        {
            $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
        }
        Else
        {
            $targetFolder = $Context.Web.GetFolderByServerRelativeUrl($uploadFolderUrl)
            $Upload = $targetFolder.Files.Add($FileCreationInfo);
        }

        $Context.Load($Upload)
        $Context.ExecuteQuery()

        Write-host -f Green "File '$SourceFile' Uploaded to '$SiteURL$DocLibName' Successfully!" $_.Exception.Message
        
    }

    Catch {
        write-host -f Red "Error Uploading Files!" $_.Exception.Message
    }

}
Function Upload-AllFilesFromDirectory()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $DocLibName,
        [Parameter(Mandatory=$true)] [string] $User,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password,
        [Parameter(Mandatory=$true)] [string] $SourceDirectory,
        [Parameter(Mandatory=$false)] [string] $TargetDirectory
    )
 
    Try {

        # Credentials
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $Password)

        #Setup the context
        $Context = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $Context.Credentials = $Credentials

        #Retrieve list
        $List = $Context.Web.Lists.GetByTitle($DocLibName)
        $Context.Load($List)
        $Context.Load($List.RootFolder)
        $Context.ExecuteQuery()
        $ServerRelativeUrlOfRootFolder = $List.RootFolder.ServerRelativeUrl
        $UploadFolderUrl=  $ServerRelativeUrlOfRootFolder + "/" + $TargetDirectory

        #Upload file
        Foreach ($File in (Get-ChildItem $SourceDirectory -File))
        {

            $FileStream = New-Object IO.FileStream($File.FullName,[System.IO.FileMode]::Open)
            $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
            $FileCreationInfo.Overwrite = $true
            $FileCreationInfo.ContentStream = $FileStream
            $FileCreationInfo.URL = $File

            If($TargetDirectory -eq $null)
            {
                $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
            }
            Else
            {
                $targetFolder = $Context.Web.GetFolderByServerRelativeUrl($uploadFolderUrl)
                $Upload = $targetFolder.Files.Add($FileCreationInfo);
            }

            $Context.Load($Upload)
            $Context.ExecuteQuery()

            Write-host -f Green "File '$File' Uploaded to '$SiteURL$DocLibName' Successfully!" $_.Exception.Message
        }
        
    }

    Catch {
        write-host -f Red "Error Uploading Files!" $_.Exception.Message
    }

}

