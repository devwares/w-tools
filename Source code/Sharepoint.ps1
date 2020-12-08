#Load SharePoint CSOM Assemblies
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll"
Add-Type -Path "C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll"
Function Get-FileFromLibrary()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $SPContext, 
        [Parameter(Mandatory=$true)] [string] $SourceFile,
        [Parameter(Mandatory=$true)] [string] $TargetFile
    )
 
    Try {
   
        #sharepoint online powershell download file from library
        $FileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($SPContext,$SourceFile)
        $WriteStream = [System.IO.File]::Open($TargetFile,[System.IO.FileMode]::Create)
        $FileInfo.Stream.CopyTo($WriteStream)
        $WriteStream.Close()
 
        Write-host -f Green "File '$SourceFile' Downloaded to '$TargetFile' Successfully!" $_.Exception.Message
  }
    Catch {
        write-host -f Red "Error Downloading File!" $_.Exception.Message
    }
}
Function Send-FileToLibrary()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $SPContext,
        [Parameter(Mandatory=$true)] [string] $DocLibName,
        [Parameter(Mandatory=$true)] [String] $SourceFile,
        [Parameter(Mandatory=$false)] [string] $TargetDirectory
    )
 
    Try {

        #Retrieve list
        $List = $SPContext.Web.Lists.GetByTitle($DocLibName)
        $SPContext.Load($List)
        $SPContext.Load($List.RootFolder)
        $SPContext.ExecuteQuery()
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
            $targetFolder = $SPContext.Web.GetFolderByServerRelativeUrl($uploadFolderUrl)
            $Upload = $targetFolder.Files.Add($FileCreationInfo);
        }

        $SPContext.Load($Upload)
        $SPContext.ExecuteQuery()

        Write-host -f Green "File '$SourceFile' Uploaded to '$SiteURL$DocLibName/$TargetDirectory' Successfully!" $_.Exception.Message
        
    }

    Catch {
        write-host -f Red "Error Uploading Files!" $_.Exception.Message
    }

}
Function Send-AllFilesFromDirectory()
{ 
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $SPContext,
        [Parameter(Mandatory=$true)] [string] $DocLibName,
        [Parameter(Mandatory=$true)] [string] $SourceDirectory,
        [Parameter(Mandatory=$false)] [string] $TargetDirectory
    )
 
    Try {

        #Retrieve list
        $List = $SPContext.Web.Lists.GetByTitle($DocLibName)
        $SPContext.Load($List)
        $SPContext.Load($List.RootFolder)
        $SPContext.ExecuteQuery()
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
                $targetFolder = $SPContext.Web.GetFolderByServerRelativeUrl($uploadFolderUrl)
                $Upload = $targetFolder.Files.Add($FileCreationInfo);
            }

            $SPContext.Load($Upload)
            $SPContext.ExecuteQuery()

            Write-host -f Green "File '$File' Uploaded to '$SiteURL$DocLibName' Successfully!" $_.Exception.Message
        }
        
    }

    Catch {
        write-host -f Red "Error Uploading Files!" $_.Exception.Message
    }

}

Function Get-AllFilesFromDirectory()
{
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $SPContext,        
        [Parameter(Mandatory=$true)] [string] $LibraryName,
        [Parameter(Mandatory=$false)] [string] $DirectoryName,
        [Parameter(Mandatory=$false)] [bool] $Recursive
    )
    Function Get-AllFilesFromFolder()
    {
        param
        (
            [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.Folder]$Folder,
            [Parameter(Mandatory=$false)] [bool] $Recursive
        )
    
        #Get All Files of the Folder
        $Ctx =  $Folder.Context
        $Ctx.load($Folder.files)
        $Ctx.ExecuteQuery()
      
        # Initialize object
        $SPFileListFromFolder = @()

        # Loop on all files in folder
        ForEach ($File in $Folder.files)
        {
            #Get the File Name or do something
            # Write-host -f Green $File.Name
            $SPFileListFromFolder += $File

        }
    
        if ($Recursive){
            #Recursively Call the function to get files of all folders
            $Ctx.load($Folder.Folders)
            $Ctx.ExecuteQuery()
    
            #Exclude "Forms" system folder and iterate through each folder
            ForEach($SubFolder in $Folder.Folders | Where {$_.Name -ne "Forms"})
            {
                $SPFileListRecursive = Get-AllFilesFromFolder -Folder $SubFolder -Recursive $true
                $SPFileListFromFolder += $SPFileListRecursive
            }
        }

        Return $SPFileListFromFolder

    }

    #Get the Library and Its Root Folder
    $Library = $SPContext.web.Lists.GetByTitle($LibraryName)
    $SPContext.Load($Library)
    $SPContext.Load($Library.RootFolder)
    $SPContext.ExecuteQuery()

    #Call the function to get Files of the Root Folder or specified Folder
    if ([string]::IsNullOrEmpty($DirectoryName)){
        $AllFilesFromDirectory = Get-AllFilesFromFolder -Folder $Library.RootFolder -Recursive $Recursive
    }
    else{
        $ServerRelativeUrlOfRootFolder = $Library.RootFolder.ServerRelativeUrl
        $TargetFolderUrl=  $ServerRelativeUrlOfRootFolder + "/" + $DirectoryName
        $TargetFolder = $SPContext.Web.GetFolderByServerRelativeUrl($TargetFolderUrl)
        $AllFilesFromDirectory = Get-AllFilesFromFolder -Folder $TargetFolder -Recursive $Recursive
    }

    Return $AllFilesFromDirectory | select-object Name, ServerRelativeUrl, Length, TimeCreated, TimeLastModified, ContentTag, Etag, IrmEnabled, MajorVersion, MinorVersion, UIVersion, UIVersionLabel 

}

Function Get-SPContext
{

    param
    (
        [Parameter(Mandatory=$true)] [string] $SiteURL,
        [Parameter(Mandatory=$true)] [string] $User,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password
    )

    Try {
        # Credentials
        $Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($User, $Password)

        #Setup the context
        $SPContext = New-Object Microsoft.SharePoint.Client.ClientContext($SiteURL)
        $SPContext.Credentials = $Credentials

        # Return context
        return $SPContext
     }

    Catch {
        write-host -f Red "Error:" $_.Exception.Message
    }

}

Function Remove-SPFile()
{
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $SPContext,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientObject] $SPFile      
    )

    Write-Host "Je supprime le fichier " $SPFile.Name
    $SPFile.Recycle()
    $SPContext.ExecuteQuery()

}