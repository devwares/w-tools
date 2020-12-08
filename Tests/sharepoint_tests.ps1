Function Test-Get-FileFromLibrary()
{

	# 1) $sourcefile and $targetfile names without spaces and special characters
	# 2) $sourcefile et $targetfile names WITH spaces and special characters
	# 3) $sourcefile on a network share
	# 4) $sourcefile with path that doesn't exist
	# 5) $targetfile with path that doesn't exist
	# 6) $targetfile being edited (ex: .docx opened in ms-word)

	Get-module | Remove-Module
	Import-Module '..\Source code'

	$Config = get-content '..\Configuration Files\sharepoint_tests.json' | ConvertFrom-Json

	$siteurl = $config.'Download-FileFromLibrary'.SiteURL
	$user = $config.'Download-FileFromLibrary'.User
	$sourcefile = $config.'Download-FileFromLibrary'.SourceFile
	$targetfile = $Config.'Download-FileFromLibrary'.TargetFile

    # Method 1 : direct input
    $SecurePassword = Read-Host -AsSecureString

    # Method 2 : plain text (not recommended)
    #$Password = $config.'Download-FileFromLibrary'.Password
    #$SecurePassword = ($Password | ConvertTo-SecureString -asPlainText -Force)

    # Method 3 : encrypted key (preferred)
    #$key = Get-Content $config.'Download-FileFromLibrary'.Encrypted-Keyfile
    #$encpassword = $config.'Download-FileFromLibrary'.Encrypted-Password
    #$SecurePassword = $encpassword | ConvertTo-SecureString -Key $key

	$SPContext = Get-SPContext -SiteURL $siteurl -User $user -Password $SecurePassword
	Get-FileFromLibrary -SPContext $SPContext -SourceFile $sourcefile -TargetFile $targetfile

}

Function Test-Send-FileToLibrary()
{

}
Function Test-Send-AllFilesFromDirectory()
{

}
Function Test-Get-AllFilesFromDirectory()
{

    $siteurl = "https://contoso.sharepoint.com/personal/john_smith_contoso_com"
    $user = "john.smith@contoso.com"
    $password = Read-Host "Enter password of $user" -AsSecureString

    $SPContext = Get-SPContext -SiteURL $siteurl -User $user -Password $password
    $LibraryName = "Documents"
    $DirectoryName = "compatibility"

    # Get all files from directory
    $filelist = Get-AllFilesFromDirectory -SPContext $SPContext -LibraryName $LibraryName -DirectoryName $DirectoryName

    # Get all files from directory (recursively)
    $filelist = Get-AllFilesFromDirectory -SPContext $SPContext -LibraryName $LibraryName -DirectoryName $DirectoryName -Recursive $True

    # File all PDF files > 1MB from list
    $filelist | select-object Name, Length, @{Label = 'UcName' ; Expression = {$_.Name.ToUpper()}} | Where-Object UcName -match ".PDF" | Where-Object Length -GT 1000000

}
Function Test-Get-SPContext
{
    $siteurl = "https://contoso.sharepoint.com/personal/john_smith_contoso_com"
    $user = "john.smith@contoso.com"
    $password = Read-Host "Enter password of $user" -AsSecureString
    $SPContext = Get-SPContext -SiteURL $siteurl -User $user -Password $password
}
Function Test-Remove-SPFile()
{

}