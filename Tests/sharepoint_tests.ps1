Function Test-Download-FileFromLibrary()
{
	# 1) $sourcefile et $targetfile avec noms de fichiers ne contenant ni espace ni caracteres speciaux
	# 2) $sourcefile et $targetfile avec noms de fichiers contenant espaces et caracteres speciaux
	# 3) $sourcefile sur chemin réseau
	# 4) $sourcefile avec un nom de fichier qui n'existe pas
	# 5) $targetfile avec un nom de fichier qui n'existe pas
	# 6) $targetfile avec un nom de fichier en cours d'édition (ex: .docx ouvert dans word)

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
	Download-FileFromLibrary -SPContext $SPContext -SourceFile $sourcefile -TargetFile $targetfile

}

Function Test-Upload-FileToLibrary()
{
	Upload-FileToLibrary  -SiteURL $SiteURL -DocLibName $DocLibName -User $User -Password $Password -SourceFile $SourceFile -TargetDirectory $TargetDirectory
}
Function Test-Upload-AllFilesFromDirectory()
{

}
Function Test-Get-AllFilesFromDirectory()
{

}
Function Test-Get-SPContext
{

}
Function Test-Remove-SPFile()
{

}