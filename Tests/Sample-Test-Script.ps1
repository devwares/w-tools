#Region Encrypt functions
function Export-EncryptedSecureString(){  
    param
    (
        [Parameter(Mandatory=$true)] [string] $KeyFile,
        [Parameter(Mandatory=$true)] [string] $PasswordFile,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password
    )

    # Create and export Key if doesn't exist
    if (Test-Path $KeyFile -PathType leaf)
    {
        $Key = Get-Content $KeyFile
    }
    else {
        $Key = New-Object Byte[] 16
        [Security.Cryptography.RNGCryptoServiceProvider]::Create().GetBytes($key)
        $Key | Out-File $KeyFile
    }

    # Create and export Password File
    $Password | ConvertFrom-SecureString -Key $Key | Out-File $PasswordFile

}

function Get-EncryptedPassword {

    Param (

        [Parameter(mandatory=$true)][string]$KeyFile,
        [Parameter(mandatory=$true)][string]$PasswordFile

    )

    $Key = Get-Content $KeyFile
    $EncryptedPassword = Get-Content $PasswordFile
    Return ($EncryptedPassword | ConvertTo-SecureString -Key $Key)

}
#Endregion

#Region Test function
function Test-Function-1 {

    Param (
        [Parameter(mandatory=$true)][pscredential]$Credentials
    )

    $Login=$Credentials.UserName
    $Pass=$Credentials.Password # Get clear password
    Write-Debug -Message "Login = ""$Login"""

    # Code Here
    $Value = "Hello World"

    Return $Value

}

function Test-Function-2 {

    param (
        [Parameter(mandatory=$true)][string]$Message
    )

    Write-Host $Message

}
#Endregion

#Region User-defined parameters
# Reference to Test Module
$ProjectPath = "C:\Users\John\Documents\Projects\Test-Project"
$ModuleFileFullPath = "$ProjectPath\Source Code\Test-Module.psm1"

# Credentials
$KeyFile = "$ProjectPath\Configuration Files\secret.key"
$PasswordFile = "$ProjectPath\Configuration Files\password.txt"
Export-EncryptedSecureString -KeyFile $KeyFile -PasswordFile $PasswordFile # Run once : (re)Generate Credentials
$TestUserName="John"
$TestPassword = Get-EncryptedPassword -KeyFile $KeyFile -PasswordFile $PasswordFile

# Other parameters
$DebugPreference="Continue"
#Endregion

#Region Call function
# (re)Import Test Module
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath

# Credential object
$TestCredentials = new-object -typename System.Management.Automation.PSCredential -argumentlist $TestUserName, $TestPassword

# Call Test functions
$Test = Test-Function-1 -Credentials $TestCredentials
Test-Function-2 $Test
#Endregion
