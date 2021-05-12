function Get-RandomCharacters($length, $characters) { 
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length } 
    $private:ofs="" 
    return [String]$characters[$random]
}

function Switch-Characters([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}

function Get-RandomPassword32(){     
    $password = Get-RandomCharacters -length 20 -characters 'abcdefghiklmnoprstuvwxyz'
    $password += Get-RandomCharacters -length 4 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
    $password += Get-RandomCharacters -length 4 -characters '1234567890'
    $password += Get-RandomCharacters -length 4 -characters '!"§$%&/()=?}][{@#*+'
    $password = Scramble-String($password)
    return $password
}

function Export-Pfx(){  
    param
    (
        [Parameter(Mandatory=$true)] [string] $Dnsname,
        [Parameter(Mandatory=$true)] [string] $Filepath,
        [Parameter(Mandatory=$false)] [Security.SecureString] $Password
    )
 
    Try {
        $cert = New-SelfSignedCertificate -certstorelocation cert:\localmachine\my -dnsname $Dnsname
        $path = 'cert:\localMachine\my\' + $cert.thumbprint
        if([String]::IsNullOrEmpty($Password)) {
            $Password = ConvertTo-SecureString -String (Get-RandomPassword32) -Force -AsPlainText
        }
        Export-PfxCertificate -cert $path -FilePath $Filepath -Password $Password      
    }
    Catch {
        write-host -f Red "Error Downloading File!" $_.Exception.Message
    }
}

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

function Get-EncryptedSecureString {

    Param (

        [Parameter(mandatory=$true)][string]$KeyFile,
        [Parameter(mandatory=$true)][string]$StringFile

    )

    $Key = Get-Content $KeyFile
    $EncryptedPassword = Get-Content $StringFile
    Return ($EncryptedPassword | ConvertTo-SecureString -Key $Key)

}

function Convert-SecureStringToPlainText 
{

    param(
        [Parameter(Mandatory=$true)] [Securestring] $SecureString
    )

    # Get unsecure client_secret
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    $PlainText = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    Return $PlainText

}

function Export-OaepSHA1EncryptedStringToBase64() {

    param
    (
        [Parameter(Mandatory=$true)] [string] $PublicCertificateFilePath,
        [Parameter(Mandatory=$true,ParameterSetName="Clear")] [string] $PlainText,
        [Parameter(Mandatory=$true,ParameterSetName="Secure")] [securestring] $SecureString
    )

    try {
        If ($SecureString) {$PlainText = Convert-SecureStringToPlainText $SecureString}
        $PublicCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PublicCertificateFilePath)
        $ByteArray = [System.Text.ASCIIEncoding]::ASCII.GetBytes($PlainText)
        $EncryptedByteArray = $PublicCert.PublicKey.Key.Encrypt($ByteArray, [System.Security.Cryptography.RSAEncryptionPadding]::OaepSHA1)  
        $EncryptedBase64String = [Convert]::ToBase64String($EncryptedByteArray)
    }
    catch{
        Throw $_.exception
    }

    Return $EncryptedBase64String

}

Function Get-MD5Hash {

    param
    (
        [Parameter(Mandatory=$true,ParameterSetName="File")] [string] $Filepath,
        [Parameter(Mandatory=$true,ParameterSetName="Bytearray")] [byte] $Bytearray
    )

    try {
        Write-Verbose "New-object MD5CryptoServiceProvider"
        $md5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
        If ([string]::IsNullOrEmpty($Filepath)) {
            Write-Verbose "Compute Hash for Bytearray"
            return [System.BitConverter]::ToString($md5.ComputeHash($Bytearray))
        }else {
            If (Test-Path -path $Filepath){
                Write-Verbose "Compute Hash for File $Filepath"
                return [System.BitConverter]::ToString($md5.ComputeHash([System.IO.File]::ReadAllBytes($Filepath)))
            }
            else {
                Throw [System.Management.Automation.ItemNotFoundException]
            }
        }
    }
    catch {
        Throw $_.exception
    }

}

Function Encrypt-StringWithPublicCertificate {

    [CmdletBinding()]
    [OutputType([System.String])]
    param(
        [Parameter(Position=0, Mandatory=$true)][ValidateNotNullOrEmpty()][System.String]
        $ClearText,
        [Parameter(Position=1, Mandatory=$true)][ValidateNotNullOrEmpty()][ValidateScript({Test-Path $_ -PathType Leaf})][System.String]
        $PublicCertFilePath
    )

    # Encrypts a string with a public key
    $PublicCert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($PublicCertFilePath)
    $ByteArray = [System.Text.Encoding]::UTF8.GetBytes($ClearText)
    $EncryptedByteArray = $PublicCert.PublicKey.Key.Encrypt($ByteArray,$true)
    $EncryptedBase64String = [Convert]::ToBase64String($EncryptedByteArray)
    
    Return $EncryptedBase64String 

} 