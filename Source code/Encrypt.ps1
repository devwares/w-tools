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