

function Send-Winscp() {
    param
    (
        [Parameter(Mandatory=$true)] [string] $Server,
        [Parameter(Mandatory=$true)] [string] $Port,
        [Parameter(Mandatory=$true)] [Security.SecureString] $Password,
        [Parameter(Mandatory=$true)] [string] $Fingerprint,
        [Parameter(Mandatory=$true)] [string] $Destination,
        [Parameter(Mandatory=$true)] [string] $Username        
    )
    Import-Module $PSScriptRoot\WinSCP
}

function Receive-Winscp(){
    param
    (
        [Parameter(Mandatory=$true)] [string] $Server,
        [Parameter(Mandatory=$true)] [string] $Port,
        [Parameter(Mandatory=$true)] [Security.SecureString] $SecurePassword,
        [Parameter(Mandatory=$true)] [string] $Fingerprint,
        [Parameter(Mandatory=$true)] [string] $Path,
        [Parameter(Mandatory=$true)] [string] $FileMask,
        [Parameter(Mandatory=$true)] [string] $Destination,
        [Parameter(Mandatory=$true)] [string] $Username
    )

    Import-Module $PSScriptRoot\WinSCP

    Try {

        Write-Host "DEBUG", "Receive-Winscp", "Server: $Server"
        Write-Host "DEBUG", "Receive-Winscp", "Port: $Port"
        Write-Host "DEBUG", "Receive-Winscp", "Username: $Username"
        Write-Host "DEBUG", "Receive-Winscp", "Destination: $Destination"

        $TotalNumberOfFileFound = 0

        # Create PSCredential
        $PSCredential = [System.Management.Automation.PSCredential]::new($Username,$SecurePassword)

        Try {
            # Create SFTPSession
            Write-Host "DEBUG", "Receive-Winscp", "Establishing session $($Server):$Port with user $Username"
            $SFTPSessionOption = New-WinSCPSessionOption -Credential $PSCredential -HostName $Server -PortNumber $Port -SshHostKeyFingerprint $Fingerprint
            $SFTPSession = New-WinSCPSession -SessionOption $SFTPSessionOption
            Write-Host "DEBUG", "Receive-Winscp", "Session etablished $($Server):$Port with user $Username"   
        } Catch {

            Throw "Could not establish SFTP session --> $($_.Exception.Message)"
        }

        # Watch 
        Write-Host "DEBUG", "Receive-Winscp", "Listing all files"
        $AllFiles = Get-WinSCPChildItem -WinSCPSession $SFTPSession -Path $Path -File
        $MatchingFiles = $AllFiles | Where-Object {$_.Name -like $FileMask}
        Write-Host "DEBUG", "Receive-Winscp", "$(($MatchingFiles | Measure-Object).count) files match the pattern ' $FileMask '"

        if($null -ne $MatchingFiles) {
        
            # Set the transfert options
            $TransfertOptions = New-WinSCPTransferOption -OverwriteMode $([WinSCP.OverwriteMode]::Overwrite) -TransferMode Binary
    
            # Log
            Write-Host "DEBUG", "Receive-Winscp", "Transfert option set to BINARY"
    
            $LastOrder = 0
            # Download file to temporary location
    
            foreach($MatchingFile in $MatchingFiles) {
                
                # Parameters
                $Now = Get-Date
                $LastOrder = $LastOrder +1
                $FileName = $MatchingFile.Name
                $NewFilepath = "$Destination\$FileName"
    
                # Log
                Write-Host "DEBUG", "Receive-Winscp", "Download file $($MatchingFile.FullName)"
                
                # Download
                Receive-WinSCPItem -WinSCPSession $SFTPSession -RemotePath $MatchingFile.FullName -LocalPath $Destination -Remove -TransferOptions $TransfertOptions    
            }
    
            #$FileFound = $True
            $TotalNumberOfFileFound += $MatchingFiles.count

        } else {
    
            Write-Host "WARNING", "Receive-Winscp", "No file matching file pattern found"
        }
    } Catch {

        # Failed
        Write-Host "WARNING", "Receive-Winscp", "$($_.Exception.Message)"
        Throw $($_.Exception.Message)

    }

    return $TotalNumberOfFileFound

}