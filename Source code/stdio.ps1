<#
.SYNOPSIS
Get file encoding

.DESCRIPTION
Get encoding format for specified file

.PARAMETER Path
Full path to file to analyse

.EXAMPLE
Get-Encoding "c:\users\john\documents\notes.txt"

.OUTPUTS
PSCustomObject with Encoding and Path as members
#>
function Get-FileEncoding
{
  param
  (
    [Parameter(Mandatory,ValueFromPipeline,ValueFromPipelineByPropertyName)]
    [Alias('FullName')]
    [string]
    $Path
  )

  process 
  {
    $bom = New-Object -TypeName System.Byte[](4)
        
    $file = New-Object System.IO.FileStream($Path, 'Open', 'Read')
    
    $null = $file.Read($bom,0,4)
    $file.Close()
    $file.Dispose()
    
    $enc = [Text.Encoding]::ASCII
    if ($bom[0] -eq 0x2b -and $bom[1] -eq 0x2f -and $bom[2] -eq 0x76) 
      { $enc =  [Text.Encoding]::UTF7 }
    if ($bom[0] -eq 0xff -and $bom[1] -eq 0xfe) 
      { $enc =  [Text.Encoding]::Unicode }
    if ($bom[0] -eq 0xfe -and $bom[1] -eq 0xff) 
      { $enc =  [Text.Encoding]::BigEndianUnicode }
    if ($bom[0] -eq 0x00 -and $bom[1] -eq 0x00 -and $bom[2] -eq 0xfe -and $bom[3] -eq 0xff) 
      { $enc =  [Text.Encoding]::UTF32}
    if ($bom[0] -eq 0xef -and $bom[1] -eq 0xbb -and $bom[2] -eq 0xbf) 
      { $enc =  [Text.Encoding]::UTF8}
        
    [PSCustomObject]@{
      Encoding = $enc
      Path = $Path
    }
  }
}

<#
.SYNOPSIS
Get file data

.DESCRIPTION
Reads file and return Bytes Array with data

.PARAMETER FilePath
Full path to file

.EXAMPLE
Get-FileData -FilePath "C:\users\john\test.txt"

.OUTPUTS
Bytes Array with file data
#>
Function Get-FileData {

    param
    (
        [Parameter(Mandatory=$true)] [string]$FilePath
    )
    Write-Debug "[Get-FileData]: Get file data"
    [Byte[]]$FileData = Get-Content -Path $FilePath -Encoding Byte -ReadCount 0
    Return $FileData
}

function Read-String
{
    param
    (
        [Parameter(Mandatory=$false)][string] $Value,
        [Parameter(Mandatory=$false)][string] $Message,
        [Parameter(Mandatory=$false)][string] $Action
    )

    # Display Message if specified
    If (![string]::IsNullOrEmpty($Message)){
        $EnteredValue = Read-Host $Message
    }
    Else {
        $EnteredValue = Read-Host
    }

    # Run Action if action specified and specified value is entered (not case sensitive)
    If ( !([string]::IsNullOrEmpty($Value)) -and ($EnteredValue.ToUpper() -eq $Value.ToUpper()) -and !([string]::IsNullOrEmpty($Action)))
        {
            Invoke-Expression -Command $Action
        }

    # Returns EnteredValue
    return $EnteredValue

}

function Read-Module
{
    param
    (
        [Parameter(Mandatory=$true)][string] $ModuleFileFullPath
    )
    $ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
    Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
    Import-Module $ModuleFileFullPath
}
