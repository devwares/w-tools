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
