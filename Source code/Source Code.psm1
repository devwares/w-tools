# Import functions
Get-ChildItem -Path $PSScriptRoot\*.ps1 | Foreach-Object{ . $_.FullName }

# Registers functions
Export-ModuleMember -Function * -Alias *