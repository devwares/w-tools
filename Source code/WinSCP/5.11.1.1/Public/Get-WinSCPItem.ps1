function Get-WinSCPItem {

    [CmdletBinding(
        HelpUri = "https://github.com/dotps1/WinSCP/wiki/Get-WinSCPItem",
        PositionalBinding = $false
    )]
    [OutputType(
        [WinSCP.RemoteFileInfo]
    )]
    
    param (
        
        [Parameter(
            Mandatory = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateScript({ 
            if ($_.Opened) { 
                return $true 
            } else { 
                throw "The WinSCP Session is not in an Open state."
            }
        })]
        [WinSCP.Session]
        $WinSCPSession,

        [Parameter(
            Mandatory = $true,
            Position = 0,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [String[]]
        $Path,

        [Parameter()]
        [String]
        $Filter = [String]::Empty
    )

    process {
        foreach ($pathValue in (Format-WinSCPPathString -Path $($Path))) {
            $pathExists = Test-WinSCPPath -WinSCPSession $WinSCPSession -Path $pathValue
            if (-not $pathExists) {
                Write-Error -Message "Cannot find path '$pathValue' because it does not exist."
                continue
            }

            $filterParameterUsed = $PSBoundParameters.ContainsKey(
                "Filter"
            )

            if ($filterParameterUsed) {
                $output = Get-WinSCPChildItem -WinSCPSession $WinSCPSession -Path $pathValue -Filter $Filter
            } else {
                try {
                    $output = $WinSCPSession.GetFileInfo(
                        $pathValue
                    )
                } catch {
                    $PSCmdlet.WriteError(
                        $_
                    )
                    continue
                }
            }

            Write-Output -InputObject $output
        }
    }
}
