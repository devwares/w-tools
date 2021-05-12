<# CHECKLIST
- new object with wrong commandline/id
- getvalue right after killing process (once mode)
- call StopWatching on "forever" CimInstanceWatcher
- call StopWatching on "until" CimInstanceWatcher
- call StopWatching on "for" CimInstanceWatcher
- set runtime with bad date format
- call startwatching on terminated process
- call startwatching after setruntime at past date

#>

$DebugPreference="Continue"

########### Get test process infos
$MatchCmdLine = 'Teams.exe" --system-initiated'
$process_id = (Get-CimInstance -ClassName Win32_Process | Where-Object{$_.CommandLine -match $MatchCmdLine}).ProcessId
$process_cmdline = (Get-CimInstance -ClassName Win32_Process | Where-Object{$_.CommandLine -match $MatchCmdLine}).CommandLine

########### (re)Import Module
$ProjectPath = "C:\Users\john\Documents\Projects\w-tools"
$ModuleFileFullPath = "$ProjectPath\Source code\Wmi.ps1"
$ModuleName = [io.path]::GetFileNameWithoutExtension($ModuleFileFullPath)
Get-Module | Where-Object -Property "Name" -Like "*$ModuleName*" | Remove-Module
Import-Module $ModuleFileFullPath

try {

    ###################### Init CimInstanceWatcher
    Write-Host "*********** INIT CimInstanceWatcher ***********" -ForegroundColor Cyan    
    #$runtime = "forever"
    #$runtime = "once"
    #$runtime = "Until 2021-03-23T15:12:30" # yyyy-MM-ddTHH:mm:ss
    $runtime = "for 10"
    $frequency = 2

    $Pw = New-Object CimInstanceWatcher($process_cmdline)
    $Pw.SetRuntime($runtime)
    $Pw.SetFrequency($frequency)
    $Pw.SetSizeInRamLimit(100000000)
    
    # Test GetLastValue before populate
    #Write-host "GetLastValue ""pagefile"" before FetchData : $($Pw.GetLastValue("win32_process\PageFileUsage"))"

    # Fetch one report
    #$Pw.FetchData()

    # Test GetLastValue after populate
    #Write-host "GetLastValue ""pagefile"" after FetchData : $($Pw.GetLastValue("win32_process\PageFileUsage"))"
    
    # Start Watching (loop on FetchData)
    $Pw.Start()

    # Show some data
    Write-host "Instant ""win32_process\pagefile"": $($Pw.GetLastValue("win32_process\PageFileUsage"))"
    Write-host "Average ""win32_process\pagefile"": $($Pw.GetAverageValue("win32_process\PageFileUsage"))"
    Write-host "Instant ""Win32_PerfFormattedData_PerfProc_Process\WorkingSetPeak"": $($Pw.GetLastValue("Win32_PerfFormattedData_PerfProc_Process\WorkingSetPeak"))"
    Write-host "Average ""Win32_PerfFormattedData_PerfProc_Process\WorkingSet"": $($Pw.GetAverageValue("Win32_PerfFormattedData_PerfProc_Process\WorkingSet"))"
    Write-host "Instant ""Win32_PerfFormattedData_PerfProc_Process\PercentProcessorTime"": $($Pw.GetLastValue("Win32_PerfFormattedData_PerfProc_Process\PercentProcessorTime"))"
    Write-host "Average ""Win32_PerfFormattedData_PerfProc_Process\PercentProcessorTime"": $($Pw.GetAverageValue("Win32_PerfFormattedData_PerfProc_Process\PercentProcessorTime"))"
    Write-host "Average ""Custom\ProcessWatcherPercentProcessorTime"": $($Pw.GetAverageValue("Custom\ProcessWatcherPercentProcessorTime"))"
    Write-host "Average ""Custom\ProcessCpuPercentTime"": $($Pw.GetAverageValue("Custom\ProcessCpuPercentTime"))"
    Write-host "Max ""Custom\ProcessWatcherWorkingSet"": $($Pw.GetMaxValue("Custom\ProcessWatcherWorkingSet"))"

}
catch {
    Write-Error $_.Exception.message
}