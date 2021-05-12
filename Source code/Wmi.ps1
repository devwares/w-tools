class Watcher{

    [int] $Id
    hidden static [int16] $WatcherCount

    Watcher() {
        $type = $this.GetType()

        if ($type -eq [Watcher])
        {
            throw("Class $type must be inherited")
        }
    }

    [void]Start(){
        throw("Must Override Method")
    }

    [void]Reset(){
        throw("Must Override Method")
    }

    [void]Stop(){
        throw("Must Override Method")
    }

    [string]ToString(){
        throw("Must Override Method")
    }

}

class CimInstanceWatcher : Watcher {

    #region Properties
    [int]$ProcessId
    [string]$CommandLine
    hidden [array]$WatchedClassesList = @("Win32_Process","Win32_PerfFormattedData_PerfProc_Process") # List of classes to include in reports
    hidden [string]$RuntimeMode # once, forever, until, for
    hidden [datetime]$RuntimeLimit # watching stops at this datetime
    hidden [int]$FrequencyMinValue = 0 # minimum number of seconds accepted for frequency
    hidden [int]$Frequency = 30 # stats fetch frequency in seconds (doesn't apply to "once" runtime mode)
    hidden [int]$Timer = 0 # Timer used for loops on getting informations at set frequency
    hidden [PSCustomObject]$CimClassesObject = @{} # debug : $CimInstancesArray
    hidden [string]$Version = "1.0"
    hidden [string]$WatchedHostName = "localhost"
    hidden [int]$SizeInRamLimit = 500000000 # Maximum size of object before size limit is reached
    #endregion

    #region Constructors
    CimInstanceWatcher() {
        Throw "ProcessId or Command-Line required"
    }

    CimInstanceWatcher([UInt32]$ProcessId) {
        $This.InitFromProcessId($ProcessId)
    }

    CimInstanceWatcher([string]$CommandLine) {   
        $This.InitFromCommandLine($CommandLine)
    }

    CimInstanceWatcher([UInt32]$ProcessId, [string]$Runtime) {
        $This.InitFromProcessId($ProcessId)
        $This.SetRuntime($Runtime)
        $This.Start() # Start watching
    }

    CimInstanceWatcher([string]$CommandLine, [string]$Runtime) {   
        $This.InitFromCommandLine($CommandLine)
        $This.SetRuntime($Runtime)
        $This.Start() # Start watching
    }

    CimInstanceWatcher([UInt32]$ProcessId, [string]$Runtime, [int]$Frequency) {
        $This.InitFromProcessId($ProcessId)
        $This.SetRuntime($Runtime)
        $This.SetFrequency($Frequency)
        $This.Start() # Start watching
    }

    CimInstanceWatcher([string]$CommandLine, [string]$Runtime, [int]$Frequency) {   
        $This.InitFromCommandLine($CommandLine)
        $This.SetRuntime($Runtime)
        $This.SetFrequency($Frequency)
        $This.Start() # Start watching
    }
    #endregion

    #region Setters
    [void]SetRuntime([string]$runtime){

        Switch ($runtime.split(" ")[0].ToLower()) {
            "once" { 
                $This.Runtimemode = "once"
                Write-Debug "Runtime set to ""Once"""
            }
            "forever" { 
                $This.Runtimemode = "forever"
                Write-Debug "Runtime set to ""Forever"". Watching will stop when process $($This.ProcessId) is terminated"   
            }
            "until" { $This.Runtimemode = "until"

                try {
                    $Limit = $runtime.split(" ")[1]
                    $This.RuntimeLimit = [System.DateTime]::ParseExact($Limit,'yyyy-MM-ddTHH:mm:ss',$null)
                    $This.RuntimeMode = "until"
                    Write-Debug "Runtime is set ""$($This.RuntimeMode) $($This.RuntimeLimit)"""
                }
                catch {
                    Throw "Unable to set time limit. Expected : ""UNTIL yyyy-MM-ddTHH:mm:ss"""
                }
            
            }
            "for" { $This.Runtimemode = "for"
                try {
                    [int]$Duration = $runtime.split(" ")[1].ToLower()
                    $Now = Get-Date
                    $This.RuntimeLimit = $Now.AddSeconds($Duration)
                    $This.RuntimeMode = "until"
                    Write-Debug "Runtime is set ""for $Duration seconds"". Watcher will stop at $($This.RuntimeLimit)"""
                }
                catch {
                    Throw "Unable to set time limit. Expected : ""FOR XX"" (where XX = number of seconds)"
                }
        
            }
            Default {Throw "Format unknown"}
        }

    }

    [void]SetSizeInRamLimit($Size){
        If ([CimInstanceWatcher]::IsNumeric($Size)) {$This.SizeInRamLimit = $Size}
        else {Throw "Incorrect format. Object size limit must be numeric"}
    }

    [void]SetFrequency([int]$seconds){
        If ($seconds -ge $This.FrequencyMinValue) {
            $This.Frequency = $seconds
            $This.Timer = 0 # reset timer for loop in watching method
            Write-Debug "Frequency set to $($This.Frequency)"
        }
        else {
            Throw "Minimum frequency value is $($This.FrequencyMinValue)"
        }
    }
    #endregion

    #region Getters
    [string]GetClassIdKey($ClassName) {
        # Define property name containing processId value for each allowed class
        switch ($ClassName.tolower()) {
            "win32_process" {return "ProcessId"}
            "win32_perfformatteddata_perfproc_process" { return "IDProcess" }
            Default {Throw "Class unknown : $ClassName"}
        }
        Throw "Class unknown : $ClassName"
    }

    [string]GetRuntime(){
        Switch ($This.RuntimeMode) {
            "once" { return "Once" }
            "forever" { return "No time limit (runs until process end)" }
            {"until","for"} { return "Will stop at $($This.RuntimeLimit)" }
            Default { return "Runtime not set"}
        }
        Return "0" # Alternate return must be set outside the switch for methods
    }

    [int]GetSizeInRamLimit(){
        Return $This.SizeInRamLimit
    }

    [string]GetFrequency(){
        Return "Frequency is $($This.Frequency) seconds"
    }

    [int]GetMaxValue($FullPropertyPath){

        # Check path to property and reports content
        try {
            $This.CheckReportPropertyPath($FullPropertyPath)
        }
        catch {
            $ErrorMsg = $_.exception.message
            If ($ErrorMsg -match "No report found for this class"){
                Write-Warning $ErrorMsg
                Return $null
             }else {
                Throw $ErrorMsg
            }
        }
        
        $ClassName = $FullPropertyPath.split("\")[0]
        $PropertyName = $FullPropertyPath.split("\")[1]

        $MaxValue = 0
        # Get values and type for Property
        $PropertyValues = $This.CimClassesObject["CLASSES"][$ClassName]["Reports"] | Select-Object -ExpandProperty $PropertyName
        If ([CimInstanceWatcher]::IsNumeric($PropertyValues[0])) {
            foreach ($PropertyValue in $PropertyValues){
                If ($PropertyValue -gt $MaxValue) {$MaxValue = $PropertyValue}
            }
        }
        else {
            Throw "Could not peak value for $($PropertyValues[0].GetType().Name) type"
        }
        return $MaxValue

    }

    [int]GetAverageValue($FullPropertyPath){

        # Check path to property and reports content
        try {
            $This.CheckReportPropertyPath($FullPropertyPath)
        }
        catch {
            $ErrorMsg = $_.exception.message
            If ($ErrorMsg -match "No report found for this class"){
                Write-Warning $ErrorMsg
                Return $null
             }else {
                Throw $ErrorMsg
            }
        }
        
        $ClassName = $FullPropertyPath.split("\")[0]
        $PropertyName = $FullPropertyPath.split("\")[1]

        $AverageValue = 0
        # Get values and type for Property
        $PropertyValues = $This.CimClassesObject["CLASSES"][$ClassName]["Reports"] | Select-Object -ExpandProperty $PropertyName
        If ([CimInstanceWatcher]::IsNumeric($PropertyValues[0])) {
            $ValuesTotal = 0
            foreach ($PropertyValue in $PropertyValues){
                $ValuesTotal += $PropertyValue
            }
            $AverageValue = $ValuesTotal / $PropertyValues.Count
        }
        else {
            Throw "Could not get average value for $($PropertyValues[0].GetType().Name) type"
        }
        return $AverageValue

    }

    [string]GetLastValue($FullPropertyPath){

        # Check path to property and reports content
        try {
            $This.CheckReportPropertyPath($FullPropertyPath)
        }
        catch {
            $ErrorMsg = $_.exception.message
            If ($ErrorMsg -match "No report found for this class"){
                Write-Warning $ErrorMsg
                Return $null
             }else {
                Throw $ErrorMsg
            }
        }
        
        $ClassName = $FullPropertyPath.split("\")[0]
        $PropertyName = $FullPropertyPath.split("\")[1]

        Return $This.CimClassesObject["CLASSES"][$ClassName]["Reports"][-1] | Select-Object -ExpandProperty $PropertyName

    }

    [object]GetReport(){Return $This.CimClassesObject}

    [string]GetReport($Format){
        $Report = $Null
        try {
            switch ($Format.ToLower()) {
                "json" {$Report = $This.CimClassesObject | ConvertTo-Json -Depth 4 | Out-String}
                Default {Write-Warning "Not supported yet : $($Format)"}
            }
        }
        catch {
            Throw $_.exception
        }
        Return $Report
    }

    #endregion

    #region Other methods
    hidden [void]InitFromCommandLine($CommandLine) {
        [CimInstanceWatcher]::WatcherCount++ # Increment count of watchers
        $This.Id = [CimInstanceWatcher]::WatcherCount # Current number of watchers is assigned to this as an ID
        $This.CommandLine = $CommandLine
        $This.ProcessId = (Get-CimInstance -ClassName Win32_Process | Where-Object{$_.CommandLine -eq $CommandLine}).ProcessId
        if ($This.ProcessId) {
            foreach ($id in $This.ProcessId){Write-Debug "Found Process Id : $id"}
        }
        Else {Throw "No process id found for this command line : ""$CommandLine"""}
        
        If (($This.ProcessId | Measure-Object).Count -gt 1) {
            Throw "Found multiple process ids with same command-line"
        }
        Else {$This.Reset()}
        
    }

    hidden [void]InitFromProcessId($ProcessId){
        [CimInstanceWatcher]::WatcherCount++ # Increment count of watchers
        $This.Id = [CimInstanceWatcher]::WatcherCount # Current number of watchers is assigned to this as an ID
        $This.ProcessId = $ProcessId
        $This.Reset() # Get informations for summary array
        $This.CommandLine = $This.CimInstance.CommandLine
        Write-Debug "CommandLine for process $ProcessId : $($This.CommandLine)"
    }

    [void]Reset(){
        # Custom Object
        $This.CimClassesObject["GENERAL"] = @{
            Version = $This.Version
            Hostname = $This.WatchedHostName
        }
        $This.CimClassesObject["CLASSES"] = @{}
        Foreach ($classname in $This.WatchedClassesList) {
            $This.CimClassesObject["CLASSES"][$classname] += @{
                ClassName = $classname
                Reports = @()
            }
        }
        $This.CimClassesObject["CLASSES"]["Custom"] += @{
            ClassName = "Custom"
            Reports = @()
        }
        Write-Debug "Reset data object"
    }

    [void]FetchData(){
        # Adds new report for every watched classes
        If (Get-Process -id $This.ProcessId -ErrorAction SilentlyContinue) {
            # Get informations for process
            foreach ($watchedclass in $This.WatchedClassesList) {
                try {
                    $key = $This.GetClassIdKey($watchedclass) # get name of property with ProcessId value
                    Write-Debug "Fetching data for $watchedclass..."
                    $report = Get-CimInstance -ClassName $watchedclass | Where-Object{$_.$key -eq $This.ProcessId} | Select-Object * # fetch data
                    $report | Add-Member -MemberType NoteProperty -name WatcherTimeStamp -Value (get-date -format 'yyyy-MM-ddTHH:mm:ss') # add timestamp to report
                    $This.CimClassesObject["CLASSES"][$watchedclass]["Reports"] += $report # adds report to report object
                    Write-Debug "Fetched data for $watchedclass at $(Get-Date) : $report"
                }
                catch {
                    Write-Warning "Could not fetch data for class $watchedclass"
                }
            }
            # Get additionnal informations about watched process and current PS process
            try {
                $report = @{}
                Write-Debug "Fetching additionnal data..."
                $report | Add-Member -MemberType NoteProperty -name WatcherTimeStamp -Value (get-date -format 'yyyy-MM-ddTHH:mm:ss') # add timestamp to report
                $report | Add-Member -MemberType NoteProperty -name TotalCpuLoadPercentage -Value (Get-WmiObject Win32_Processor | Measure-Object -Property LoadPercentage -Average | Select-Object Average).Average # add cpu load to report
                $memory = (Get-WmiObject Win32_OperatingSystem | Select-Object TotalVisibleMemorySize, FreePhysicalMemory) # Get physical memory stats
                $report | Add-Member -MemberType NoteProperty -name TotalMemoryLoad -Value ($Memory.TotalVisibleMemorySize - $Memory.FreePhysicalMemory) # memory usage
                $TotalCpuTime = (Get-CimInstance -ClassName Win32_PerfRawData_Counters_ProcessorInformation | Where-Object {$_.Name -eq "_Total"}).PercentProcessorTime
                $ProcessCpuTime = (Get-CimInstance -ClassName Win32_PerfRawData_PerfProc_Process | Where-Object {$_.IDProcess -eq $This.ProcessId}).PercentProcessorTime
                $report | Add-Member -MemberType NoteProperty -name ProcessCpuPercentTime -Value ((100 / $TotalCpuTime) * $ProcessCpuTime) # precise cpu process %
                $PwPercentProcessorTime =  (Get-CimInstance -ClassName Win32_PerfRawData_PerfProc_Process | Where-Object {$_.IDProcess -eq $global:pid}).PercentProcessorTime
                $report | Add-Member -MemberType NoteProperty -name ProcessWatcherPercentProcessorTime -Value ((100 / $TotalCpuTime) * $PwPercentProcessorTime) # precise cpu process % for This Process Watcher instance (powershell.exe)
                $PwWorkingSet =  (Get-CimInstance -ClassName Win32_PerfFormattedData_PerfProc_Process | Where-Object {$_.IDProcess -eq $global:pid}).WorkingSet
                $report | Add-Member -MemberType NoteProperty -name ProcessWatcherWorkingSet -Value $PwWorkingSet # working set (memory) for This Process Watcher instance (powershell.exe)
                $This.CimClassesObject["CLASSES"]["Custom"]["Reports"] += $report # adds report to report object
            }
            catch {
                Write-Warning "Could not fetch additional data for process"
            }
            Write-Debug "End of fetch data"
        }
        else {
            Throw "Fetch data failed (process not found $($This.ProcessId))"
        }
    }

    [void]Start(){

        If ($This.RuntimeMode) {

            If ($This.RuntimeMode -eq "once") { Write-Warning "Watcher could not be started in ""once"" runtime mode"}
            else {
                
                Write-Debug "Watcher started at $(get-date)"
                $This.Timer = $This.Frequency # first request will be triggered immediately
                $RequestExecutionTime = 0
                $Continue = $True

                # Watch at set frequency until stop condition is met (method, process terminated, timelimit reached)
                Do{

                    # Put informations into Reports when Frequency is reached by Timer
                    If ($This.Timer -ge $This.Frequency) {
                        try {
                            Write-Debug "Requesting process informations at $(Get-Date)"
                            $RequestExecutionTime = [math]::Round((Measure-Command {$This.FetchData()}).TotalSeconds) # measure time while fetching data
                            $This.Timer = 0 # Reset timer for next informations request
                        }
                        catch {
                            Write-Warning "No answer from process $($this.processid)"
                        }
                    }else {
                        $RequestExecutionTime = 0
                    }
                    
                    # Adds request execution time to Timer (this time has to be substracted to time-to-wait so as to respect frequency time)
                    $This.Timer = $This.Timer + 1 + $RequestExecutionTime
                    Start-Sleep 1

                    # Check stop condition depending on set runtime mode
                    If ($This.RuntimeMode -eq "forever") {
                        $Continue = Get-Process -id $This.ProcessId -ErrorAction SilentlyContinue }
                    Else{
                        $Continue = (Get-Date) -lt $This.RuntimeLimit
                    }
                
                    Write-Debug "Timer : $($This.Timer) seconds"
                    Write-Debug "*This* Process size in RAM :   $((Get-Process -id $global:pid).WS)"
                    Write-Debug "Limit size in RAM :            $($This.SizeInRamLimit)"

                }While($Continue -and !$This.SizeInRamLimitReached())

            }

        }
        else {
            Throw "Runtime must be set before you start watching process"
        }

    }

    [void]Stop(){
        Write-Warning "Not implemented yet"
    }

    hidden [boolean]CheckReportPropertyPath([string]$FullPropertyPath) {
        # Check Format
        $count = ($FullPropertyPath.ToCharArray() | Where-Object {$_ -eq '\'} | Measure-Object).Count
        if ($count -ne 1) {
            Throw "Invalid format. Expected : Classname\PropertyName (e.g : ""Win32_Process\PageFileUsage)"""
        }
        # Split property path
        $ClassName = $FullPropertyPath.split("\")[0]
        $PropertyName = $FullPropertyPath.split("\")[1]
        # Check if at least one report exists, else returns $null
        if (($This.CimClassesObject["CLASSES"][$ClassName]["Reports"]).Count -eq 0) {
            Throw "No report found for this class : $ClassName"
        }
        # Check if property exists
        if ($null -eq ($This.CimClassesObject["CLASSES"][$ClassName]["Reports"] | Get-Member $PropertyName)) {
            Throw "Unknown property : $PropertyName"
        }
        Return $True
    }

    static Hidden [bool]IsNumeric($Value){
        $typeValue = $Value.getTypeCode().value__
        if ($typeValue -ge 5 -and $typeValue -le 15) {return $True}
        else {return $False}
    }
    
    Hidden [bool]SizeInRamLimitReached(){
        # Measure size of , return True if size > $This.SizeInRamLimit
        $CurrentSize = (Get-Process -id $global:pid).WS
        If ($CurrentSize -gt $This.SizeInRamLimit) {
            Return $True
        }
        Return $False
    }

    [string]ToString(){
        return "Id = $($this.Id) | ProcessId = $($this.ProcessId) | Version = $($this.Version)"
    }
    #endregion

}