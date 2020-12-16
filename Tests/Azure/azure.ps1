Function Test-Find-AzVms()
{
    Connect-AzAccount
    $idList = @("532rg642-af1b-58j7-8d3c-98658455540t")
    $ExportFilePath = "C:\users\john\documents\test.csv"

    # Store all Windows VMs in an array
    $AzVmList = Find-AzVms –SubscriptionList $IdList | Where-Object {$_.OsType -EQ "Windows"} 

    # List all VMs informations in a file
    Find-AzVms –SubscriptionList $IdList –ExportFilePath $ExportFilePath -Delimiter ";"
}