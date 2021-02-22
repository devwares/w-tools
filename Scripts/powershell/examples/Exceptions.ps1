#https://docs.microsoft.com/fr-fr/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.1
function List-Exceptions {
    [appdomain]::CurrentDomain.GetAssemblies() | ForEach {
        Try {
            $_.GetExportedTypes() | Where {
                $_.Fullname -match 'Exception'
            }
        } Catch {}
    } | Select FullName  | Sort-Object -Property FullName
}

Function Test-Exceptions {

    try {

        Write-Host "Function Top Level Try"

        try {

            Write-Host "Function Nested Try"
            #Invoke-Expression "dir c:\" | Out-Null # No Exception
            #Invoke-Expression "command that doesn't exist" | Out-Null # ParseException
            #Invoke-Expression "command that does not exist" | Out-Null # CommandNotFoundException
            #Throw "Random error" # Unknown type

        }
        catch [System.Management.Automation.ParseException] {

            Write-Warning "ParseException ! Transmitting."
            throw $_.Exception

        }
        catch {
            Write-Warning "Oops... unknown error ! Transmitting."
            throw $_.Exception
        }
        finally {

            Write-Host """Finally"" block is always executed"

        }

        Write-Host "This message is not displayed when exception is thrown"

    }
    Catch {
        Write-Warning "Transmitting too !"
        Throw $_.Exception
    }

}

Try {

    Test-Exceptions

}
Catch {

    Write-Error "Test Function has thrown an error"
    Write-Warning "I am still running instructions yet :)"
    Write-Host "Exception type :" $_.Exception.gettype()

}
Finally
{

    Write-host "Final instructions are always displayed."

}

Write-host "Program terminated. This message will always be displayed too." 
