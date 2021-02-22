#https://docs.microsoft.com/fr-fr/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.1
function List-Exceptions {

    <# Common Exception Classes
    SystemException	: A failed run-time check;used as a base class for other.
    AccessException	: Failure to access a type member, such as a method or field.
    ArgumentException : An argument to a method was invalid.
    ArgumentNullException :	A null argument was passed to a method that doesn't accept it.
    ArgumentOutOfRangeException : Argument value is out of range.
    ArithmeticException	: Arithmetic over - or underflow has occurred.
    ArrayTypeMismatchException : Attempt to store the wrong type of object in an array.
    BadImageFormatException	: Image is in the wrong format.
    CoreException : Base class for exceptions thrown by the runtime.
    DivideByZeroException : An attempt was made to divide by zero.
    FormatException  :The format of an argument is wrong.
    IndexOutOfRangeException : An array index is out of bounds.
    InvalidCastExpression : An attempt was made to cast to an invalid class.
    InvalidOperationException : A method was called at an invalid time.
    MissingMemberException : An invalid version of a DLL was accessed.
    NotFiniteNumberException : A number is not valid.
    NotSupportedException : Indicates sthat a method is not implemented by a class.
    NullReferenceException : Attempt to use an unassigned reference.
    OutOfMemoryException : Not enough memory to continue execution.
    StackOverflowException : A stack has overflown.
    #>

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
