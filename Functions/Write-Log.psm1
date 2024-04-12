<#
.SYNOPSIS
    Creates a log file and writes log messages to it.

.DESCRIPTION
    This script creates a log directory if it doesn't exist and defines the log file path based on the current date and time. 
    It also contains a function called Write-Log that takes a log message as input and writes it to the log file.

.PARAMETER Message
    The log message to be written to the log file.

.EXAMPLE
    Write-Log -Message "This is a sample log message"

    This example demonstrates how to use the Write-Log function to write a log message to the log file.

.NOTES
    Author: Phillip Anderson
    Date:   10/04/2024
#>
<#
.SYNOPSIS
    Creates a log file and writes log messages to it.

.DESCRIPTION
    This script creates a log directory if it doesn't exist and defines the log file path based on the current date and time. 
    It also contains a function called Write-Log that takes a log message as input and writes it to the log file.

.PARAMETER Message
    The log message to be written to the log file.

.EXAMPLE
    Write-Log -Message "This is a sample log message"

    This example demonstrates how to use the Write-Log function to write a log message to the log file.

.NOTES
    Author: Phillip Anderson
    Date:   10/04/2024
#>
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,
        
        [Parameter(Mandatory = $false)]
        [ValidateSet("INFO", "WARNING", "ERROR")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory = $false)]
        [string]$LogDirectory = "C:\Support"
    )
    
    try {
        if (-not (Test-Path -Path $LogDirectory)) {
            New-Item -ItemType Directory -Path $LogDirectory | Out-Null
        }

        $logFile = Join-Path -Path $LogDirectory -ChildPath "$(Get-Date -Format 'dd.MM.yyyy.HH.mm.ss')-log.txt"
        $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
        $logMessage = "$timestamp - $Level - $Message"
        
        Add-Content -Path $logFile -Value $logMessage
        Write-Output $logMessage
    } catch {
        Write-Error "Failed to write log: $_"
    }
}

#Export-ModuleMember -Function Write-Log

 # End of Write-Log function
