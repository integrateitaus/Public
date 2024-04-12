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

# Create the log directory if it doesn't exist
$logDirectory = "C:\Support"
if (-not (Test-Path -Path $logDirectory)) {
    New-Item -ItemType Directory -Path $logDirectory | Out-Null
}

# Define the log file path
$logFile = Join-Path -Path $logDirectory -ChildPath "$(Get-Date -Format 'dd.MM.yyyy.HH.mm.ss')-log.txt"

# Function to write log messages
function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message
    )
    
    $timestamp = Get-Date -Format "dd-MM-yyyy HH:mm:ss"
    $logMessage = "$timestamp - $Message"
    
    Add-Content -Path $logFile -Value $logMessage
    Write-Output $logMessage
}

 # End of Write-Log function