# Copyright (c) 2025 Broadcom. All Rights Reserved.
# Broadcom Confidential. The term "Broadcom" refers to Broadcom Inc.
# and/or its subsidiaries.
#
# =============================================================================
#
# SOFTWARE LICENSE AGREEMENT
#
#
# Copyright (c) CA, Inc. All rights reserved.
#
#
# You are hereby granted a non-exclusive, worldwide, royalty-free license
# under CA, Inc.'s copyrights to use, copy, modify, and distribute this
# software in source code or binary form for use in connection with CA, Inc.
# products.
#
#
# This copyright notice shall be included in all copies or substantial
# portions of the software.
#
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
# FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
#
# =============================================================================
#
#
# VCF.PS.Toolbox - Utility Functions Module
#
# This module provides essential utility functions for the VCF PowerShell Toolbox,
# including logging, environment setup, timing operations, JSON processing, file
# operations, network validation, and user interface helpers. These functions are
# designed to be reusable across multiple VCF automation scripts and provide
# consistent error handling and logging capabilities.
#
# Key Features:
# - Multi-level logging with color-coded console output and file logging
# - Log level filtering (DEBUG, INFO, ADVISORY, WARNING, EXCEPTION, ERROR)
# - Standardized error handling with structured error result objects
# - Environment information gathering for troubleshooting
# - High-precision operation timing and performance measurement
# - Safe JSON file parsing with comprehensive error handling
# - JSON validation (missing properties, null values, file validation)
# - Interactive user input collection with validation
# - Configurable yes/no choice menus for user confirmation
# - Array validation for missing properties in configuration objects
# - File operations (existence, locking, disk space validation)
# - Network utilities (IP address validation, CIDR range checking)
# - Exception handling with detailed inner exception traversal
# - Secure string conversion for API authentication
#
# Last modified: 2025-01-30
#
Function Test-LogLevel {

    <#
        .SYNOPSIS
        Determines if a message should be displayed based on the configured log level.

        .DESCRIPTION
        Compares the message type against the configured log level threshold to determine
        if the message should be displayed on screen. All messages are always written to
        the log file regardless of level.

        The log level hierarchy from lowest to highest is:
        DEBUG < INFO < ADVISORY < WARNING < EXCEPTION < ERROR

        .PARAMETER messageType
        The type/severity of the log message to check.

        .PARAMETER configuredLevel
        The minimum log level configured for screen output.

        .EXAMPLE
        Test-LogLevel -messageType "DEBUG" -configuredLevel "INFO"
        Returns $false because DEBUG is below INFO threshold.

        .EXAMPLE
        Test-LogLevel -messageType "ERROR" -configuredLevel "INFO"
        Returns $true because ERROR is at or above INFO threshold.

        .OUTPUTS
        Boolean
        Returns $true if the message should be displayed, $false otherwise.

    #>
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$ConfiguredLevel,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$MessageType
    )

    $messageLevel = $Script:logLevelHierarchy[$MessageType]
    $configuredLevelValue = $Script:logLevelHierarchy[$ConfiguredLevel]

    return ($messageLevel -ge $configuredLevelValue)
}
Function Write-ErrorAndReturn {

    <#
        .SYNOPSIS
        Writes an error message and returns a standardized error result.

        .DESCRIPTION
        This function provides a standardized way to handle errors by logging the error
        message and returning a consistent error result object. This replaces the need
        for throw statements and provides better error handling consistency.

        USAGE GUIDELINES:
        - Use in Helper/Validation/Utility functions (not main workflow functions)
        - Allows caller to decide how to handle the error (propagate, retry, or exit)
        - Always check the returned Success property in the caller

        Error Handling Pattern:
        1. Helper function calls Write-ErrorAndReturn to return structured error
        2. Caller checks $result.Success
        3. Caller decides: propagate error, retry operation, or exit script

        .PARAMETER ErrorMessage
        The error message to log and include in the result.

        .PARAMETER ErrorCode
        Optional error code for categorization. Defaults to "ERR_UNKNOWN".

        Error Code Categories:
        - Connection Errors (1xxx):
          ERR_NOT_CONNECTED_SDDC, ERR_NOT_CONNECTED_VCENTER, ERR_CONNECTION_TIMEOUT,
          ERR_CONNECTION_FAILED, ERR_AUTH_FAILED, ERR_TOKEN_EXPIRED

        - Validation Errors (2xxx):
          ERR_INVALID_PARAMETER, ERR_INVALID_JSON, ERR_MISSING_PARAMETER,
          ERR_FILE_NOT_FOUND, ERR_INVALID_CREDENTIALS, ERR_VALIDATION_FAILED

        - Resource Errors (3xxx):
          ERR_CLUSTER_NOT_FOUND, ERR_IMAGE_NOT_FOUND, ERR_DOMAIN_NOT_FOUND,
          ERR_RESOURCE_NOT_FOUND, ERR_VCENTER_NOT_FOUND, ERR_HOST_NOT_FOUND

        - Operation Errors (4xxx):
          ERR_COMPLIANCE_CHECK_FAILED, ERR_TRANSITION_FAILED, ERR_IMPORT_FAILED,
          ERR_DELETE_FAILED, ERR_TASK_FAILED, ERR_OPERATION_FAILED

        - Task Errors (5xxx):
          ERR_TASK_IN_PROGRESS, ERR_TASK_CANCELLED, ERR_TASK_TIMEOUT,
          ERR_TASK_UNKNOWN_STATE, ERR_RETRY_FAILED

        - JSON/Configuration Errors (6xxx):
          ERR_JSON_PARSE, ERR_JSON_FORMAT, ERR_CONFIG_INVALID,
          ERR_REMEDIATION_OPTIONS_INVALID

        .EXAMPLE
        # Helper function returns error object
        Function Get-ClusterInfo {
            if (-not $cluster) {
                return Write-ErrorAndReturn `
                    -errorMessage "Cluster '$clusterName' not found in workload domain '$workloadDomainName'" `
                    -errorCode "ERR_CLUSTER_NOT_FOUND"
            }
            return @{ Success = $true; Cluster = $cluster }
        }

        .EXAMPLE
        # Caller checks result and decides how to handle
        $result = Get-ClusterInfo -clusterName $clusterName -workloadDomainName $workloadDomainName
        if (-not $result.Success) {
            Write-LogMessage -type ERROR -message "Failed to get cluster info: $($result.ErrorMessage)"
            exit 1  # Main workflow decides to exit
        }
        $cluster = $result.Cluster

        .OUTPUTS
        PSCustomObject
        Returns a hashtable with Success=$false, ErrorMessage, and ErrorCode properties.

        .NOTES
        Error Handling: This is a utility function used by helper/validation functions to return
        standardized error objects. Do NOT use 'exit 1' in helper functions; use this
        function instead to allow the caller to control error handling.

    #>
    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String]$ErrorCode = "ERR_UNKNOWN",
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$ErrorMessage
    )

    Write-LogMessage -type ERROR -message $ErrorMessage

    return @{
        Success = $false
        ErrorMessage = $ErrorMessage
        ErrorCode = $ErrorCode
    }
}
Function Exit-WithCode {

    <#
        .SYNOPSIS
        Exits the script with a standardized exit code and optional final message.

        .DESCRIPTION
        This function provides a centralized exit point that ensures consistent exit code usage,
        optional cleanup operations, and clear logging before script termination. Using this
        function instead of direct 'exit' calls improves automation integration and debugging.

        Benefits of standardized exit codes:
        - CI/CD pipelines can distinguish between failure types and implement appropriate retry logic
        - Monitoring systems can categorize failures for better alerting and reporting
        - Debugging is faster with clear failure category indication
        - Follows PowerShell and Unix conventions for exit codes

        Exit Code Categories (see $Script:ExitCodes):
        0  - SUCCESS: Operation completed successfully
        1  - GENERAL_ERROR: Unspecified error
        2  - PARAMETER_ERROR: Invalid parameters or validation failure
        3  - CONNECTION_ERROR: Failed to connect to SDDC Manager or vCenter
        4  - AUTHENTICATION_ERROR: Authentication or credential failure
        5  - RESOURCE_NOT_FOUND: Cluster, host, workload domain, or image not found
        6  - OPERATION_FAILED: Operation (transition, import, compliance) failed
        7  - TASK_FAILED: Background task failed or timed out
        8  - CONFIGURATION_ERROR: JSON or configuration file error
        9  - PRECONDITION_ERROR: Prerequisites not met (modules, versions)
        10 - USER_CANCELLED: User cancelled the operation

        .PARAMETER ExitCode
        The exit code to return to the shell. Use values from $Script:ExitCodes hashtable
        for consistency and self-documentation.

        .PARAMETER Message
        Optional final message to log before exiting. If ExitCode is 0, logs as INFO.
        Otherwise logs as ERROR.

        .PARAMETER NoCleanup
        Skip optional cleanup operations before exit. Use this when cleanup has already
        been performed or is not desired.

        .EXAMPLE
        Exit-WithCode -exitCode $Script:ExitCodes.PARAMETER_ERROR -message "Invalid cluster name format"

        Exits with code 2 and logs an error message about invalid parameters.

        .EXAMPLE
        Exit-WithCode -exitCode $Script:ExitCodes.SUCCESS -message "Transition completed successfully"

        Exits with code 0 and logs a success message.

        .EXAMPLE
        Exit-WithCode -exitCode $Script:ExitCodes.CONNECTION_ERROR -message "Failed to connect to SDDC Manager" -noCleanup

        Exits with code 3, logs error, and skips cleanup operations.

        .OUTPUTS
        None. This function terminates the script with the specified exit code.

        .NOTES
        This function should be used for all script exits except in the main menu's exit option,
        which may have its own cleanup logic. Using this consistently throughout the script
        ensures predictable exit behavior for automation and debugging.
    #>
    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNull()] [Int]$ExitCode,
        [Parameter(Mandatory = $false)] [AllowEmptyString()] [String]$Message,
        [Parameter(Mandatory = $false)] [Switch]$NoCleanup
    )

    Write-LogMessage -type DEBUG -message "Entered Exit-WithCode function..."

    # Log final message if provided.
    if ($Message) {
        if ($ExitCode -eq 0) {
            Write-LogMessage -type INFO -message $Message
        } else {
            Write-LogMessage -type ERROR -message $Message
        }
    }

    # Optional cleanup logic for error exits.
    if (-not $NoCleanup -and $ExitCode -ne 0) {
        Write-LogMessage -type DEBUG -message "Exit code $ExitCode indicates failure."
    }

    # Log the exit code for debugging.
    Write-LogMessage -type DEBUG -message "Script exiting with code $ExitCode"

    # Exit with the specified code.
    exit $ExitCode
}
Function Write-LogMessage {

    <#
        .SYNOPSIS
        Writes a severity-based color-coded message to the console and/or log file.

        .DESCRIPTION
        The Write-LogMessage function provides centralized logging functionality with support for
        different message types (INFO, ERROR, WARNING, EXCEPTION, ADVISORY, DEBUG). Messages are displayed
        on the console with color coding based on severity and written to a log file with timestamps.
        This function supports flexible output control allowing messages to be suppressed from either
        the console or log file as needed.

        Screen output is filtered based on the configured log level threshold (set via the -LogLevel
        script parameter). Only messages at or above the configured level are displayed on screen.
        All messages are always written to the log file regardless of their severity level.

        Log level hierarchy (lowest to highest):
        DEBUG < INFO < ADVISORY < WARNING < EXCEPTION < ERROR

        .PARAMETER Message
        The message content to be logged and/or displayed. Can be an empty string if needed.

        .PARAMETER Type
        The severity level of the message. Valid values are:
        - DEBUG (Gray): Debug information for troubleshooting and development
        - INFO (Green): General information messages
        - ADVISORY (Yellow): Advisory information for user guidance
        - WARNING (Yellow): Warning conditions that may need attention
        - EXCEPTION (Cyan): Exception details and stack traces
        - ERROR (Red): Error conditions that require attention
        Default value is "INFO".

        .PARAMETER SuppressOutputToScreen
        When specified, prevents the message from being displayed on the console regardless of log level.

        .PARAMETER SuppressOutputToFile
        When specified, prevents the message from being written to the log file.

        .PARAMETER PrependNewLine
        When specified, adds a blank line before displaying the message on the console.
        This parameter has no effect when SuppressOutputToScreen is used or when the message
        is filtered by log level threshold.

        .PARAMETER AppendNewLine
        When specified, adds a blank line after displaying the message on the console.
        This parameter has no effect when SuppressOutputToScreen is used or when the message
        is filtered by log level threshold.

        .EXAMPLE
        Write-LogMessage -type INFO -message "Process started successfully"
        Displays an informational message in green on the console and writes the message to the log file.

        .EXAMPLE
        Write-LogMessage -type ERROR -message "Failed to connect to server" -prependNewLine
        Displays an error message in red with a blank line before it, and logs it to the file.

        .EXAMPLE
        Write-LogMessage -type WARNING -message "Configuration file not found, using defaults" -suppressOutputToScreen
        Writes a warning message to the log file only, without displaying it on the console.

        .EXAMPLE
        Write-LogMessage -type ADVISORY -message "Consider updating your configuration" -suppressOutputToFile
        Displays an advisory message on the console only, without writing it to the log file.

        .EXAMPLE
        Write-LogMessage -type DEBUG -message "Variable value: $myVar = $($myVar)"
        Displays a debug message in gray on the console (only if log level is DEBUG) and writes it to the log file.

        .NOTES
        The function relies on the $Script:LogFile, $Script:logOnly, and $Script:configuredLogLevel variables being set.
        The log file path should be established using the New-LogFile function before calling this function.
        The $Script:configuredLogLevel should be set during script initialization.

        .OUTPUTS
        None
        This function does not return a value. It writes messages to console and/or log file.
    #>

    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$AppendNewLine,
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [String]$Message,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$PrependNewLine,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$SuppressOutputToFile,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$SuppressOutputToScreen,
        [Parameter(Mandatory = $false)] [ValidateSet("INFO", "ERROR", "WARNING", "EXCEPTION", "ADVISORY", "DEBUG")] [String]$Type = "INFO"
    )

    # Define color mapping for different message types.
    $msgTypeToColor = @{
        "INFO" = "Green";
        "ERROR" = "Red" ;
        "WARNING" = "Yellow" ;
        "ADVISORY" = "Yellow" ;
        "EXCEPTION" = "Cyan";
        "DEBUG" = "Gray"
    }

    # Get the appropriate color for the message type.
    $messageColor = $msgTypeToColor.$Type

    # Create timestamp for log file entries (MM-dd-yyyy_HH:mm:ss format)
    $timeStamp = Get-Date -Format "MM-dd-yyyy_HH:mm:ss"

    # Determine if message should be displayed based on log level threshold.
    $shouldDisplay = Test-LogLevel -MessageType $Type -ConfiguredLevel $Script:configuredLogLevel

    # Add blank line before message if requested and not in log-only mode and meets log level threshold.
    if ($PrependNewLine -and (-not ($Script:logOnly -eq "enabled")) -and $shouldDisplay) {
        Write-Host ""
    }

    # Display message to console with color coding (unless suppressed, in log-only mode, or below log level threshold).
    if (-not $SuppressOutputToScreen -and $Script:logOnly -ne "enabled" -and $shouldDisplay) {
        Write-Host -ForegroundColor $messageColor "[$Type] $Message"
    }

    # Add blank line after message if requested and not in log-only mode and meets log level threshold.
    if ($AppendNewLine -and (-not ($Script:logOnly -eq "enabled")) -and $shouldDisplay) {
        Write-Host ""
    }

    # Write message to log file (unless suppressed).
    if (-not $SuppressOutputToFile) {
        $logContent = '[' + $timeStamp + '] ' + '(' + $Type + ')' + ' ' + $Message
        try {
            Add-Content -ErrorVariable ErrorMessage -Path $Script:LogFile $logContent
        }
        catch {
            # Handle log file write failures gracefully.
            Write-Host "Failed to add content to log file $Script:LogFile."
            Write-Host $errorMessage
        }
    }
}
Function Show-Version {

    <#
        .SYNOPSIS
        Displays or logs the version of the VCF.Powershell.Toolbox module.

        .DESCRIPTION
        The Show-Version function displays or logs the current version of the
        VCF.Powershell.Toolbox module. When called without the -silence parameter,
        it displays the version to the console. With -silence, it only logs the
        version to the log file for audit purposes.

        The version is retrieved from the module manifest (VCF.Powershell.Toolbox.psd1).

        .PARAMETER Silence
        When specified, suppresses console output and only logs the version to the log file.
        This is useful for automated scenarios where console output should be minimized
        while maintaining audit trail in logs.

        .EXAMPLE
        Show-Version

        Displays the module version to the console and logs it to the file.
        Output: "VCF.Powershell.Toolbox Module Version: 1.0.0.2"

        .EXAMPLE
        Show-Version -silence

        Logs the module version to the file only without console output.

        .NOTES
        This function is typically called during environment setup to record
        the module version in log files for troubleshooting purposes.
    #>

    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$Silence
    )

    Write-LogMessage -type DEBUG -message "Entered Show-Version function..."

    # Get the module version from the loaded module manifest
    $moduleVersion = "Unknown"
    try {
        $manifestPath = Join-Path -Path $PSScriptRoot -ChildPath "VCF.Powershell.Toolbox.psd1"

        if (Test-Path $manifestPath) {
            $manifest = Import-PowerShellDataFile -Path $manifestPath -ErrorAction SilentlyContinue
            if ($manifest -and $manifest.ModuleVersion) {
                $moduleVersion = $manifest.ModuleVersion
            }
        } else {
            # Fallback: Try to get version from loaded module
            $loadedModule = Get-Module -Name "VCF.Powershell.Toolbox" -ErrorAction SilentlyContinue
            if ($loadedModule) {
                $moduleVersion = $loadedModule.Version
            }
        }
    } catch {
        Write-LogMessage -type DEBUG -message "Unable to retrieve module version: $_"
    }

    if (-not $Silence) {
        Write-LogMessage -type INFO -message "VCF.Powershell.Toolbox Module Version: $moduleVersion"
    } else {
        Write-LogMessage -type DEBUG -message "VCF.Powershell.Toolbox Module Version: $moduleVersion"
    }
}
Function Get-EnvironmentSetup {

    <#
        .SYNOPSIS
        The function Get-EnvironmentSetup logs user environment details.

        .DESCRIPTION
        The function facilitates troubleshooting by populating each day's log files with useful runtime details.

        .EXAMPLE
        Get-EnvironmentSetup

        .OUTPUTS
        None
        This function does not return a value. It logs environment setup information.
    #>

    Write-LogMessage -type DEBUG -message "Entered Get-EnvironmentSetup function..."

    # Get PowerShell version information.
    $powerShellRelease = $($PSVersionTable.PSVersion).ToString()

    # Check for installed PowerCLI modules (VCF and VMware versions).
    $vcfPowerCliRelease = (Get-Module -ListAvailable -Name VCF.PowerCLI -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1).Version
    $vmwarePowerCliRelease = (Get-Module -ListAvailable -Name VMware.PowerCLI -ErrorAction SilentlyContinue | Sort-Object Version -Descending | Select-Object -First 1).Version

    # Start with basic OS information from PowerShell automatic variables.
    $operatingSystem = $($PSVersionTable.OS)

    # Enhanced macOS information - sw_vers provides more user-friendly OS details than Darwin kernel info.
    if ($IsMacOS) {
        try {
            $macOsName = (sw_vers --productName)
            $macOsRelease = (sw_vers --productVersion)
            $macOsVersion = "$macOsName $macOsRelease"
        } catch [Exception] {
            # If sw_vers fails, we'll fall back to the basic OS info from $PSVersionTable.
        }
    }
    if ($macOsVersion) {
        $operatingSystem = $macOsVersion
    }

    # Enhanced Windows information - Get-ComputerInfo provides more detailed OS information.
    if ($IsWindows) {
        try {
            $windowsProductInformation = (Get-ComputerInfo -ProgressAction SilentlyContinue) | Select-Object OSName,OSVersion
            $windowsVersion = "$($windowsProductInformation.OSName) $($windowsProductInformation.OSVersion)"
        } catch [Exception] {
            # If Get-ComputerInfo fails, we'll fall back to the basic OS info from $PSVersionTable.
        }
    }
    if ($windowsVersion) {
        $operatingSystem = $windowsVersion
    }

    Show-Version -silence

    Write-LogMessage -type DEBUG -message "Client PowerShell version is $powerShellRelease"

    if ($vcfPowerCliRelease) {
        Write-LogMessage -type DEBUG -message "Client VCF.PowerCLI version is $vcfPowerCliRelease."
    }
    if ($vmwarePowerCliRelease) {
        Write-LogMessage -type DEBUG -message "Client VMware.PowerCLI version is $vmwarePowerCliRelease."
    }
    if (-not $vcfPowerCliRelease -and -not $vmwarePowerCliRelease) {
        Write-LogMessage -type ERROR -message "Client PowerCLI not installed. Please install VCF.PowerCLI or VMware.PowerCLI module."
        exit 1
    }

    Write-LogMessage -type DEBUG -message "Client Operating System is $operatingSystem"
}
Function New-LogFile {

    <#
        .SYNOPSIS
        Creates a log file with automatic directory structure and environment logging.

        .DESCRIPTION
        The New-LogFile function establishes the logging infrastructure for the VCF PowerShell
        Toolbox by creating a timestamped log file in a specified directory. The function creates
        one log file using the format mm-dd-yyyy, ensuring logs are organized chronologically.
        If the log directory doesn't exist, it will be created automatically. When a new log file
        is created, the function automatically calls Get-EnvironmentSetup to record system
        information for troubleshooting purposes.

        The function sets the following script-scoped variables:
        - $Script:logFolder: Path to the log directory
        - $Script:logFile: Full path to the current log file

        .PARAMETER Prefix
        Specifies the prefix for the log file name. The final log file will be named
        "{Prefix}-{mm-dd-yyyy}.log". Default value is "VCF.PS.Toolbox".

        .PARAMETER Directory
        Specifies the directory name where log files will be stored, relative to the script root.
        The directory will be created if it doesn't exist. Default value is "logs".

        .EXAMPLE
        New-LogFile
        Creates a log file with default settings: "logs/VCF.PS.Toolbox-01-15-2024.log"

        .EXAMPLE
        New-LogFile -directory "audit" -prefix "SecurityAudit"
        Creates a log file: "audit/SecurityAudit-01-15-2024.log"

        .NOTES
        This function should be called before any Write-LogMessage calls to ensure the log
        infrastructure is properly initialized. The function will exit the script if it
        cannot create the required log directory.
    #>

    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String]$Directory = "logs",
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String]$Prefix = "VCF.Powershell.Toolbox"
    )

    # Generate timestamp for daily log file naming (yyyy-MM-dd format)
    $fileTimeStamp = Get-Date -Format "yyyy-MM-dd"

    # Set script-scoped variables for log directory and file paths.
    $Script:logFolder = Join-Path -Path $PSScriptRoot -ChildPath $Directory
    $Script:logFile = Join-Path -Path $Script:logFolder -ChildPath "$Prefix-$fileTimeStamp.log"

    # Create log directory if it doesn't exist.
    if (-not (Test-Path -Path $Script:logFolder -PathType Container) ) {
        Write-Information "LogFolder not found, creating $Script:logFolder" -InformationAction Continue
        New-Item -ItemType Directory -Path $Script:logFolder | Out-Null
        if (-not $?) {
            Write-Information "Failed to create directory $Script:logFile. Exiting." -InformationAction Continue
            exit 1
        }
    }

    # Create the log file if it doesn't exist for today.
    # When creating a new log file, automatically capture environment details for troubleshooting.
    if (-not (Test-Path $Script:logFile)) {
        New-Item -type File -Path $Script:logFile | Out-Null
        Get-EnvironmentSetup
    }
}
Function Start-ProcessTimer {

    <#
        .SYNOPSIS
        Initializes and starts a high-precision stopwatch for operation timing.

        .DESCRIPTION
        The Start-ProcessTimer function creates and starts a System.Diagnostics.Stopwatch
        object for measuring elapsed time of operations. This provides a consistent way
        to begin timing across the VCF PowerShell Toolbox functions. The returned stopwatch
        object should be used with the Stop-ProcessTimer function for consistent logging.

        .OUTPUTS
        System.Diagnostics.Stopwatch
        A started stopwatch object that can be used to measure elapsed time.

        .EXAMPLE
        $timer = Start-ProcessTimer
        # Perform some operation
        Stop-ProcessTimer -timer $timer -operation "Data processing" -interval "Seconds"

        .NOTES
        This function is designed to be paired with Stop-ProcessTimer for complete
        timing functionality with automatic logging.
    #>

    return [System.Diagnostics.Stopwatch]::StartNew()
}
Function Stop-ProcessTimer {
    <#
        .SYNOPSIS
        Stops a stopwatch timer and logs the elapsed time for the specified operation.

        .DESCRIPTION
        The Stop-ProcessTimer function stops a System.Diagnostics.Stopwatch object and
        automatically logs the elapsed time to the log file. The elapsed time is calculated
        and rounded to 2 decimal places based on the specified interval (milliseconds,
        seconds, or minutes). This provides consistent timing and logging across all
        VCF PowerShell Toolbox operations.

        .PARAMETER Timer
        The System.Diagnostics.Stopwatch object to stop. This should be a stopwatch
        that was started using the Start-ProcessTimer function.

        .PARAMETER Operation
        A descriptive name for the operation that was being timed. This will be included
        in the log message for identification purposes.

        .PARAMETER Interval
        The time unit for reporting the elapsed time. Valid values are:
        - "Milliseconds": Reports time in milliseconds (ms)
        - "Seconds": Reports time in seconds (s)
        - "Minutes": Reports time in minutes (min)

        .EXAMPLE
        $timer = Start-ProcessTimer
        # Perform vCenter connection
        Stop-ProcessTimer -timer $timer -operation "vCenter connection" -interval "Seconds"
        # Logs: "vCenter connection took 2.35 Seconds to complete."

        .EXAMPLE
        $timer = Start-ProcessTimer
        # Perform quick operation
        Stop-ProcessTimer -timer $timer -operation "API call" -interval "Milliseconds"
        # Logs: "API call took 250.75 Milliseconds to complete."

        .NOTES
        The elapsed time is automatically logged to the file (not displayed on console).
        All timing measurements are rounded to 2 decimal places for consistency.
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateSet("Milliseconds","Seconds","Minutes")] [String]$Interval,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Operation,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [System.Diagnostics.Stopwatch]$Timer
    )

    # Stop the stopwatch to capture final elapsed time.
    $Timer.Stop()

    # Calculate elapsed time based on requested interval and round to 2 decimal places.
    switch ($Interval) {
        "Milliseconds" {
            $elapsedInterval = [math]::Round(($Timer.elapsed.totalMilliseconds), 2)
        }
        "Seconds" {
            $elapsedInterval = [math]::Round(($Timer.elapsed.totalSeconds), 2)
        }
        "Minutes" {
            $elapsedInterval = [math]::Round(($Timer.elapsed.totalMinutes), 2)
        }
    }

    # Log the timing result (suppressed from console output).
    Write-LogMessage -type INFO -suppressOutputToScreen -message "$Operation took $elapsedInterval $Interval to complete."
}
Function ConvertFrom-JsonSafely {

    <#
        .SYNOPSIS
        Safely loads and validates JSON content from a file with comprehensive error handling.

        .DESCRIPTION
        The ConvertFrom-JsonSafely function provides a robust way to load JSON files with
        built-in validation and error handling. The function reads the file content, removes
        empty lines that could cause JSON parsing issues, and converts the content to a
        PowerShell object. If JSON validation fails, the function logs detailed error
        information including the file path and specific parsing error, then exits the
        script to prevent further execution with invalid data.

        This function standardizes JSON loading across the VCF PowerShell Toolbox and
        ensures consistent error reporting for troubleshooting.

        .PARAMETER JsonFilePath
        The full path to the JSON file to load and parse. The file must exist and
        contain valid JSON content.

        .EXAMPLE
        $Config = ConvertFrom-JsonSafely -JsonFilePath "C:\configs\settings.json"
        Loads application settings from a JSON file with error handling.

        .NOTES
        This function will terminate script execution (exit) if JSON parsing fails.
        Empty lines are automatically filtered out before JSON parsing to handle
        files that may have formatting issues.
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonFilePath
    )

    Write-LogMessage -type DEBUG -message "Entered ConvertFrom-JsonSafely function..."

    try {
        # Read file content, filter out empty lines, and convert from JSON,
        # Empty line filtering prevents JSON parsing issues with poorly formatted files,
        return (Get-Content $JsonFilePath) | Select-String -Pattern "^\s*$" -NotMatch | ConvertFrom-Json

    }
    catch {
        # Handle JSON parsing errors with detailed, user-friendly logging.
        $errorMessage = $_.Exception.Message

        Write-LogMessage -type ERROR -message "JSON validation failed for file: $JsonFilePath"
        Write-Host ""

        # Extract the specific JSON error and location
        if ($errorMessage -match "Bad JSON escape sequence: \\([A-Za-z])\..*'([^']+)'.*line (\d+).*position (\d+)") {
            $badChar = $matches[1]
            $jsonPath = $matches[2]
            $lineNum = $matches[3]
            $position = $matches[4]

            Write-LogMessage -type ERROR -message "Invalid escape sequence: '\$badChar' in JSON property '$jsonPath'"
            Write-LogMessage -type ERROR -message "Location: Line $lineNum, Position $position"
            Write-Host ""
            Write-LogMessage -type ERROR -message "Common causes:"
            Write-LogMessage -type ERROR -message "  1. Windows file paths must use forward slashes (/) or escaped backslashes (\\\\)"
            Write-LogMessage -type ERROR -message "     Example: `"C:/Users/Admin/file.yml`" or `"C:\\\\Users\\\\Admin\\\\file.yml`""
            Write-LogMessage -type ERROR -message "  2. Backslash (\) is a special character in JSON and must be escaped"
            Write-Host ""
            Write-LogMessage -type ERROR -message "Please correct the JSON syntax in '$JsonFilePath' at line $lineNum and try again."
        }
        elseif ($errorMessage -match "Conversion from JSON failed with error: (.+?)\. Path '([^']+)'.*line (\d+).*position (\d+)") {
            $jsonError = $matches[1]
            $jsonPath = $matches[2]
            $lineNum = $matches[3]
            $position = $matches[4]

            Write-LogMessage -type ERROR -message "JSON parsing error: $jsonError"
            Write-LogMessage -type ERROR -message "Property: '$jsonPath'"
            Write-LogMessage -type ERROR -message "Location: Line $lineNum, Position $position"
            Write-Host ""
            Write-LogMessage -type ERROR -message "Please correct the JSON syntax in '$JsonFilePath' and try again."
        }
        else {
            # Fallback for unexpected error formats
            Write-LogMessage -type ERROR -message "JSON parsing error: $errorMessage"
        }

        # Exit script execution to prevent continuing with invalid data.
        exit 1
    }
}
Function New-ChoiceMenu {

    <#
        .SYNOPSIS
        Presents an interactive yes/no choice menu to the user with a configurable default.

        .DESCRIPTION
        The New-ChoiceMenu function creates a standardized interactive prompt that presents
        the user with a yes/no decision. The function uses PowerShell's built-in choice
        prompt functionality to provide a consistent user experience across the VCF PowerShell
        Toolbox. The user can select options using Y/N keys or simply press Enter to accept
        the default choice.

        The function returns an integer value (0 for Yes, 1 for No) that can be used in
        conditional logic to determine the user's decision.

        .PARAMETER Question
        The question or prompt text to display to the user. This should be a clear,
        concise question that can be answered with yes or no.

        .PARAMETER DefaultAnswer
        The default answer that will be selected if the user presses Enter without
        making a selection. Valid values are "Yes" or "No" (case-sensitive).

        .OUTPUTS
        System.Int32
        Returns 0 if the user selects Yes, or 1 if the user selects No.

        .EXAMPLE
        $decision = New-ChoiceMenu -question "Would you like to create the log folder?" -defaultAnswer "Yes"
        if ($decision -eq 0) {
            Write-Host "User chose Yes"
        } else {
            Write-Host "User chose No"
        }

        .EXAMPLE
        $continue = New-ChoiceMenu -question "Do you want to proceed with the operation?" -defaultAnswer "No"
        Creates a prompt with "No" as the default, requiring explicit user confirmation.

        .NOTES
        This function requires an interactive PowerShell session and will not work in
        non-interactive or headless environments. The default answer parameter is
        case-sensitive and must be exactly "Yes" or "No".
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$DefaultAnswer,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Question
    )

    # Create a collection to hold the choice options
    $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]

    # Add Yes and No options with keyboard shortcuts (&Y and &N)
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes', "Yes"))
    $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No', "No"))

    # Set the default choice based on the DefaultAnswer parameter
    # Index 0 = Yes, Index 1 = No
    # Note: $title is intentionally $null as we use $Question for the prompt text
    $title = $null
    if ($DefaultAnswer -eq "Yes") {
        $decision = $host.UI.PromptForChoice($title, $Question, $choices, 0)
    }
    else {
        $decision = $host.UI.PromptForChoice($title, $Question, $choices, 1)
    }

    return $decision
}
Function Show-AnyKey {

    <#
        .SYNOPSIS
        Pauses script execution and waits for user to press any key before continuing.

        .DESCRIPTION
        The Show-AnyKey function provides a standardized way to pause script execution
        and wait for user acknowledgment before proceeding. This is commonly used after
        displaying information, completing operations, or before returning to menus.
        The function only operates in interactive mode and is automatically bypassed
        when the script is running in headless mode.

        The function displays a yellow-colored prompt message and captures any keystroke
        without echoing it to the console, providing a clean user experience.

        .EXAMPLE
        Show-AnyKey
        Displays "Press any key to continue..." and waits for user input.

        .NOTES
        This function checks the $headless variable and only prompts for input when
        $headless equals "disabled". In headless mode, the function returns immediately
        without any user interaction, allowing scripts to run unattended.

        The function uses RawUI.ReadKey with 'NoEcho,IncludeKeyDown' options to capture
        keystrokes without displaying them on the console.
    #>

    # Only prompt for user input when not in headless mode.
    if (-not $Script:headless -or $Script:headless -eq "disabled") {
        Write-Host "`nPress any key to continue...`n" -ForegroundColor Yellow;
        # Capture keystroke without echoing to console
        $host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown') | Out-Null
    }
}
Function Test-EmptyValue {

     <#
        .SYNOPSIS
        Validates that a string value is not null or empty, exiting the script if validation fails.

        .DESCRIPTION
        The Test-EmptyValue function provides a standardized way to validate required string
        parameters or variables throughout the VCF PowerShell Toolbox. When a value is found
        to be null or empty, the function logs an error message identifying the specific field
        that failed validation and immediately terminates script execution with exit code 1.

        This function is designed to be used for critical validations where continuing execution
        with missing data would lead to unpredictable results or failures. It provides consistent
        error reporting and ensures that scripts fail fast when required data is missing.

        .PARAMETER FieldName
        A descriptive name for the field or variable being validated. This name will be included
        in the error message to help identify which specific field failed validation.

        .PARAMETER Value
        The string value to validate. The parameter allows empty strings to be passed (using
        [AllowEmptyString()]) so that the function can properly detect and report empty values.

        .EXAMPLE
        Test-EmptyValue -FieldName "Username" -Value $username
        Validates that the Username variable is not null or empty, logging "Username is empty." if validation fails.

    #>

   Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$FieldName,
        [Parameter(Mandatory = $true)] [AllowEmptyString()] [String]$Value
    )

    if ([String]::IsNullOrEmpty($Value)) {
        Write-LogMessage -type ERROR -message "$FieldName is empty."
        exit 1
    }
}
Function Get-InteractiveInput {

    <#
        .SYNOPSIS
        Prompts the user for input and returns the value.

        .DESCRIPTION
        The Get-InteractiveInput function provides a standardized way to prompt the user for input and return the value.
        This function is designed to be used for interactive input throughout the VCF PowerShell Toolbox.

        .PARAMETER PromptMessage
        The message to display to the user.

        .PARAMETER AsSecureString
        When specified, the function will prompt the user for input as a secure string.

        .OUTPUTS
        System.String
        Returns the user's input as a string.

        .EXAMPLE
        $username = Get-InteractiveInput -promptMessage "Enter your username:"
        Prompts the user for a username and returns the value.
    #>

    Param(
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$AsSecureString,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$PromptMessage
    )

    do {
        if ($AsSecureString) {
            $value = Read-Host $PromptMessage -asSecureString
        } else {
            $value = Read-Host $PromptMessage
        }
    } while ($value -eq "")

    return $value
}
Function Test-ArrayMissingProperties {

    <#
        .SYNOPSIS
        Checks for missing properties in objects within an array and returns detailed validation results.

        .DESCRIPTION
        The Test-ArrayMissingProperties function validates that all objects in an array contain
        the specified required properties. This function is useful for validating configuration
        data, API responses, or any collection of objects that should have a consistent schema.

        The function returns a comprehensive validation result that includes:
        - Overall validation status (pass/fail).
        - List of missing properties per object.
        - Summary of validation issues.
        - Detailed error information for troubleshooting.

        .PARAMETER Array
        The array of objects to validate. Each object in the array will be checked for
        the presence of the required properties.

        .PARAMETER ArrayName
        A descriptive name for the array being validated, used in error messages and
        logging to help identify the source of validation failures.

        .PARAMETER RequiredProperties
        An array of property names that must be present in each object. Property names
        are case-sensitive and must match exactly.

        .PARAMETER StopOnFirstError
        When specified, the function will stop validation and return immediately upon
        finding the first missing property, rather than validating the entire array.

        .OUTPUTS
        System.Management.Automation.PSCustomObject
        Returns an object with the following properties:
        - isValid: Boolean indicating if all validations passed
        - MissingProperties: Array of objects detailing missing properties per item
        - ErrorCount: Total number of validation errors found
        - Summary: Human-readable summary of validation results.

        .EXAMPLE
        $users = @(
            @{ name = "John"; email = "john@example.com" },
            @{ name = "Jane" },
            @{ email = "bob@example.com" }
        )
        $result = Test-ArrayMissingProperties -array $users -requiredProperties @("name", "email") -arrayName "Users"

        if (-not $result.isValid) {
            return
        }

        .EXAMPLE
        $config = ConvertFrom-JsonSafely -jsonFilePath "config.json"
        $validationResult = Test-ArrayMissingProperties -array $config.servers -requiredProperties @("hostname", "username", "password") -arrayName "ServerConfiguration" -stopOnFirstError

        if (-not $validationResult.isValid) {
            return
        }

        .NOTES
        This function is designed to work with arrays of PSCustomObject, hashtables, or any
        objects that support property access via PSObject.Properties. The function logs
        detailed validation results and integrates with the VCF PowerShell Toolbox logging
        infrastructure for consistent error reporting.
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [Array]$Array,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$ArrayName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String[]]$RequiredProperties,
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [Switch]$StopOnFirstError
    )

    # Initialize validation result object
    $validationResult = [PSCustomObject]@{
        isValid = $true
        MissingProperties = @()
        ErrorCount = 0
        Summary = ""
    }

    Write-LogMessage -type INFO -message "Starting validation of $ArrayName with $($Array.count) items for properties: $($RequiredProperties -join ', ')" -suppressOutputToScreen

    # Validate each object in the array
    for ($i = 0; $i -lt $Array.count; $i++) {
        $currentObject = $Array[$i]
        $missingProps = @()

        # Check each required property
        foreach ($property in $RequiredProperties) {
            # Handle different object types (PSCustomObject, Hashtable, etc.)
            $hasProperty = $false

            if ($currentObject -is [System.Collections.Hashtable]) {
                $hasProperty = $currentObject.ContainsKey($property)
            } elseif ($currentObject.PSObject.Properties[$property]) {
                $hasProperty = $true
            }

            if (-not $hasProperty) {
                $missingProps += $property
                $validationResult.errorCount++
            }
        }

        # Record missing properties for this object
        if ($missingProps.count -gt 0) {
            $validationResult.isValid = $false
            $missingPropertyInfo = [PSCustomObject]@{
                Index = $i
                Missing = $missingProps
            }
            $validationResult.missingProperties += $missingPropertyInfo

            Write-LogMessage -type ERROR -appendNewLine -message "$ArrayName item at index $i is missing required properties: $($missingProps -join ', ')"

            # Stop on first error if requested
            if ($StopOnFirstError) {
                break
            }
        }
    }

    # Generate summary message and log message
    if ($validationResult.isValid) {
        $validationResult.summary = "$ArrayName validation passed. All $($Array.count) item(s) contain required properties."
        Write-LogMessage -type INFO -suppressOutputToScreen -message $validationResult.summary
    } else {
        $affectedItems = $validationResult.missingProperties.count
        $validationResult.summary = "$ArrayName validation failed. $affectedItems of $($Array.count) items are missing required properties ($($validationResult.errorCount) total missing properties)."
        Write-LogMessage -type ERROR -suppressOutputToScreen -message $validationResult.summary
    }

    return $validationResult
}
Function Get-JsonDataWithValidation {

    <#
        .SYNOPSIS
        Loads and validates JSON file existence and parseability with consistent error handling.

        .DESCRIPTION
        Common helper function for JSON validation functions that handles file existence checking
        and JSON parsing with consistent error handling and logging. This function eliminates
        code duplication across Test-JsonMissingProperties and Test-JsonNullValues by centralizing
        the common file validation and parsing logic.

        The function performs two critical validations:
        1. Verifies the JSON file exists at the specified path
        2. Attempts to parse the JSON file using ConvertFrom-JsonSafely

        If either validation fails, the function updates the provided ValidationResult object
        with appropriate error information and returns $null. On success, it returns the parsed
        JSON data and stores it in the ValidationResult.JsonData property.

        .PARAMETER JsonFilePath
        Path to the JSON file to load and validate.

        .PARAMETER JsonObjectName
        Name of the JSON object for error messages and logging (e.g., "InputConfiguration", "SupervisorConfiguration").
        This name is used to provide context in error messages.

        .PARAMETER ValidationResult
        Reference to the validation result object to update on error. The function will set
        IsValid, ErrorCount, and Summary properties on validation failure.

        .OUTPUTS
        PSCustomObject - Parsed JSON data on success, or $null if validation failed.

        .EXAMPLE
        $jsonData = Get-JsonDataWithValidation -JsonFilePath $JsonFilePath -JsonObjectName $JsonObjectName -ValidationResult ([ref]$validationResult)
        if ($null -eq $jsonData) {
            return $validationResult
        }

        Loads JSON data and returns early if validation fails.

        .NOTES
        This function is a helper for Test-JsonMissingProperties and Test-JsonNullValues.

        Error Handling:
        • Updates ValidationResult object with error details
        • Logs errors using Write-LogMessage
        • Returns $null on any validation failure
        • Preserves parsed JSON data in ValidationResult.JsonData on success
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonFilePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonObjectName,
        [Parameter(Mandatory = $true)] [ref]$ValidationResult
    )

    Write-LogMessage -type DEBUG -message "Validating and loading JSON file: $JsonFilePath"

    # Validate that the JSON file exists.
    if (-not (Test-Path -Path $JsonFilePath -PathType Leaf)) {
        $ValidationResult.Value.IsValid = $false
        $ValidationResult.Value.ErrorCount = 1
        $ValidationResult.Value.Summary = "$JsonObjectName validation failed: File $JsonFilePath does not exist."
        Write-LogMessage -type ERROR -message $ValidationResult.Value.Summary
        return $null
    }

    # Load and parse the JSON file.
    try {
        $jsonData = ConvertFrom-JsonSafely -JsonFilePath $JsonFilePath
        $ValidationResult.Value.JsonData = $jsonData
        return $jsonData
    }
    catch {
        $ValidationResult.Value.IsValid = $false
        $ValidationResult.Value.ErrorCount = 1
        $ValidationResult.Value.Summary = "$JsonObjectName validation failed: Unable to parse JSON file $JsonFilePath. Error: $_"
        Write-LogMessage -type ERROR -message $ValidationResult.Value.Summary
        return $null
    }
}
Function Test-JsonFile {

    <#
        .SYNOPSIS
        Validates JSON file existence and content with proper resource management and comprehensive error handling.

        .DESCRIPTION
        The Test-JsonFile function provides robust validation of JSON files by checking both file existence
        and JSON content validity. It uses the .NET System.Text.Json.JsonDocument class for efficient
        parsing and implements proper resource disposal to prevent memory leaks.

        Key features:
        - File existence validation with detailed error reporting
        - Strict JSON parsing using System.Text.Json.JsonDocument
        - Proper resource disposal using try/finally blocks
        - Comprehensive error handling with specific exception types
        - Integration with the script's logging system
        - Performance optimized for large JSON files

        The function will return $true if the file exists and contains valid JSON, $false otherwise.
        All errors are logged using the Write-LogMessage system for consistent error reporting.

        .PARAMETER JsonFilePath
        The absolute path to the JSON file to be validated. This parameter is mandatory and must
        point to an existing file. The path can be either a local file path or a UNC path.

        .EXAMPLE
        Test-JsonFile -JsonFilePath "C:\config\settings.json"
        Returns $true if the file exists and contains valid JSON, $false otherwise.

        .EXAMPLE
        if (Test-JsonFile -JsonFilePath $configPath) {
            Write-Host "Configuration file is valid"
            $config = Get-Content $configPath | ConvertFrom-Json
        }

        .OUTPUTS
        System.Boolean
        Returns $true if the file exists and contains valid JSON content, $false otherwise.

        .NOTES
        - Uses System.Text.Json.JsonDocument for efficient JSON validation
        - Implements proper resource disposal to prevent memory leaks
        - All validation errors are logged using Write-LogMessage
        - Function is optimized for performance with large JSON files
        - Compatible with both Windows PowerShell 5.1 and PowerShell 7+
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateScript({ if ([string]::IsNullOrWhiteSpace($_)) { throw "JSON file path cannot be null, empty, or contain only whitespace characters." }; if ($_.Length -gt 260) { throw "JSON file path cannot exceed 260 characters. Current length: $($_.Length)" }; if ($_ -match '[<>"|?*]') { throw "JSON file path contains invalid characters: $($matches[0])" }; return $true })] [ValidateNotNullOrEmpty()] [String]$JsonFilePath
    )

    Write-LogMessage -type DEBUG -message "Entered Test-JsonFile function..."

    # Validate file existence first.
    if (-not (Test-Path -Path $JsonFilePath -PathType Leaf)) {
        Write-LogMessage -type ERROR -message "JSON file not found: '$JsonFilePath'"
        return $false
    }

    # Validate file is actually a file (not a directory)
    $fileInfo = Get-Item -Path $JsonFilePath -ErrorAction SilentlyContinue
    if ($fileInfo -and $fileInfo.PSIsContainer) {
        Write-LogMessage -type ERROR -message "Specified path is a directory, not a file: '$JsonFilePath'"
        return $false
    }

    # Check if file is readable.
    try {
        $null = Get-Content -Path $JsonFilePath -TotalCount 1 -ErrorAction Stop
    } catch [System.UnauthorizedAccessException] {
        Write-LogMessage -type ERROR -message "Access denied reading JSON file: '$JsonFilePath'. Please check file permissions."
        return $false
    } catch [System.IO.IOException] {
        Write-LogMessage -type ERROR -message "I/O error reading JSON file: '$JsonFilePath'. File may be locked or corrupted."
        return $false
    } catch {
        Write-LogMessage -type ERROR -message "Unexpected error reading JSON file: '$JsonFilePath': $($_.Exception.Message)"
        return $false
    }

    # Validate JSON content.
    $jsonDocument = $null
    try {
        Write-LogMessage -type INFO -suppressOutputToScreen -message "Validating JSON content in file: '$JsonFilePath'"

        # Read file content
        $content = Get-Content -Path $JsonFilePath -Raw -ErrorAction Stop

        # Check for empty file.
        if ([string]::IsNullOrWhiteSpace($content)) {
            Write-LogMessage -type ERROR -message "JSON file is empty or contains only whitespace: '$JsonFilePath'"
            return $false
        }

        # Load and validate JSON using System.Text.Json for strict parsing.
        Add-Type -AssemblyName System.Text.Json -ErrorAction Stop

        # Parse JSON with strict validation.
        $jsonDocument = [System.Text.Json.JsonDocument]::Parse($content)

        # If we reach here, JSON is valid.
        Write-LogMessage -type INFO -suppressOutputToScreen -message "JSON file validation successful: '$JsonFilePath'"
        return $true

    } catch [System.Text.Json.JsonException] {
        # Handle JSON parsing errors specifically.
        Write-LogMessage -type ERROR -message "Invalid JSON format in file: '$JsonFilePath'"
        Write-LogMessage -type ERROR -message "JSON parsing error: $($_.Exception.Message)"
        return $false
    } catch [System.ArgumentException] {
        # Handle argument exceptions (e.g., invalid UTF-8 encoding)
        Write-LogMessage -type ERROR -message "Invalid content encoding in JSON file: '$JsonFilePath'"
        Write-LogMessage -type ERROR -message "Encoding error: $($_.Exception.Message)"
        return $false
    } catch [System.IO.FileNotFoundException] {
        # Handle case where file was deleted between existence check and read.
        Write-LogMessage -type ERROR -message "JSON file was deleted during validation: '$JsonFilePath'"
        return $false
    } catch [System.OutOfMemoryException] {
        # Handle very large files that exceed memory limits.
        Write-LogMessage -type ERROR -message "JSON file too large to process: '$JsonFilePath'. File may exceed available memory."
        return $false
    } catch {
        # Handle any other unexpected exceptions.
        Write-LogMessage -type ERROR -message "Unexpected error during JSON validation for file: '$JsonFilePath'"
        Write-LogMessage -type ERROR -message "Error details: $($_.Exception.Message)"
        return $false
    } finally {
        # Ensure proper resource disposal.
        if ($jsonDocument) {
            try {
                $jsonDocument.Dispose()
                Write-LogMessage -type INFO -suppressOutputToScreen -message "JSON document resources properly disposed for: '$JsonFilePath'"
            } catch {
                Write-LogMessage -type WARNING -suppressOutputToScreen -message "Warning: Could not dispose JSON document resources for: '$JsonFilePath': $($_.Exception.Message)"
            }
        }
    }
}
Function Get-JsonPropertyValue {

    <#
        .SYNOPSIS
        Extracts a property value from a JSON object using dot-notation path.

        .DESCRIPTION
        The Get-JsonPropertyValue function navigates nested JSON objects, PSCustomObjects, or Hashtables
        using a dot-notation property path (e.g., "parent.child.property") and returns the value as a string.
        This helper function separates the concern of property extraction from validation logic.

        .PARAMETER InputData
        The input data object (JSON, PSCustomObject, Hashtable, or String) to extract the value from.

        .PARAMETER PropertyPath
        Optional. The dot-notation path to the property (e.g., "common.vCenterName"). If not specified
        and InputData is a string, returns the string directly. If not specified and InputData is an
        object, converts the entire object to a string.

        .OUTPUTS
        System.String
        Returns the extracted property value as a string, or $null if extraction fails.

        .EXAMPLE
        $value = Get-JsonPropertyValue -InputData $config -PropertyPath "common.vCenterName"
        Extracts the vCenterName property from the common section of the config object.

        .NOTES
        This is a helper function used by Test-JsonPropertyFormat to separate property extraction
        from validation logic, improving testability and maintainability.
    #>

    Param (
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String]$PropertyPath,
        [Parameter(Mandatory = $true)] [AllowNull()] [AllowEmptyString()] $InputData
    )

    Write-LogMessage -type DEBUG -message "Entered Get-JsonPropertyValue function..."

    try {
        # Handle null input
        if ($null -eq $InputData) {
            Write-LogMessage -type ERROR -suppressOutputToScreen -message "Input data is null"
            return $null
        }

        # If InputData is already a string, return it directly.
        if ($InputData -is [String]) {
            Write-LogMessage -type DEBUG -message "Input is already a string with length: $($InputData.Length)"
            return $InputData
        }

        # If PropertyPath is specified, extract the property value.
        if ($PropertyPath) {
            Write-LogMessage -type DEBUG -message "Extracting property '$PropertyPath' from input object"

            # Split property path by dots to navigate nested properties.
            $pathParts = $PropertyPath.Split('.')
            $currentObject = $InputData

            foreach ($part in $pathParts) {
                if ($null -eq $currentObject) {
                    Write-LogMessage -type ERROR -suppressOutputToScreen -message "Property path '$PropertyPath' contains null value at '$part'"
                    return $null
                }

                # Handle PSCustomObject, Hashtable, and regular object property access.
                if ($currentObject -is [PSCustomObject]) {
                    $currentObject = $currentObject.$part
                } elseif ($currentObject -is [Hashtable]) {
                    $currentObject = $currentObject[$part]
                } else {
                    try {
                        $currentObject = $currentObject.$part
                    } catch {
                        Write-LogMessage -type ERROR -suppressOutputToScreen -message "Cannot access property '$part' in path '$PropertyPath': $($_.Exception.Message)"
                        return $null
                    }
                }
            }

            # Convert the final property value to string.
            $result = if ($null -eq $currentObject) { "" } else { $currentObject.ToString() }
            Write-LogMessage -type DEBUG -message "Extracted value: '$result' (length: $($result.Length))"
            return $result
        }
        # If no PropertyPath specified, convert entire object to string.
        else {
            $result = $InputData.ToString()
            Write-LogMessage -type DEBUG -message "Converted entire object to string (length: $($result.Length))"
            return $result
        }
    }
    catch {
        Write-LogMessage -type ERROR -suppressOutputToScreen -message "Error extracting property value: $($_.Exception.Message)"
        return $null
    }
}
Function Test-JsonMissingProperties {

    <#
        .SYNOPSIS
        Validates JSON file content for missing required properties with support for nested properties.

        .DESCRIPTION
        The Test-JsonMissingProperties function provides comprehensive validation of JSON files
        to ensure all required properties are present. It supports nested property validation
        using dot notation (e.g., "common.vCenter.name") and provides detailed reporting of
        missing properties with their expected structure.

        This function is particularly useful for validating configuration files, API payloads,
        or any JSON data that must conform to a specific schema. It integrates with the VCF
        PowerShell Toolbox logging infrastructure for consistent error reporting.

        .PARAMETER JsonFilePath
        The full path to the JSON file to validate. The file must exist and contain valid JSON content.

        .PARAMETER JsonObjectName
        A descriptive name for the JSON object being validated, used in error messages and
        logging to help identify the source of validation failures.

        .PARAMETER RequiredProperties
        An array of property names (using dot notation for nested properties) that must be present
        in the JSON object. Examples: "name", "config.database.host", "settings.security.enabled"

        .PARAMETER ShowExpectedStructure
        When specified, the function will include the expected JSON structure for missing
        properties in the validation results, helpful for troubleshooting and documentation.

        .PARAMETER StopOnFirstError
        When specified, the function will stop validation and return immediately upon
        finding the first missing property, rather than validating all properties.

        .OUTPUTS
        System.Management.Automation.PSCustomObject
        Returns an object with the following properties:
        - IsValid: Boolean indicating if all validations passed
        - MissingProperties: Array of missing property paths
        - ExpectedStructure: Suggested JSON structure for missing properties (if ShowExpectedStructure is used)
        - ErrorCount: Total number of missing properties
        - Summary: Human-readable summary of validation results
        - JsonData: The loaded JSON object (if validation passes)

        .EXAMPLE
        $validationResult = Test-JsonMissingProperties -JsonFilePath "config.json" -RequiredProperties @("database.host", "database.port", "api.key") -JsonObjectName "Configuration"

        if (-not $validationResult.IsValid) {
            Write-Host "Validation failed: $($validationResult.Summary)"
            return
        }
        $config = $validationResult.JsonData

        .EXAMPLE
        $requiredProps = @(
            "common.vCenterName",
            "common.VcenterUser",
            "common.esxHost"
        )
        $result = Test-JsonMissingProperties -JsonFilePath "input.json" -RequiredProperties $requiredProps -JsonObjectName "InputConfiguration" -ShowExpectedStructure

        .NOTES
        This function uses the existing ConvertFrom-JsonSafely function for safe JSON loading
        and integrates with the VCF PowerShell Toolbox logging infrastructure. Nested properties
        are accessed using dot notation, and the function provides detailed error reporting
        for missing properties at any depth in the JSON structure.
    #>

    Param (
        [Parameter(Mandatory = $false)] [Switch]$ShowExpectedStructure,
        [Parameter(Mandatory = $false)] [Switch]$StopOnFirstError,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonFilePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonObjectName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String[]]$RequiredProperties
    )

    Write-LogMessage -type DEBUG -message "Entered Test-JsonMissingProperties function..."

    # Initialize validation result object.
    $validationResult = [PSCustomObject]@{
        IsValid = $true
        MissingProperties = @()
        ExpectedStructure = @{}
        ErrorCount = 0
        Summary = ""
        JsonData = $null
    }

    Write-LogMessage -type INFO -suppressOutputToScreen -message "Validating $($RequiredProperties.Count) required properties: $($RequiredProperties -join ', ')"

    # Load and validate the JSON file using helper function.
    $jsonData = Get-JsonDataWithValidation -JsonFilePath $JsonFilePath -JsonObjectName $JsonObjectName -ValidationResult ([ref]$validationResult)
    if ($null -eq $jsonData) {
        return $validationResult
    }

    # Helper function to check if a nested property exists using dot notation.
    Function Test-NestedProperty {
        param($Object, $PropertyPath)

        # Split the property path into individual segments using dot as delimiter.
        $properties = $PropertyPath -split '\.'
        # Start traversal from the root object.
        $currentObject = $Object

        # Iterate through each property segment in the path.
        foreach ($Property in $properties) {
            # Handle hashtable objects - use ContainsKey for property existence check.
            if ($currentObject -is [System.Collections.Hashtable]) {
                if (-not $currentObject.ContainsKey($Property)) {
                    return $false
                }
                # Move to the next level in the hierarchy.
                $currentObject = $currentObject[$Property]
            }
            # Handle PowerShell custom objects - check PSObject.Properties collection.
            elseif ($currentObject.PSObject.Properties[$Property]) {
                # Move to the next level in the hierarchy.
                $currentObject = $currentObject.$Property
            }
            # Property doesn't exist in current object - path is invalid.
            else {
                return $false
            }
        }

        # Successfully traversed the entire path.
        return $true
    }

    # Validate each required property.
    foreach ($Property in $RequiredProperties) {
        $propertyExists = Test-NestedProperty -Object $jsonData -PropertyPath $Property

        if (-not $propertyExists) {
            $validationResult.IsValid = $false
            $validationResult.MissingProperties += $Property
            $validationResult.ErrorCount++

            Write-LogMessage -type ERROR -message "$JsonObjectName (in JSON file $JsonFilePath) is missing required property '$Property'."

            # Generate expected structure if requested.
            if ($ShowExpectedStructure) {
                $pathParts = $Property -split '\.'
                $structure = $validationResult.ExpectedStructure
                $current = $structure

                for ($i = 0; $i -lt $pathParts.Count - 1; $i++) {
                    if (-not $current.ContainsKey($pathParts[$i])) {
                        $current[$pathParts[$i]] = @{}
                    }
                    $current = $current[$pathParts[$i]]
                }
                $current[$pathParts[-1]] = "<value>"
            }

            # Stop on first error if requested.
            if ($StopOnFirstError) {
                break
            }
        }
    }

    # Generate summary message.
    if ($validationResult.IsValid) {
        $validationResult.Summary = "$JsonObjectName validation passed. All $($RequiredProperties.Count) required properties are present."
        Write-LogMessage -type INFO -suppressOutputToScreen -message $validationResult.Summary
    }
    else {
        $validationResult.Summary = "$JsonObjectName validation failed. $($validationResult.ErrorCount) of $($RequiredProperties.Count) required properties are missing: $($validationResult.MissingProperties -join ', ')"
        Write-LogMessage -type ERROR -message $validationResult.Summary
    }

    # Store the JSON data in the result for caller use.
    $validationResult.JsonData = $jsonData

    return $validationResult
}
Function Test-JsonNullValues {

    <#
        .SYNOPSIS
        Validates that specified JSON properties are not null.

        .DESCRIPTION
        The Test-JsonNullValues function checks whether specified properties in a JSON file
        contain null values. This is a complementary validation to Test-JsonMissingProperties,
        which only checks if keys exist. This function ensures that existing keys also have
        non-null values.

        This validation is critical because PowerShell's JSON parsing will include properties
        with null values in the object structure, making them technically "present" but unusable.
        Configuration files must have actual values, not nulls, for deployment to succeed.

        .PARAMETER JsonFilePath
        The full path to the JSON file to validate. The file must exist and contain valid JSON content.

        .PARAMETER JsonObjectName
        A descriptive name for the JSON object being validated, used in error messages and
        logging to help identify the source of validation failures.

        .PARAMETER RequiredProperties
        An array of property names (using dot notation for nested properties) that must have
        non-null values. Examples: "name", "config.database.host", "settings.security.enabled"

        .PARAMETER StopOnFirstError
        When specified, the function will stop validation and return immediately upon
        finding the first null value, rather than validating all properties.

        .OUTPUTS
        System.Management.Automation.PSCustomObject
        Returns an object with the following properties:
        - IsValid: Boolean indicating if all validations passed (no null values found)
        - NullProperties: Array of property paths that contain null values
        - ErrorCount: Total number of properties with null values
        - Summary: Human-readable summary of validation results
        - JsonData: The loaded JSON object (if validation passes)

        .EXAMPLE
        $validationResult = Test-JsonNullValues -JsonFilePath "config.json" -RequiredProperties @("database.host", "database.port", "api.key") -JsonObjectName "Configuration"

        if (-not $validationResult.IsValid) {
            Write-Host "Validation failed: $($validationResult.Summary)"
            return
        }

        .EXAMPLE
        $requiredProps = @(
            "common.vCenterName",
            "common.VcenterUser",
            "common.esxHost"
        )
        $result = Test-JsonNullValues -JsonFilePath "input.json" -RequiredProperties $requiredProps -JsonObjectName "InputConfiguration"

        .NOTES
        - This function is designed to work in conjunction with Test-JsonMissingProperties
        - First check if keys exist (Test-JsonMissingProperties), then check if values are non-null (Test-JsonNullValues)
        - Uses Get-JsonPropertyValue to retrieve nested property values
        - Integrates with VCF PowerShell Toolbox logging infrastructure
        - Null values in arrays or objects are also detected
    #>

    Param (
        [Parameter(Mandatory = $false)] [Switch]$StopOnFirstError,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonFilePath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$JsonObjectName,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String[]]$RequiredProperties
    )

    Write-LogMessage -type DEBUG -message "Entered Test-JsonNullValues function..."

    # Initialize validation result object.
    $validationResult = [PSCustomObject]@{
        IsValid = $true
        NullProperties = @()
        ErrorCount = 0
        Summary = ""
        JsonData = $null
    }

    Write-LogMessage -type DEBUG -message "Checking $($RequiredProperties.Count) properties for null values: $($RequiredProperties -join ', ')"

    # Load and validate the JSON file using helper function.
    $jsonData = Get-JsonDataWithValidation -JsonFilePath $JsonFilePath -JsonObjectName $JsonObjectName -ValidationResult ([ref]$validationResult)
    if ($null -eq $jsonData) {
        return $validationResult
    }

    # Validate each property for null values.
    foreach ($Property in $RequiredProperties) {
        # Use Get-JsonPropertyValue to retrieve the property value.
        $propertyValue = Get-JsonPropertyValue -InputData $jsonData -PropertyPath $Property

        # Check if the value is null.
        if ($null -eq $propertyValue) {
            $validationResult.IsValid = $false
            $validationResult.NullProperties += $Property
            $validationResult.ErrorCount++

            Write-LogMessage -type ERROR -message "$JsonObjectName (in JSON file $JsonFilePath) property '$Property' has a null value. Please provide a valid value."

            # Stop on first error if requested.
            if ($StopOnFirstError) {
                break
            }
        }
    }

    # Generate summary message.
    if ($validationResult.IsValid) {
        $validationResult.Summary = "$JsonObjectName null value validation passed. All $($RequiredProperties.Count) required properties have non-null values."
        Write-LogMessage -type DEBUG -message $validationResult.Summary
    }
    else {
        $validationResult.Summary = "$JsonObjectName null value validation failed. $($validationResult.ErrorCount) of $($RequiredProperties.Count) required properties have null values: $($validationResult.NullProperties -join ', ')"
        Write-LogMessage -type ERROR -message $validationResult.Summary
    }

    # Store the JSON data in the result for caller use.
    $validationResult.JsonData = $jsonData

    return $validationResult
}
Function Get-CleanErrorMessage {

    <#
        .SYNOPSIS
        Extracts clean error messages from JSON error responses.

        .DESCRIPTION
        The Get-CleanErrorMessage function attempts to extract localized or default error
        messages from JSON-formatted error responses. This function standardizes error message
        extraction throughout the module, eliminating code duplication and ensuring consistent
        error message handling.

        The function checks for error messages in the following priority order:
        1. "localized" field - User-friendly localized error message
        2. "default_message" field - Default error message
        3. Original error message - Falls back to the input if no clean message is found

        This function is used throughout the module to extract clean, user-friendly error
        messages from API responses that may contain JSON-formatted error details.

        .PARAMETER ErrorMessage
        The raw error message that may contain JSON-formatted error details. This can be
        a plain string or a JSON string containing error information.

        .EXAMPLE
        $cleanError = Get-CleanErrorMessage -ErrorMessage $_.Exception.Message
        Write-LogMessage -type ERROR -message "Operation failed: $cleanError"

        Extracts a clean error message from an exception and logs it.

        .EXAMPLE
        $cleanMessage = Get-CleanErrorMessage -ErrorMessage $errorResponse
        if ($cleanMessage) {
            Write-Host "Error: $cleanMessage"
        }

        Extracts a clean error message from an API response for display to the user.

        .OUTPUTS
        System.String
        Returns the cleanest available error message. If no clean message is found in the
        JSON response, returns the original error message unchanged.

        .NOTES
        This function uses regex pattern matching to extract error messages from JSON strings.
        The patterns match common JSON error response formats used by vCenter and VCF APIs.

        Error Message Priority:
        - "localized" field is preferred as it provides user-friendly messages
        - "default_message" field is used if "localized" is not available
        - Original message is returned if neither field is found
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$ErrorMessage
    )

    switch -Regex ($ErrorMessage) {
        '"localized":"([^"]+)"' {
            return $matches[1]
        }
        '"default_message":"([^"]+)"' {
            return $matches[1]
        }
        default {
            return $ErrorMessage
        }
    }
}
Function ConvertFrom-SecureString {

    <#
        .SYNOPSIS
        Converts a SecureString to a plain text string for API authentication.

        .DESCRIPTION
        Safely converts a SecureString to plain text using secure memory operations.
        The plain text is extracted and the secure memory is immediately cleared.
        This function should only be used when plain text is absolutely required for API calls.

        .PARAMETER SecureString
        The SecureString to convert to plain text.

        .OUTPUTS
        String
        Returns the plain text password string.

        .EXAMPLE
        $securePassword = Read-Host "Enter password" -AsSecureString
        $plainText = ConvertFrom-SecureString -SecureString $securePassword

        Converts a SecureString obtained from user input to plain text for API authentication.

        .NOTES
        SECURITY: This function converts SecureString to plain text, which is less secure.
        - Uses Marshal operations for secure memory handling
        - Automatically clears secure memory after conversion
        - Should be used only when plain text is required for API calls
        - Caller is responsible for clearing the returned plain text string immediately after use
    #>

    Param (
        [Parameter(Mandatory = $true)] [System.Security.SecureString]$SecureString
    )

    $decodedPasswordInterimStep = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($SecureString)
    $plainText = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($decodedPasswordInterimStep)
    [System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($decodedPasswordInterimStep)
    return $plainText
}
Function Test-FileLocked {

    <#
        .SYNOPSIS
        Tests if a file is locked by another process.

        .DESCRIPTION
        Attempts to open a file with read/write access to determine if it is locked
        by another process. If the file cannot be opened, it is considered locked.
        This function uses proper resource disposal to prevent memory leaks.

        .PARAMETER FilePath
        The full path to the file to test.

        .EXAMPLE
        Test-FileLocked -FilePath "C:\ISO\file.iso"
        Returns $true if the file is locked, $false otherwise.

        .OUTPUTS
        Boolean
        Returns $true if the file is locked, $false if it is not locked or does not exist.

        .NOTES
        This function attempts to open the file with exclusive access. If the file
        does not exist, the function returns $false.
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$FilePath
    )

    if (-not (Test-Path $FilePath)) {
        return $false
    }

    $fileStream = $null
    try {
        $fileStream = [System.IO.File]::Open($FilePath, 'Open', 'ReadWrite', 'None')
        $fileStream.Close()
        return $false
    } catch {
        return $true
    } finally {
        if ($fileStream) {
            try {
                $fileStream.Dispose()
            } catch {
                # Ignore cleanup errors - file stream may already be closed or disposed.
            }
        }
    }
}
Function Test-DiskSpace {

    <#
        .SYNOPSIS
        Checks if sufficient disk space is available for a file operation.

        .DESCRIPTION
        The Test-DiskSpace function checks the available free space on the drive where
        the specified file path is located and compares it against the required size.
        The function includes a safety buffer to account for filesystem overhead and
        other operations.

        .PARAMETER FilePath
        The full path to the file that will be created. The function will determine
        the drive from this path.

        .PARAMETER MinimumBufferBytes
        Minimum buffer in bytes to require even if RequiredSize is not specified.
        Defaults to 104857600 (100MB).

        .PARAMETER RequiredSize
        The size in bytes required for the operation. If not specified, a minimum
        buffer (100MB) is checked.

        .PARAMETER SafetyBufferPercent
        Percentage of additional space to require as a safety buffer. Defaults to 10%.

        .EXAMPLE
        Test-DiskSpace -FilePath "C:\ISO\file.iso" -RequiredSize 1073741824
        Checks if at least 1GB (plus 10% buffer) is available on C: drive.

        .EXAMPLE
        Test-DiskSpace -FilePath "D:\Downloads\file.iso" -RequiredSize 524288000 -SafetyBufferPercent 20
        Checks if at least 500MB (plus 20% buffer) is available on D: drive.

        .OUTPUTS
        Hashtable
        Returns a hashtable with the following keys:
        - HasEnoughSpace: Boolean indicating if sufficient space is available
        - AvailableSpace: Int64 - Available free space in bytes
        - RequiredSpace: Int64 - Required space (including buffer) in bytes
        - Drive: String - Drive letter or path where the file will be saved
        - ErrorMessage: String - Error message if check failed (null on success)

        .NOTES
        The function uses Get-PSDrive to get disk space information, which works on
        Windows, macOS, and Linux. The safety buffer helps prevent failures due to
        filesystem overhead, temporary files, or concurrent operations.
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$FilePath,
        [Parameter(Mandatory = $false)] [ValidateRange(0, [Int64]::MaxValue)] [Int64]$MinimumBufferBytes = 104857600,
        [Parameter(Mandatory = $false)] [ValidateRange(0, [Int64]::MaxValue)] [Int64]$RequiredSize = 0,
        [Parameter(Mandatory = $false)] [ValidateRange(0, 100)] [int]$SafetyBufferPercent = 10
    )

    Write-LogMessage -type DEBUG -message "Entered Test-DiskSpace function for file: $FilePath"

    try {
        # Get the drive root from the file path.
        $driveRoot = Split-Path -Path $FilePath -Qualifier
        if ([string]::IsNullOrEmpty($driveRoot)) {
            # For Unix-like systems, get the root directory.
            $driveRoot = (Split-Path -Path $FilePath -Parent)
            while ($driveRoot -ne (Split-Path -Path $driveRoot -Parent)) {
                $driveRoot = Split-Path -Path $driveRoot -Parent
            }
        }

        # Get drive information.
        $drive = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Root -eq $driveRoot }
        if (-not $drive) {
            # Try alternative method for Unix systems or network paths.
            $driveInfo = [System.IO.DriveInfo]::new($driveRoot)
            $availableSpace = $driveInfo.AvailableFreeSpace
            $driveName = $driveInfo.Name
        } else {
            $availableSpace = $drive.Free
            $driveName = $drive.Name
        }

        # Calculate required space (file size + safety buffer).
        $bufferSize = if ($RequiredSize -gt 0) {
            [Math]::Max($RequiredSize * $SafetyBufferPercent / 100, $MinimumBufferBytes)
        } else {
            $MinimumBufferBytes
        }
        $requiredSpace = $RequiredSize + $bufferSize

        # Format bytes for human-readable output
        $availableSpaceMB = [math]::Round($availableSpace / 1MB, 2)
        $requiredSpaceMB = [math]::Round($requiredSpace / 1MB, 2)
        $bufferSizeMB = [math]::Round($bufferSize / 1MB, 2)

        Write-LogMessage -type DEBUG -message "Drive: $driveName, Available: $availableSpaceMB MB ($availableSpace bytes), Required: $requiredSpaceMB MB ($requiredSpace bytes)"

        # Check if enough space is available.
        if ($availableSpace -lt $requiredSpace) {
            return @{
                HasEnoughSpace = $false
                AvailableSpace = $availableSpace
                RequiredSpace = $requiredSpace
                Drive = $driveName
                ErrorMessage = "Insufficient disk space on drive '$driveName'. Available: $availableSpaceMB MB ($availableSpace bytes), Required: $requiredSpaceMB MB ($requiredSpace bytes) (including $bufferSizeMB MB safety buffer)"
            }
        }

        return @{
            HasEnoughSpace = $true
            AvailableSpace = $availableSpace
            RequiredSpace = $requiredSpace
            Drive = $driveName
            ErrorMessage = $null
        }
    }
    catch {
        Write-LogMessage -type ERROR -message "Error checking disk space for file '$FilePath': $($_.Exception.Message)"
        return @{
            HasEnoughSpace = $false
            AvailableSpace = 0
            RequiredSpace = $requiredSpace
            Drive = "Unknown"
            ErrorMessage = "Failed to check disk space: $($_.Exception.Message)"
        }
    }
}
Function Test-Filepath {

    <#
        .SYNOPSIS
        Tests if a specified file exists at the given file path.

        .DESCRIPTION
        The function Test-Filepath validates whether a file exists at the specified path.
        If the file exists, it logs a success message. If the file does not exist,
        it logs an error message and throws an exception to stop script execution.

        .PARAMETER Description
        A descriptive name for the file being tested, used in log messages.

        .PARAMETER FilePath
        The absolute path to the file that needs to be validated for existence.

        .EXAMPLE
        Test-Filepath -FilePath "c:\argocd.yml" -Description "ArgoCD configuration"

        .NOTES
        This function throws an exception if the file is not found, which will stop
        script execution. Use try-catch blocks if you need to handle this gracefully.
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Description,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$FilePath
    )

    Write-LogMessage -type DEBUG -message "Entered Test-Filepath function..."

    if (Test-Path -Path $FilePath -PathType Leaf) {
        Write-LogMessage -type INFO -message "Found the `"$Description`" file on disk: `"$FilePath`"."
    } else {
        Write-LogMessage -type ERROR -message "Failed to find `"$Description`" file on disk: `"$FilePath`" not found. Exiting."
        throw "Deployment failed. Check logs for details."
    }
}
Function Test-CommandAvailability {

    <#
        .SYNOPSIS
        Tests if a specified command/utility is available in the system PATH.

        .DESCRIPTION
        This function checks whether a given command or executable is available and accessible
        through the system PATH. It can be used to verify that required tools or utilities
        are installed before attempting to use them in the script. If the command is not found,
        the function will log an error and throw an exception to stop script execution.

        .PARAMETER Command
        The name of the command or executable to test for availability.

        .PARAMETER Description
        A human-readable description of the command for use in error messages.

        .EXAMPLE
        Test-CommandAvailability -Command "vcf" -Description "vcf-cli"

        .NOTES
        This function throws an exception if the command is not found, which will stop
        script execution. Use try-catch blocks if you need to handle this gracefully.
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Command,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$Description
    )

    Write-LogMessage -type DEBUG -message "Entered Test-CommandAvailability function..."

    if (Get-Command $Command -ErrorAction SilentlyContinue) {
        Write-LogMessage -type INFO -suppressOutputToScreen -message "Executable $Command found in PATH. Proceeding."
    } else {
        Write-LogMessage -type ERROR -message "Executable `"$Command`" not found in PATH.  $Description is required for the script to proceed. Exiting"
        throw "Deployment failed. Check logs for details."
    }
}
Function Test-IpAddressInCidrRange {

    <#
        .SYNOPSIS
        Tests if an IP address falls within a specified CIDR network range.

        .DESCRIPTION
        The Test-IpAddressInCidrRange function validates whether a given IP address
        is contained within a specified CIDR network range. This is useful for validating
        that starting IP addresses, gateway addresses, or other IP configurations fall
        within expected network boundaries.

        The function performs the following validation:
        1. Validates the format of both the IP address and CIDR notation
        2. Parses the CIDR range to extract network address and subnet mask
        3. Converts both IP addresses to binary format for comparison
        4. Applies the subnet mask to determine network membership
        5. Returns true if the IP is within the range, false otherwise

        .PARAMETER CidrRange
        The CIDR network range (e.g., "192.168.1.0/24"). Must be in valid CIDR notation
        with format: IP/prefix where prefix is 0-32.

        .PARAMETER IpAddress
        The IP address to test (e.g., "192.168.1.100"). Must be a valid IPv4 address.

        .EXAMPLE
        Test-IpAddressInCidrRange -IpAddress "192.168.1.100" -CidrRange "192.168.1.0/24"
        Returns $true because 192.168.1.100 is within the 192.168.1.0/24 network.

        .EXAMPLE
        Test-IpAddressInCidrRange -IpAddress "10.0.0.5" -CidrRange "192.168.1.0/24"
        Returns $false because 10.0.0.5 is not within the 192.168.1.0/24 network.

        .EXAMPLE
        Test-IpAddressInCidrRange -IpAddress "172.16.50.1" -CidrRange "172.16.0.0/16"
        Returns $true because 172.16.50.1 is within the 172.16.0.0/16 network.

        .OUTPUTS
        Boolean
        Returns $true if the IP address is within the CIDR range, $false otherwise.

        .NOTES
        This function only supports IPv4 addresses and CIDR notation.
        The function validates input formats before performing range checks.
    #>

    Param(
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$CidrRange,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$IpAddress
    )

    Write-LogMessage -type DEBUG -message "Entered Test-IpAddressInCidrRange function..."

    try {
        # Validate IP address format.
        if ($IpAddress -notmatch '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$') {
            Write-LogMessage -type ERROR -message "Invalid IP address format: $IpAddress"
            return $false
        }

        # Validate CIDR range format.
        if ($CidrRange -notmatch '^((25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\/([0-9]|[1-2][0-9]|3[0-2])$') {
            Write-LogMessage -type ERROR -message "Invalid CIDR range format: $CidrRange"
            return $false
        }

        # Split CIDR into network address and prefix length.
        $cidrParts = $CidrRange.Split('/')
        $networkAddress = $cidrParts[0]
        $prefixLength = [int]$cidrParts[1]

        # Convert IP addresses to 32-bit integers.
        Function ConvertTo-IpInt {
            param([String]$IpString)
            $octets = $IpString.Split('.')
            return ([int64]$octets[0] -shl 24) -bor ([int64]$octets[1] -shl 16) -bor ([int64]$octets[2] -shl 8) -bor [int64]$octets[3]
        }

        # Calculate subnet mask from prefix length.
        if ($prefixLength -eq 0) {
            $subnetMask = 0
        } else {
            $subnetMask = [int64][Math]::Pow(2, 32) - [int64][Math]::Pow(2, (32 - $prefixLength))
        }

        # Convert addresses to integers.
        $ipInt = ConvertTo-IpInt -IpString $IpAddress
        $networkInt = ConvertTo-IpInt -IpString $networkAddress

        # Apply subnet mask to both addresses.
        $ipNetwork = $ipInt -band $subnetMask
        $cidrNetwork = $networkInt -band $subnetMask

        # Check if the IP is in the same network.
        $isInRange = ($ipNetwork -eq $cidrNetwork)

        if ($isInRange) {
            Write-LogMessage -type DEBUG -message "IP address $IpAddress is within CIDR range $CidrRange"
        } else {
            Write-LogMessage -type DEBUG -message "IP address $IpAddress is not within CIDR range $CidrRange"
        }

        return $isInRange
    }
    catch {
        Write-LogMessage -type ERROR -message "Error validating IP address in CIDR range: $($_.Exception.Message)"
        return $false
    }
}
Function Test-ValidCidrRange {

    <#
        .SYNOPSIS
        Validates that an IP count corresponds to a valid CIDR block range.

        .DESCRIPTION
        The Test-ValidCidrRange function checks if a given IP address count corresponds to a valid
        CIDR range (/8 to /32). The value must be a power of 2 AND within the valid range.
        This ensures IP address counts correspond to complete, valid CIDR blocks.

        Valid CIDR ranges (IPv4):
        - 1 IP = 2^0 = /32 (single host)
        - 2 IPs = 2^1 = /31 (point-to-point)
        - 4 IPs = 2^2 = /30
        - 8 IPs = 2^3 = /29
        - 16 IPs = 2^4 = /28
        - 32 IPs = 2^5 = /27
        - 64 IPs = 2^6 = /26
        - 128 IPs = 2^7 = /25
        - 256 IPs = 2^8 = /24
        - 512 IPs = 2^9 = /23
        - 1024 IPs = 2^10 = /22
        - ... up to ...
        - 16,777,216 IPs = 2^24 = /8 (maximum)

        Values larger than 16,777,216 (e.g., 2^25 = 33,554,432) are powers of 2 but correspond
        to CIDR prefixes smaller than /8, which are invalid.

        .PARAMETER InputText
        The value to validate as a power of 2.

        .PARAMETER PropertyPath
        Optional. The property path for error messages.

        .OUTPUTS
        System.Boolean
        Returns $true if the value is a power of 2 within valid CIDR range, $false otherwise.

        .EXAMPLE
        $isValid = Test-ValidCidrRange -InputText "512"
        Validates that "512" corresponds to a valid CIDR range (/23).
        Returns: $true

        .EXAMPLE
        $isValid = Test-ValidCidrRange -InputText "511"
        Validates that "511" corresponds to a valid CIDR range.
        Returns: $false (511 is not a power of 2)

        .EXAMPLE
        $isValid = Test-ValidCidrRange -InputText "33554432"
        Validates that "33554432" corresponds to a valid CIDR range.
        Returns: $false (would be /7, outside valid range)

        .NOTES
        The function uses bitwise AND operation to check if a number is a power of 2.
        A power of 2 in binary has exactly one bit set (e.g., 8 = 1000, 16 = 10000).
        The check (n & (n-1)) == 0 returns true only for powers of 2.
    #>

    Param (
        [Parameter(Mandatory = $false)] [ValidateNotNullOrEmpty()] [String]$PropertyPath,
        [Parameter(Mandatory = $true)] [ValidateNotNullOrEmpty()] [String]$InputText
    )

    Write-LogMessage -type DEBUG -message "Entered Test-ValidCidrRange function..."

    Write-LogMessage -type DEBUG -message "Validating CIDR range for IP count: '$InputText'"

    # Attempt to parse as integer.
    $number = $null
    $isInteger = [int]::TryParse($InputText, [ref]$number)

    if (-not $isInteger) {
        $pathInfo = if ($PropertyPath) { " for property '$PropertyPath'" } else { "" }
        Write-LogMessage -type ERROR -message "CIDR range validation failed${pathInfo}: Value '$InputText' is not a valid integer"
        return $false
    }

    # Check if number is positive.
    if ($number -le 0) {
        $pathInfo = if ($PropertyPath) { " for property '$PropertyPath'" } else { "" }
        Write-LogMessage -type ERROR -message "CIDR range validation failed${pathInfo}: Value $number must be positive"
        return $false
    }

    # Check if number is a power of 2 using bitwise AND.
    # A power of 2 has only one bit set in binary representation.
    # Example: 8 = 1000, 8-1 = 0111, 1000 & 0111 = 0000.
    # Non-power: 7 = 0111, 7-1 = 0110, 0111 & 0110 = 0110 (not zero)
    $isPowerOfTwo = ($number -band ($number - 1)) -eq 0

    if (-not $isPowerOfTwo) {
        $pathInfo = if ($PropertyPath) { " for property '$PropertyPath'" } else { "" }
        Write-LogMessage -type ERROR -message "CIDR range validation failed${pathInfo}: Value $number is not a power of 2"
        return $false
    }

    # Check if the power of 2 corresponds to a valid CIDR prefix (/8 to /32).
    # Maximum valid: 2^24 = 16,777,216 (/8)
    # Minimum valid: 2^0 = 1 (/32)
    $maxValidCidr = [Math]::Pow(2, 24)
    if ($number -gt $maxValidCidr) {
        $pathInfo = if ($PropertyPath) { " for property '$PropertyPath'" } else { "" }
        Write-LogMessage -type ERROR -message "CIDR range validation failed${pathInfo}: Value $number exceeds maximum valid CIDR range (16,777,216 = /8)"
        return $false
    }

    Write-LogMessage -type DEBUG -message "CIDR range validation passed: $InputText is a valid power of 2"
    return $true
}
Function Write-ExceptionDetails {

    <#
        .SYNOPSIS
        Logs detailed exception information including inner exceptions.

        .DESCRIPTION
        Traverses exception chain and logs detailed information about each exception,
        including type, message, source, and stack trace. Also detects common exception
        types (SSL, network, authentication) and provides additional context.

        .PARAMETER Exception
        The exception object to log details for.

        .PARAMETER LogType
        Log message type (ERROR, WARNING, DEBUG). Default is ERROR.

        .PARAMETER MaxDepth
        Maximum depth to traverse inner exceptions. Default is 5.

        .EXAMPLE
        try {
            # Some operation
        } catch {
            Write-ExceptionDetails -Exception $_.Exception
        }

        Logs detailed exception information including inner exceptions.

        .NOTES
        - Automatically detects SSL/TLS, network, and authentication exceptions.
        - Provides additional context for common exception types.
    #>

    Param (
        [Parameter(Mandatory = $true)] [ValidateNotNull()] [System.Exception]$Exception,
        [Parameter(Mandatory = $false)] [ValidateSet('ERROR', 'WARNING', 'DEBUG')] [string]$LogType = 'ERROR',
        [Parameter(Mandatory = $false)] [ValidateRange(1, 10)] [int]$MaxDepth = 5
    )

    Write-LogMessage -type $LogType -message "Exception Message: $($Exception.Message)"
    Write-LogMessage -type $LogType -message "Exception Type: $($Exception.GetType().FullName)"
    if ($Exception.Source) {
        Write-LogMessage -type $LogType -message "Exception Source: $($Exception.Source)"
    }
    if ($Exception.StackTrace) {
        Write-LogMessage -type DEBUG -message "Exception StackTrace: $($Exception.StackTrace)"
    }

    $currentException = $Exception
    $depth = 0
    while ($currentException.InnerException -and $depth -lt $MaxDepth) {
        $depth++
        $currentException = $currentException.InnerException
        Write-LogMessage -type $LogType -message "Inner exception (depth $depth): $($currentException.Message)"
        Write-LogMessage -type $LogType -message "Inner exception type: $($currentException.GetType().FullName)"
        if ($currentException.Source) {
            Write-LogMessage -type $LogType -message "Inner exception source: $($currentException.Source)"
        }

        if ($currentException -is [System.Net.Http.HttpRequestException]) {
            Write-LogMessage -type $LogType -message "HttpRequestException detected - this is typically an SSL/TLS issue"
        }
        if ($currentException -is [System.Security.Authentication.AuthenticationException]) {
            Write-LogMessage -type $LogType -message "AuthenticationException detected - SSL handshake failed"
        }
        if ($currentException -is [System.Net.Sockets.SocketException]) {
            Write-LogMessage -type $LogType -message "SocketException detected - network connectivity issue"
            if ($currentException.SocketErrorCode) {
                Write-LogMessage -type $LogType -message "SocketErrorCode: $($currentException.SocketErrorCode)"
            }
        }
        if ($currentException.StackTrace) {
            Write-LogMessage -type DEBUG -message "Inner exception stack trace: $($currentException.StackTrace)"
        }
    }
}