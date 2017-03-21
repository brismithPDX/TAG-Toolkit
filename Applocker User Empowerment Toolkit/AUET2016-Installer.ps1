#===========================================================================
# Set up Dependencies and Variables
#===========================================================================
param($InstallModes)

#Core Configuration Variables
[bool]$Dev = $false
[bool] $InstallLog = "$env:TEMP\AUETInstall.log"

Function Write-Log {
    [CmdletBinding()]
    Param(
    [Parameter(Mandatory=$False)]
    [ValidateSet("INFO","WARN","ERROR","FATAL","DEBUG")]
    [String]
    $Level = "INFO",

    [Parameter(Mandatory=$True)]
    [string]
    $Message,

    [Parameter(Mandatory=$False)]
    [string]
    $logfile
    )

    $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
    $Line = "$Stamp $Level $Message"
    If($logfile) {
        Add-Content $logfile -Value $Line
    }
    Else {
        Write-Output $Line
    }
}
#===========================================================================
# Perform Installation
#===========================================================================

Switch ($InstallModes){
    1 {# Installation
        if ($Dev){
            Write-Log -Level INFO -Message "[DEV-MODE] InstallMode = $InstallModes Begining Installation" -logfile $InstallLog
        }
        Write-Log -Level INFO -Message "Begining Installation" -logfile $InstallLog
        # Create Required Directories
        try{
            New-Item -ItemType Directory -Force -Path $env:ProgramFiles\AUET2016
            New-Item -ItemType Directory -Force -Path $env:ProgramFiles\AUET2016\Management
            Write-Log -Level INFO -Message "Created application directories" -logfile $InstallLog
        }
        catch{
            Write-Log -Level ERROR -Message "Creating application directories FAILED" -logfile $InstallLog
        }

        # Move install items to required directories
        try{
            Move-Item AUET2016-EventStrapper.ps1 $env:ProgramFiles\AUET2016\AUET2016-EventStrapper.ps1 -Force
            Move-Item AUET_BlockedEXE_Resolver_BootstrapingTask.xml $env:ProgramFiles\AUET2016\Management\AUET_BlockedEXE_Resolver_BootstrapingTask.xml -Force
            Write-Log -Level INFO -Message "Installing application files" -logfile $InstallLog
        }
        catch{
            Write-Log -Level ERROR -Message "Installing application FAILED" -logfile $InstallLog
        }

        # Register new Required Task Scheduler item
        try{
            Register-ScheduledTask -Xml (Get-Content $env:ProgramFiles\AUET2016\Management\AUET_BlockedEXE_Resolver_BootstrapingTask.xml | out-string) -TaskName "AUET 2016 bootstrapping Trigger"
            Write-Log -Level INFO -Message "Registered application task scheduling trigger" -logfile $InstallLog
        }
        catch{
            Write-Log -Level ERROR -Message "Registering application task scheduling trigger FAILED" -logfile $InstallLog
        }
    }
    2 {# Un-Installation
        if ($Dev){
                Write-Log -Level INFO -Message "[DEV-MODE] InstallMode = $InstallModes Begining Un-Installation" -logfile $InstallLog
            }
        Write-Log -Level INFO -Message "Begining Un-Installation" -logfile $InstallLog
        # Delete Application Directories
        try{
            New-Item -ItemType Directory -Force -Path $env:ProgramFiles\AUET2016
            New-Item -ItemType Directory -Force -Path $env:ProgramFiles\AUET2016\Management
            Write-Log -Level INFO -Message "Removed application directories" -logfile $InstallLog
        }
        catch{
            Write-Log -Level ERROR -Message "Removing application directories FAILED" -logfile $InstallLog
        }

        # Remove Required Task Scheduler item
        try{
            Unregister-ScheduledTask -TaskName "AUET 2016 bootstrapping Trigger" -Confirm $false
            Write-Log -Level INFO -Message "Removed application task scheduling trigger" -logfile $InstallLog
        }
        catch{
            Write-Log -Level ERROR -Message "Removing application task scheduling trigger FAILED" -logfile $InstallLog
        }
    }
    default {# Strange Case
        if ($Dev){
                Write-Log -Level WARN -Message "[DEV-MODE] Could not understand install mode: $InstallModes Exiting Installation" -logfile $InstallLog
            }
        Write-Log -Level WARN -Message "Could not understand install mode: $InstallModes Exiting Installation" -logfile $InstallLog
        
    }

}