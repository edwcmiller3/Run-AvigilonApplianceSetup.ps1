# Run-AvigilonApplianceSetup.ps1
Script to automate the configuration of an Avigilon HD Video Appliance

Script with GUI to configure appliance according to the needs of our K-12 clients.
Because of the default Execution Policy in PowerShell on Windows 7 Embedded, this is run using a batch file that calls the PowerShell script.

@ECHO OFF

PowerShell.exe -ExecutionPolicy Bypass -Command "& '\\path\to\script\Run-AvigilonApplianceSetup.ps1'"

PAUSE
