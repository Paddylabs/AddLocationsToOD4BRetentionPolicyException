<#
  .SYNOPSIS
  Adds Users OneDrives to a O365 Retention Policy exception list
  .DESCRIPTION
  Takes a CSV list of Users OneDrive urls and adds them to the specified O365 Retention Policy
  exception list.
  .PARAMETER
  None
  .EXAMPLE
  None
  .INPUTS
  A valid csv file of your choice with the header of "OneDriveUrl" and of course each users
  OneDriveURL in that column
  .OUTPUTS
  Errors.txt
  .NOTES
  Author:        Patrick Horne
  Creation Date: 24/08/20
  Requires:      
  Change Log:
  V1.0:         Initial Development
  V2.0:         Improved Functions from BrettMillerB
#>

#Requires -Module ExchangeOnlineManagement

[CmdletBinding()]
param (
    [Parameter(Mandatory)]
    [string]
    $UPN,

    [Parameter(Mandatory)]
    $PolicyName
)

# Declare functions to use in the script
function Get-OpenFileDialog {
    [CmdletBinding()]
    param (
        [string]
        $Directory = [Environment]::GetFolderPath('Desktop'),
        
        [string]
        $Filter = 'CSV (*.csv)| *.csv'
    )

    Add-Type -AssemblyName System.Windows.Forms
    $openFileDialog = [System.Windows.Forms.OpenFileDialog]::new()
    $openFileDialog.InitialDirectory = $Directory
    $openFileDialog.Filter = $Filter
    $openFileDialog.ShowDialog()
    $openFileDialog
}
function Import-ValidCSV {
    param (
        [parameter(Mandatory)]
        [ValidateScript({Test-Path $_ -type leaf})]
        [string]
        $inputFile,

        [string[]]
        $requiredColumns
    )

    $csvImport = Import-Csv -LiteralPath C:\temp\Users.csv # $inputFile
    $requiredColumns | ForEach-Object {
        if ($_ -notin $csvImport[0].psobject.properties.name) {
            Write-Error "$inputFile is missing the $_ column"
            exit 10
        }
    }

    $csvImport
}

# Prompt File Dialog Picker filtered on CSV files
$csvPath = Get-OpenFileDialog

# Import users from CSV and validate CSV Properties present
$Users = Import-ValidCSV -inputFile $csvpath.FileName -requiredColumns OneDriveUrl,Test

Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName $UPN

try {
    $OneDrivePolicy = Get-RetentionCompliancePolicy -Identity $PolicyName -ErrorAction stop 
}
catch {
    Write-Host "The specified Retention Policy was not found" -ForegroundColor Yellow
    exit 10
}

$ErrorLog = 'Errors.txt'

foreach ($user in $users) {

        try {
            Set-RetentionCompliancePolicy -Identity $OneDrivePolicy.name -AddOneDriveLocationException $user.OneDriveUrl -ErrorAction stop
        }
        catch {
            $ErrorMessage = $_.Exception.Message
            Write-Host $ErrorMessage -ForegroundColor Red
            $ErrorMessage | Out-File -FilePath $ErrorLog -Append
            
        } 
}

Get-RetentionCompliancePolicy -Identity $OneDrivePolicy.name -DistributionDetail | Select-Object -ExpandProperty OneDriveLocationException