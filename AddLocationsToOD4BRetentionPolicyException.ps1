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
#>

# Add Type and Load the Systems Forms
Add-Type -AssemblyName System.Windows.Forms
$csvpath = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Filter = 'CSV (*.csv)| *.csv' 
}

# Show the Dialog
$null = $csvpath.ShowDialog()

function import-ValidCSV
{
        param
        (
                [parameter(Mandatory=$true)]
                [ValidateScript({test-path $_ -type leaf})]
                [string]$inputFile,
                [string[]]$requiredColumns
        )
        $csvImport = import-csv -LiteralPath $inputFile
        $inputTest = $csvImport | Get-Member
        foreach ($requiredColumn in $requiredColumns)
        {
                if (!($inputTest | Where-Object {$_.name -eq $requiredColumn}))
                {
                        write-error "$inputFile is missing the $requiredColumn column"
                        exit 10
                }
        }

        $csvImport
}

$Users = import-ValidCSV -inputFile $csvpath.FileName -requiredColumns "OneDriveUrl"

$UPN = Read-Host -Prompt 'Enter your O365 Login'
Import-Module ExchangeOnlineManagement
Connect-IPPSSession -UserPrincipalName $UPN
$ErrorLog = 'Errors.txt'

try {
    $OneDrivePolicy = Get-RetentionCompliancePolicy -Identity "Test Retention" -ErrorAction stop 
}
catch {
    Write-Host "Retention Policy not found" -ForegroundColor Yellow
    exit 10
} 

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