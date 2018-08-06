

#function to browse for CSV file from Windows Explorer (cleaner input

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

#requires -runasadministrator
set-executionpolicy remotesigned

$usercredential=get-credential
$inputfile = Get-FileName "C:\temp"
$inputdata = get-content $inputfile
#$CSVfileandpath = "c:\temp\delegation_comma.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session

import-csv $inputfile | foreach {get-mailbox $_.PrimaryMailboxEmail | set-mailbox -ForwardingSmtpAddress $null}

#check script result
Write-host "validate scripts results"
import-csv $inputfile | foreach {get-mailbox $_.PrimaryMailboxEmail | select name, ForwardingSmtpAddress, DeliverToMailboxAndForward}


