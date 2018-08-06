
#requires -runasadministrator



Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

set-executionpolicy remotesigned


$inputfile = Get-FileName "C:\temp"
$inputdata = get-content $inputfile
$usercredential=get-credential
#$CSVfileandpath = "c:\temp\delegation_comma.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

import-csv $inputfile | foreach {get-mailbox -identity $_.PrimaryMailboxEmail | set-mailbox -ForwardingSmtpAddress $_.migrationAddress -DeliverToMailboxAndForward $true}

#check script result
Write-host "validate scripts results"
import-csv $inputfile | foreach {get-mailbox -identity $_.PrimaryMailboxEmail | select name, ForwardingSmtpAddress, DeliverToMailboxAndForward}

