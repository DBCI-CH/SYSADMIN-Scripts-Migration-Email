
set-executionpolicy remotesigned
#requires -runasadministrator

#function to browse for CSV file from Windows Explorer (cleaner input)

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}



$usercredential=get-credential
$inputfile = Get-FileName "C:\temp"
$inputdata = get-content $inputfile

$calendar = $_.primarymailboxemail +":\calendar"
$calendrier = $_.primarymailboxemail + ":\calendrier"

#$CSVfileandpath = "c:\temp\delegation_comma.csv"

$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
import-pssession $Session


import-csv $inputfile | foreach {Add-MailboxfolderPermission -identity ($_.PrimaryMailboxEmail + ":\calendrier") -User $_.delegate -AccessRights $_.delegaterights}
import-csv $inputfile | foreach {get-mailboxfolderpermission -identity ($_.primarymailboxemail + ":\calendrier")} | select identity,user,accessrights

import-csv $inputfile | ForEach {Add-MailboxFolderPermission -Identity ($_.PrimaryMailboxEmail + ":\Calendar") -User $_.delegate -AccessRights $_.DelegateRights}
import-csv $inputfile | foreach {get-mailboxfolderpermission -identity ($_.primarymailboxemail + ":\calendar") } | select identity,user,accessrights

