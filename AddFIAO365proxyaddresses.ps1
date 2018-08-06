# #requires -runasadministrator
#defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}

$inputfile = Get-FileName "C:\temp"
$inputdata = get-content $inputfile

set-executionpolicy remotesigned
import-module activedirectory
import-csv $inputfile  | Foreach {Set-ADUser -identity $_.username -Add @{'ProxyAddresses'=@(("smtp:{0}@{1}"-f $_.samaccountname, 'fiao365.onmicrosoft.com'))} }
import-csv $inputfile | Foreach {Set-ADUser -identity $_.username -replace @{mailnickname=$_.username}}
import-csv $inputfile | Foreach {Set-ADUser -identity $_.username -replace @{'targetaddress'=@("SMTP:{0}@{1}"-f $_.username, 'fia.com')}}