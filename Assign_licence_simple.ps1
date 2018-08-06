

#requires -runasadministrator
#requires -module msonline
install-module msonline
#import-module activedirectory


#set window #defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned
$usercredential=Get-Credential -message "please confirm your Office365 Admin account (if different)" ("$env:username" + "@fia.com")
Connect-MsolService -Credential $UserCredential

#Sets path (default c:\temp to csv file input)
$inputfile = Get-FileName "C:\temp"
$users = get-content $inputfile

$LicenseNameE3="FIAO365:ENTERPRISEPACK"
$LicenseNameE1="FIAO365:STANDARDPACK"
#giveslicenseoptions to disable (all but exchange online) ref : https://docs.microsoft.com/en-us/office365/enterprise/powershell/disable-access-to-services-while-assigning-user-licenses
$options=new-msollicenseoptions -Accountskuid FIAO365:ENTERPRISEPACK -disabledplans "BPOS_S_TODO_2","FORMS_PLAN_E3","STREAM_O365_E3","Deskless","FLOW_O365_P2","POWERAPPS_O365_P2","Teams1","PROJECTWORKMANAGEMENT","YAMMER_ENTERPRISE","RMS_S_ENTERPRISE","OFFICESUBSCRIPTION","SWAY","INTUNE_O365","MCOSTANDARD","SHAREPOINTWAC","SHAREPOINTENTERPRISE"                      
           

$l = $users.count
$i = 1

  Write-host "Assignment of $LicenseNameE3 for $Primarymailboxemail ...$i/$l"
import-csv $inputfile | foreach {Set-MsolUser -UserPrincipalName $_.Primarymailboxemail  -UsageLocation $_.Location

  
  Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -AddLicenses $LicenseNameE3 
  #disable unwanted plans 
  write-host "disable unused plans for $_primarymailboxemail ... $i/$l"
  set-msoluserlicense -UserPrincipalName $_.Primarymailboxemail -licenseoptions $options
  Write-host "Done."
  $i += 1

  }

$i =1


if($_.mailboxtype -eq "Shared"){$IsShared = $True} 
if ($IsSHARED) {
import-csv $inputfile | foreach {
 write-host "	Set Shared Mailbox Type for $_.Primarymailboxemail "
    Set-MailBox $_.Primarymailboxemail -type shared
    Start-Sleep -s 10
    write-host "Remove license for $_.Primarymailboxemail "
    Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -RemoveLicenses $LicenseNameE3

  $i += 1}
  }
<#
  import-csv $inputfile | foreach {
write-host "mailbox checkup"
#>

Write-Host "License assignment finished. Wait for at least 5 minute before language configuration. Press any key to continue..."
$tt=read-host

<#
$i=1
$Users | forEach-Object {
  $Primarymailboxemail  = $_.PrimarySmtpAddress
  $SpecialMailbox=$_.RecipientTypeDetails
    
  if($SpecialMailbox -eq "SharedMailbox"){$IsShared = $True} else {$IsShared = $False}
  if($SpecialMailbox -eq "RoomMailbox"){$IsRoom   = $true} else {$IsRoom = $False}
  if($SpecialMailbox -eq "EquipmentMailbox"){$IsResource   = $true} else {$IsResource = $False}

  $Language = $_.Language
  $TimeZone = $_.TimeZone
  
  Write-host "Configuration of regitonal settings for $Primarymailboxemail  - TimeZone-$TimeZone | Language-$Language...$i/$l"
      
  Set-MailboxRegionalConfiguration $Primarymailboxemail  -TimeZone $TimeZone –Language $Language -LocalizeDefaultFolderName:$true
 
  if($IsShared)
    {
    write-host "	Set Shared Mailbox for $Primarymailboxemail "
    Set-MailBox $Primarymailboxemail  -type shared
    Start-Sleep -s 10
    Set-MsolUserLicense -UserPrincipalName $Primarymailboxemail  -RemoveLicenses $LicenseName
    }

  if($IsRoom)
    {
     write-host "	Set Room Mailbox for $Primarymailboxemail "
     Set-MailBox $Primarymailboxemail  -type Room
     Start-Sleep -s 10
     Set-MsolUserLicense -UserPrincipalName $Primarymailboxemail  -RemoveLicenses $LicenseName
     }
  if($IsResource)
    {
    write-host "	Set Resource Mailbox for $Primarymailboxemail "
    Set-MailBox $Primarymailboxemail  -type Equipment
    Start-Sleep -s 10
    Set-MsolUserLicense -UserPrincipalName $Primarymailboxemail  -RemoveLicenses $LicenseName
    }
  $i += 1
}

#>

# Functions used in script 

Function Get-FileName($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}