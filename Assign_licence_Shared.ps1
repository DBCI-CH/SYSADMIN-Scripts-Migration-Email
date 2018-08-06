

#requires -runasadministrator
#requires -module msonline
install-module msonline
#import-module activedirectory

Function Get-FileName ($initialDirectory)
{
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.initialDirectory = $initialDirectory
    $OpenFileDialog.filter = "CSV (*.csv)| *.csv"
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}



#source : https://community.spiceworks.com/scripts/show/1712-start-countdown 
Function Start-Countdown 
{   <#
    .SYNOPSIS
        Provide a graphical countdown if you need to pause a script for a period of time
    .PARAMETER Seconds
        Time, in seconds, that the function will pause
    .PARAMETER Messge
        Message you want displayed while waiting
    .EXAMPLE
        Start-Countdown -Seconds 30 -Message Please wait while Active Directory replicates data...
    .NOTES
        Author:            Martin Pugh
        Twitter:           @thesurlyadm1n
        Spiceworks:        Martin9700
        Blog:              www.thesurlyadmin.com
       
        Changelog:
           2.0             New release uses Write-Progress for graphical display while couting
                           down.
           1.0             Initial Release
    .LINK
        http://community.spiceworks.com/scripts/show/1712-start-countdown
    #>
    Param(
        [Int32]$Seconds = 10,
        [string]$Message = "Pausing for 10 seconds..."
    )
    ForEach ($Count in (1..$Seconds))
    {   Write-Progress -Id 1 -Activity $Message -Status "Waiting for $Seconds seconds, $($Seconds - $Count) left" -PercentComplete (($Count / $Seconds) * 100)
        Start-Sleep -Seconds 1
    }
    Write-Progress -Id 1 -Activity $Message -Status "Completed" -PercentComplete 100 -Completed
}


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


import-csv $inputfile | foreach {Set-MsolUser -UserPrincipalName $_.Primarymailboxemail  -UsageLocation $_.Location

    Write-host "Assignment of $LicenseNameE3 for entry $_.Primarymailboxemail ...$i/$l"
  Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -AddLicenses $LicenseNameE3 
  #disable unwanted plans 
  write-host "disable unused plans for entry $_.Primarymailboxemail ... $i/$l"
  set-msoluserlicense -UserPrincipalName $_.Primarymailboxemail -licenseoptions $options
  Write-host "Done."
  $i += 1

  }

$i =1


Write-Host "License assignment finished. We wait 3 minutes (sleep in script) before following to shared section (mailbox creation). "
Start-countdown -seconds 180 - Message "Please wait"


write-host "connect to exchange online over powershell"
write-host "set mailbox to shared and remove license"
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session

import-csv $inputfile | foreach {
 write-host "	Set Shared Mailbox Type for $_.Primarymailboxemail "
    Set-MailBox $_.Primarymailboxemail -type shared
    Start-Sleep -s 10
    write-host "Remove license for $_.Primarymailboxemail "
    Set-MsolUserLicense -UserPrincipalName $_.Primarymailboxemail  -RemoveLicenses $LicenseNameE3

  $i += 1}

Write-host "check script result for shared mailboxes"
import-csv $inputfile | foreach {get-mailbox -identity $_.primarymailboxemail | select name,userprincipalname,recipienttypedetails}
 


# Functions used in script 


