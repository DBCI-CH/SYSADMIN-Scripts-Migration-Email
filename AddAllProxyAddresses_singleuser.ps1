# #requires -runasadministrator
#defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned
import-module activedirectory

$username=[Microsoft.VisualBasic.Interaction]::InputBox("please input the user you want to add proxyaddresses to ","User to Add Proxy Adresses to","Accountingfia")
get-aduser -Filter {samaccountName -eq $username} -searchbase "dc=aitfia,dc=com"  | Foreach {Set-ADUser -identity $_.samaccountname -Add @{'ProxyAddresses'=@(("SMTP:{0}@{1}"-f $_.samaccountname, 'fia.com'),("smtp:{0}@{1}" -f $_.samaccountname, 'fiao365.onmicrosoft.com'),("smtp:{0}@{1}" -f $_.samaccountname, 'europe.fia.com'),("smtp:{0}@{1}" -f $_.samaccountname, 'fia.eu'),("smtp:{0}@{1}" -f $_.samaccountname, 'fiaf1medical.com'),("smtp:{0}@{1}" -f $_.samaccountname, 'fiaf1technical.com'),("smtp:{0}@{1}" -f $_.samaccountname, 'europe.fia.com'))} }
get-aduser -Filter {samaccountName -eq $username} -searchbase "dc=aitfia,dc=com" | Foreach {Set-ADUser -identity $_.samaccountname -Add @{mailnickname=$_.samaccountname}}
get-aduser -Filter {samaccountName -eq $username} -searchbase "dc=aitfia,dc=com" | Foreach {Set-ADUser -identity $_.samaccountname -replace @{'targetaddress'=@("SMTP:{0}@{1}"-f $_.samaccountname, 'fia.com')}}