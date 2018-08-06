# #requires -runasadministrator
#defines system box windows assembly type without input
[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null

set-executionpolicy remotesigned
import-module activedirectory

$username=[Microsoft.VisualBasic.Interaction]::InputBox("please input the user you want to add proxyaddresses to ","User to Add Proxy Adresses to","Accountingfia")
get-aduser -Filter {samaccountName -eq $username} -searchbase "dc=aitfia,dc=com"  | Foreach {Set-ADUser -identity $_.samaccountname -Add @{'ProxyAddresses'=@(("smtp:{0}@{1}"-f $_.samaccountname, 'fiao365.onmicrosoft.com'))} }
