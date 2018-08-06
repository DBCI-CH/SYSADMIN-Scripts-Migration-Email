set-executionpolicy remotesigned
$CSVfileandpath = "c:\temp\delegation_comma.csv"
get-Member $_.primarymailboxuser 