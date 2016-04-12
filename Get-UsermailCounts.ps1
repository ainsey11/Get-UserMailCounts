############################################################
# Gets Mail Counts for a daily breakdown, sends to users 
# defined in $recipients variable
# Author : Robert Ainsworth
# Web : https://ainsey11.com
###########################################################

$UserList = "C:\Scripting\Get-UserDailyMailCounts\Userlist.csv" # put a list of mailboxes in here, see csv in repo for example
$temp = "C:\Scripting\Get-UserDailyMailCounts\UserMailCounts.txt"
$body = "Please see attached file for your report"
$subject = "Daily Mail Counts for Staff"
$recipients = "" #Send report to
$mailserver = "" #Mail Server to use to send report
$from = "" #address report comes from

$Users = Get-Content $UserList
        foreach ($user in $Users){
            [Int] $intSent = $intRec = 0
Get-TransportServer -WarningAction SilentlyContinue | Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-1) -End (Get-Date) -Sender "$user" -EventID RECEIVE | ? {$_.Source -eq "STOREDRIVER"} | ForEach { $intSent++ }
Get-TransportServer -WarningAction SilentlyContinue| Get-MessageTrackingLog -ResultSize Unlimited -Start (Get-Date).AddDays(-1) -End (Get-Date) -Recipients "$User" -EventID DELIVER | ForEach { $intRec++ }
$userMailsSent = "Sent:",$intSent
$userMailsRecieved=  "Recieved", $intRec


$user +"  " +"Recieved:" + $IntRec |Out-File -FilePath $temp -Append 
$user + " "+ "Sent:" + $IntSent |Out-File -FilePath $temp -Append 

 }

Send-MailMessage -Attachments $temp -Body $body -Subject $subject -To $recipients -From $from -SmtpServer $mailserver
del $temp -Force
