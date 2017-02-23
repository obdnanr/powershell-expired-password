# Written by Brandon Kelley 2016

Import-Module ActiveDirectory

$From = "vpnadmin@company.com"
$Subject = "Your Laptop VPN account's password is going to expire"
$SMTPServer = "smtpserver"

Send-MailMessage -From $From -to admin@company.com -Subject "Expired PW script started" -Body "Expired PW script started" -SmtpServer $SMTPServer

$myfile = "C:\users\user\desktop\test.csv"

###### Mount the Entrust Database 
New-PSDrive -name entrust2 -PSProvider ActiveDirectory -Root "DC=domain,DC=Local" -Scope Global -Server ldapserver.company.com:389
Set-Location entrust2:

###### Get a list of users with email addresses and the password expired field in it and export to a csv 
Get-ADuser -Filter * -properties * | 
select-object -Property name,mail,samaccountname,UserPasswordExpiryDate | 
Export-CSV -path $myfile -notypeinformation -Encoding UTF8

###### Setup date stuff
$HowManyDaysAgo = "-7"
$pastDate = (Get-Date).AddDays($HowManyDaysAgo)

$cutoff = "1"
$futuredate = (Get-Date).AddDays($cutoff)
write-host $pastDate
write-host $futuredate

###### Open file and if password expiry date is less then or equal to 7 days, email user

$csv = Import-CSV -path $myfile -Header name,mail,samaccountname,UserPasswordExpiryDate 

ForEach ($line in $csv) {
    $UserPasswordExpiryDate = [datetime]::ParseExact($line.UserPasswordExpiryDate,"M/d/yyyy h:m:s tt",$null)
    $name = $line.name
    $mail = $line.mail
    $samaccountname = $line.samaccountname
    #$UserPasswordExpiryDate = $line.UserPasswordExpiryDate 
    $Body = "$name, your laptop VPN account's password for account $samaccountname, is expiring on $UserPasswordExpiryDate.  Please visit https://reset.company.com to change your password.  If you no longer need a VPN account and wish to stop receiving these emails please reply with DELETE MY ACCOUNT."
    If ($line.UserPasswordExpiryDate -ne $null -and $line.mail -ne $null) { 
    If ($UserPasswordExpiryDate -ge $pastDate -and $UserPasswordExpiryDate -le $futuredate) {Send-MailMessage -From $From -To $line.mail -Subject $Subject -Body $Body -SmtpServer $SMTPServer} 
    }
}
    
Remove-PSDrive entrust2

Send-MailMessage -From $From -to admin@company.com -Subject "Expired PW script has ended" -Body "Expired PW script has ended" -SmtpServer $SMTPServer
