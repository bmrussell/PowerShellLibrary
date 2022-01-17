Param([string]$To, [string]$From, [string]$Subject, [string]$Body, [string]$SmtpServer)

if ($Subject -eq "") {
	$Subject = "Test email sent from $($env:computername)"
}

if ($From -eq "") {
	$From = "$($env:computername)@$($env:userdnsdomain)"
}

if ($Body -eq "") {
	$Bsg = $Subject
}

Send-MailMessage -From $From -To ($To) -Subject ($Subject) -Body ($msg) -Priority Normal -SmtpServer $SmtpServer

# SSL: -Port 465 -UseSsl
