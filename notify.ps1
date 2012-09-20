#Mail Message
$user = ""
$pass = ""

$from = "from@email"
$to = "to@email"
$subject "CA has outstanding certificate requests"

#SMTP Server
$server = "some.fqdn.here"
$port = 25
$timeout = 30000

function SendMail($msg){
  $creds = New-Object System.Net.NetworkCredential($user,$pass)

	$client = New-Object System.Net.Mail.SmtpClient $server, $port
	$client.EnableSsl = $true
	$client.Timeout = $timeout
	$client.UseDefaultCredentials = $false
	$client.Credentials = $creds
	$message = New-Object System.Net.Mail.MailMessage $from,$to,$subject,$msg
	$client.Send($message)
}

$ret = (certutil -pingadmin)
$alive = $false
$ret | foreach { if ($_.IndexOf("is alive") -ge 0) {$alive = $true;} }

if ($alive -eq $false){
	Write-Output "Cannot find CA, are you running this locally?"
}else{
	$raw = certutil -view -out "Request ID, Request Submission Date, Request Common Name, Requester Name, Request Email Address, Request Disposition" -Restrict "Request Disposition=9"
	$pending = ($raw | Select-String "Maximum Row Index: ").ToString().Replace("Maximum Row Index: ","")
	if ($pending -ne 0){
		$msg = "There are ($pending) pending certificate requests waiting to be processed"
		write-host $msg
		SendMail("$msg, `r`n`r`n`r`n" + ($raw | foreach {"$_`r`n"}))
	}else{
		write-host "There are no pending certificate requests"
	}
}