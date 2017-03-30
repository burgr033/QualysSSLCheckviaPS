$version = "0.2"
$pathApp = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)"
$fileApp = "$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path ))"
$fileLog = $fileApp + ".log"

#List of Hosts you want to check
$hosts = "www.github.com"

$date = get-date -Format "dd.MM.yyyy"
$report = "$pathApp\$($fileApp)-report_$date.csv"

#E-Mail Details
$emailFrom = ''
$emailTo = ''
$subject="TLS Certificate Checks $date"
$smtpserver=''
$style = "<style>BODY{font-family: Arial; font-size: 10pt;}"
$style = $style + "TABLE{border: 1px solid black; border-collapse: collapse;}"
$style = $style + "TH{border: 1px solid black; background: #dddddd; padding: 5px; }"
$style = $style + "TD{border: 1px solid black; padding: 5px; }"
$style = $style + "</style>"


#  *********************************************************************
function Write-Log {
	param (
		[String] $Message
	)
	("{0} [{1}]{2} {3}" -f (Get-Date -Format "yyyy-MM-dd--HH:mm:ss"),([System.Diagnostics.Process]::GetCurrentProcess().id),   $Message, $_ ) | Out-File -FilePath "$($pathApp)\\$($fileLog)" -append -encoding ascii
}
#  *********************************************************************
# remove log from last run
if ( Test-path "$($pathApp)\\$($fileLog)" ) {
    Remove-Item -Path "$($pathApp)\\$($fileLog)"
}
#Log initial details
$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
$Whoami = whoami # Simple, could use $env as well
Write-Log "Running script $($MyInvocation.MyCommand.Path) at $Date" 
Write-Log "Admin: $Admin" 
Write-Log "User: $Whoami" 
Write-Log "Bound parameters: $($PSBoundParameters | Out-String)" 


#Create new Item
Write-Log "create output file"
New-Item $report -ItemType File -ErrorAction SilentlyContinue
# add content to file
"Grad erreicht;IP-Adresse;Hostname;Progress;Report-Link;Warnungen vorhanden?" | Add-Content $report

Write-Log "start tls check..."
foreach($hostname in $hosts){
    $requestURI = "https://api.ssllabs.com/api/v2/analyze?host=$hostname&publish=off&startnew=on"
    Write-Log "tls check: $($hostname), Request URL $($requestURI) ------------------" 

    
    $webRequest = Invoke-RestMethod  $requestURI
    Write-Log "start web request :" . $webRequest|out-string
    while($webRequest.status -ne "READY"){
        if ( $webRequest -eq $null ) {
            Write-Log "WebRequest is Null"
        } else {        
            Write-Log "wait for status equal 'ready' [$($webRequest.status)]" 
        }
        Write-Log "Response $($webRequest.endpoints)"
        sleep 15

        $webRequest = Invoke-RestMethod  $requestURI
        Write-Log "start web request :" . $webRequest|out-string
    }
    Write-Log "tls report is available"     
    foreach ($item in $webRequest.endpoints){
         $warnings = "Keine Warnungen"
        if($item.haswarnings -eq $true){
            $warnings = "Warnungen gefunden!"
        }
        $newline = "$($item.grade);$($item.ipAddress);$($hostname);$($item.progress)%;https://www.ssllabs.com/ssltest/analyze.html?d=$hostname;$warnings"
        $newline | Add-Content $report
    }

}
Write-Log "send mail"
$message = New-Object System.Net.Mail.MailMessage ($emailfrom, $emailTo)
$message.Subject = $subject
$message.IsBodyHTML = $true
$cont = Import-CSV $report -Delimiter ";" | Convertto-html -As Table -Head $style
$message.Body = $cont
$smtp=new-object Net.Mail.SmtpClient($smtpServer)
$smtp.Send($message)
Write-Log "mail sent"
Remove-Item $report
Write-Log "report removed"
