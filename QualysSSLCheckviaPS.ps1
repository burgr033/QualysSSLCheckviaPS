$version = "0.3"
$pathApp = "$(Split-Path -Parent $MyInvocation.MyCommand.Path)"
$fileApp = "$([System.IO.Path]::GetFileNameWithoutExtension($MyInvocation.MyCommand.Path ))"
$fileLog = $fileApp + ".log"

#List of Hosts you want to check
$hosts = "github.com"

$date = get-date -Format "dd.MM.yyyy"
$report = "$pathApp\$($fileApp)-report_$date.csv"

#csv delimiter
$delimiter = ";"

#E-Mail Details
$emailFrom = 'service@example.com'
$emailTo = 'admin@example.com'
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
"GRADE$($delimiter)IPADDRESS$($delimiter)HOSTNAME$($delimiter)PROGRESS$($delimiter)LINK$($delimiter)WARNINGS?" | Add-Content $report

Write-Log "start TLS check..."
foreach($hostname in $hosts){
    $requestURI = "https://api.ssllabs.com/api/v2/analyze?host=$hostname&publish=off&startNew=on"
    #$requestURI = "https://api.ssllabs.com/api/v2/analyze?host=$hostname&publish=off"
    Write-Log "tls check: $($hostname), Request URL $($requestURI)" 

    
    $webRequest = Invoke-RestMethod  $requestURI
    
    Foreach($endpoint in $($webRequest.endpoints)){
    
    }
    while($webRequest.status -ne "READY"){
        if ( $webRequest -eq $null ) {
            Write-Log "WebRequest is Null"
        } else {
            Foreach($endpoint in $($webRequest.endpoints)){
                Write-Log "found endpoint: $($endpoint.serverName) IP: [$($endpoint.ipAddress)] Progress: $($endpoint.progress)"
            }
            Write-Log "wait for status equal '[READY]' current: [$($webRequest.status)]" 
        }
        sleep 15

        $webRequest = Invoke-RestMethod  $requestURI
    }
    Write-Log "tls report is available"     
    foreach ($item in $webRequest.endpoints){
         $warnings = "Keine Warnungen"
        if($item.haswarnings -eq $true){
            $warnings = "Warnungen gefunden!"
        }
        $newline = "$($item.grade)$($delimiter)$($item.ipAddress)$($delimiter)$($hostname)$($delimiter)$($item.progress)%$($delimiter)https://www.ssllabs.com/ssltest/analyze.html?d=$hostname$($delimiter)$warnings"
        $newline | Add-Content $report
    }

}
Write-Log "send mail"
$message = New-Object System.Net.Mail.MailMessage ($emailfrom, $emailTo)
$message.Subject = $subject
$message.IsBodyHTML = $true
$cont = Import-CSV $report -Delimiter "$delimiter" | Convertto-html -As Table -Head $style
$message.Body = $cont
$smtp=new-object Net.Mail.SmtpClient($smtpServer)
try{
    Write-Log "sending mail"
    $smtp.Send($message)
}
catch{
    Write-Log "mail could not be sent; generating html report"
    Set-Content -Value $cont -Path "$($report).html"
}
finally{
    Remove-Item $report
    Write-Log "csv report removed"
}
