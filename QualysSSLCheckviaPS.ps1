#List of Hosts you want to check
$hosts = "Host1,Host2,Host3"

#Path of Report
$report = "$ENV:USERPROFILE\Desktop\TLSCertificatesReport_$(Get-Date -format "MM-dd-yyyy").csv"

#Create new Item
New-Item $report -ItemType File


"GRADE;IPADDRESS;SERVERNAME;PROGRESS;LINK;WARNINGS?" | Add-Content $report

foreach($hostname in $hosts){

    $requestURI = "https://api.ssllabs.com/api/v2/analyze?host=$hostname&publish=off&startnew=on"

    $out = Invoke-WebRequest $requestURI | ConvertFrom-Json

    while($out.status -ne "READY"){

        sleep 15

        $out = Invoke-WebRequest $requestURI | ConvertFrom-Json

    }

    foreach ($item in $out.endpoints){
        $warnings = "No Warnings"
        if($item.haswarnings -eq $true){
            $warnings = "Warnings found"
        }
        $newline = "$($item.grade);'$($item.ipAddress);$($item.servername);$($item.progress)%;https://www.ssllabs.com/ssltest/analyze.html?d=$hostname;$warnings"
        $newline | Add-Content $report
    }

}
