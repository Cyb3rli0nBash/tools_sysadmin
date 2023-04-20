$events = Get-EventLog -LogName System -InstanceId 5838, 5839, 5840, 5841, 4625, 4766, 4767, 5819 | Select-Object TimeGenerated, Source, InstanceId, Message
$events += Get-EventLog -LogName Security -InstanceId 5838, 5839, 5840, 5841, 4625, 4766, 4767, 5819| Select-Object TimeGenerated, Source, InstanceId, Message
$events = $events | Group-Object -Property Source, InstanceId, Message | ForEach-Object { $_.Group[0] }

$filteredEvents = @()
foreach ($event in $events) {
    if ($event.InstanceId -eq "4769" -and $event.Message | Select-String -Pattern "(?msi)Ticket Encryption Type: 0x17\b") {
        $filteredEvents += $event
    } elseif ($event.InstanceId -ne "4769") {
        $filteredEvents += $event
    }
}

$events = $filteredEvents | Select-Object TimeGenerated, Source, InstanceId, Message

if ($events.Count -gt 0) {
    $html = "<html><head><style>table { border-collapse: collapse; width: 100%; } th, td { text-align: left; padding: 8px; } th { background-color: #4CAF50; color: white; } tr:nth-child(even){background-color: #f2f2f2}</style></head><body><h2>Events with Instance IDs 5838, 5839, 5840, 5841, 4625, 4766, 4767, 5819, or 4769 (with Ticket Encryption Type 0x17) in System and Security logs on Domain Controller</h2><table><tr><th>TimeGenerated</th><th>Source</th><th>InstanceId</th><th>Message</th></tr>"
    foreach ($event in $events) {
        $html += "<tr><td>" + $event.TimeGenerated + "</td><td>" + $event.Source + "</td><td>" + $event.InstanceId + "</td><td>" + $event.Message + "</td></tr>"
    }
    $html += "</table></body></html>"
    $reportPath = Join-Path $PSScriptRoot "events_reportt.html"
    $html | Out-File -Encoding UTF8 -FilePath $reportPath
    Write-Host "Events report has been generated and saved to '$reportPath'." -ForegroundColor Green
} else {
    Write-Host "No events with Instance IDs 5838, 5839, 5840, 5841, 4625, 4766, 4767, 5819, or 4769 (with Ticket Encryption Type 0x17) were found in the System or Security logs on Domain Controller." -ForegroundColor Green
}
