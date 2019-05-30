Param(
    [string[]]$Calendars,
    [string]$ApplicationId,
    [string]$ApplicationSecret
)

write-output "$(Get-Date) Script started"

# Setup variables
$startVMs = @()
$stopVMs = @()

# Ensures you do not inherit an AzureRMContext in your runbook
Disable-AzureRmContextAutosave –Scope Process

# Set up an AzureRMContext using the Run As account
$connection = Get-AutomationConnection -Name AzureRunAsConnection
$null = Connect-AzureRmAccount -ServicePrincipal -Tenant $connection.TenantID -ApplicationID $connection.ApplicationID -CertificateThumbprint $connection.CertificateThumbprint
$null = Select-AzureRmSubscription -SubscriptionId $connection.SubscriptionID

# grab all VMs (where Tag "AutomatedStopStart" exists)
$azureVMs = Get-AzureRmVM -Status | Where-Object {$_.Tags.ContainsKey("AutomatedStopStart")}

# Set up the Graph connection using the registered application identity
$uri = "https://login.microsoftonline.com/$($connection.TenantID)/oauth2/v2.0/token"
$body = @{
    client_id     = $ApplicationId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $ApplicationSecret
    grant_type    = "client_credentials"
}

# Obtain an OAuth 2.0 Access Token
$tokenRequest = Invoke-WebRequest -Method Post -Uri $uri -ContentType "application/x-www-form-urlencoded" -Body $body -UseBasicParsing
$token = ($tokenRequest.Content | ConvertFrom-Json).access_token

# process all Calendars
foreach($calendar in $Calendars) {

    # scope VMs to process to this Calendar
    $calendarVMs = $azureVMs | Where-Object {$_.Tags["AutomatedStopStart"] -like $Calendar}

    # only work if there are VMs for this Calendar
    if($calendarVMs) {
        # Specify Graph query URI to call
        $uri = "https://graph.microsoft.com/v1.0/users/$($calendar)/calendar/calendarView?StartDateTime=$([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss'))&EndDateTime=$([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ss'))"

        # Run Graph API query 
        $query = Invoke-RestMethod -Method Get -Uri $uri -ErrorAction Stop -Headers @{
            Authorization      = "Bearer $token"
            'Content-Type'     = "application/json"
        }

        # Create an array of VMs based on the subjects in the returned Calendar events 
        $processVMs = @()
        foreach($calendarItem in $query.Value) {
            $processVMs += "$($calendarItem.Subject.ToLower())"
            write-output "  -- $($calendarItem.Subject) should be running ($(Get-Date ($calendarItem.start.dateTime)) -> $(Get-Date ($calendarItem.end.dateTime)))"
        }
    
        # arrays for VMs that need to be turned on and of
        $startVMs += $calendarVMs | Where-Object {$_.Powerstate -like "VM deallocated" -and $processVMs.Contains("$($_.Name.ToLower())")}
        $stopVMs += $calendarVMs | Where-Object {$_.Powerstate -like "VM running" -and -not $processVMs.Contains("$($_.Name.ToLower())")}
    }
}

# process arrays of VMs
$startVMs | ForEach-Object {
    write-output "Starting:         $($_.Name)"
    $null = $_ | Start-AzureRmVM -AsJob -ErrorAction SilentlyContinue
}
$stopVMs | ForEach-Object {
    write-output "Deallocating:     $($_.Name)"
    $null = $_ | Stop-AzureRmVM -AsJob -Force -ErrorAction SilentlyContinue
}

write-output "$(Get-Date) Script ended"