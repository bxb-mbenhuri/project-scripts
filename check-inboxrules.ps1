# Check for Exchange Online Management module
if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
    Write-Host "Exchange Online Management module not found. Installing..." -ForegroundColor Yellow
    Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber
}

# Import the module
Import-Module ExchangeOnlineManagement

# Connect to Exchange Online using modern authentication
try {
    Write-Host "Connecting to Exchange Online..." -ForegroundColor Cyan
    Write-Host "A sign-in window will appear. Please enter your credentials and complete MFA if required." -ForegroundColor Yellow
    Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    Write-Host "Successfully connected to Exchange Online" -ForegroundColor Green
}
catch {
    Write-Host "Failed to connect to Exchange Online: $_" -ForegroundColor Red
    exit
}

# Main menu loop
do {
    # Get user input
    $userEmail = Read-Host "Enter the user's email address (or 'exit' to quit)"
    
    # Exit condition
    if ($userEmail -eq 'exit') {
        break
    }
    
    # Validate email format
    if ($userEmail -notmatch "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$") {
        Write-Host "Invalid email format. Please try again." -ForegroundColor Red
        continue
    }
    
    # Get inbox rules
    try {
        Write-Host "Retrieving inbox rules for $userEmail..." -ForegroundColor Cyan
        $rules = Get-InboxRule -Mailbox $userEmail -ErrorAction Stop
        
        if ($rules) {
            Write-Host "`nFound $($rules.Count) rules for $userEmail`n" -ForegroundColor Green
            
            # Create custom output object
            $output = $rules | Select-Object @{Name='Rule Name'; Expression={$_.Name}}, 
                                         @{Name='Enabled'; Expression={$_.Enabled}},
                                         @{Name='Priority'; Expression={$_.Priority}},
                                         @{Name='Description'; Expression={$_.Description -replace "`r`n", " "}}
            
            # Display results in a formatted table
            $output | Format-Table -AutoSize -Wrap | Out-String -Width 4096 | Write-Host
            
            # Option to export to CSV
            $export = Read-Host "`nExport results to CSV? (Y/N)"
            if ($export -eq 'Y' -or $export -eq 'y') {
                $filePath = "$env:USERPROFILE\Desktop\$($userEmail.Split('@')[0])_Rules_$(Get-Date -Format 'yyyyMMdd').csv"
                $output | Export-Csv -Path $filePath -NoTypeInformation
                Write-Host "Results exported to: $filePath" -ForegroundColor Green
            }
        }
        else {
            Write-Host "No inbox rules found for $userEmail" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "Error retrieving rules: $_" -ForegroundColor Red
    }
    
    Write-Host "`n----------------------------------------`n"
} while ($true)

# Disconnect from Exchange Online
Write-Host "Disconnecting from Exchange Online..." -ForegroundColor Cyan
Disconnect-ExchangeOnline -Confirm:$false
Write-Host "Session disconnected. Goodbye!" -ForegroundColor Green