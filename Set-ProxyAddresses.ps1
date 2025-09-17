# Set-ProxyAddresses.ps1
# Script to set primary and secondary proxy addresses for AD users from CSV

param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath
)

# Import Active Directory module
Import-Module ActiveDirectory -ErrorAction Stop

# Check if CSV file exists
if (!(Test-Path $CsvPath)) {
    Write-Error "CSV file not found: $CsvPath"
    exit 1
}

# Import CSV file
try {
    $users = Import-Csv -Path $CsvPath
    Write-Host "Successfully imported CSV with $($users.Count) users" -ForegroundColor Green
}
catch {
    Write-Error "Failed to import CSV: $_"
    exit 1
}

# Validate CSV headers
$requiredHeaders = @('SAMaccount', 'UPN', 'ProxyAddress1', 'ProxyAddress2')
$csvHeaders = $users[0].PSObject.Properties.Name

foreach ($header in $requiredHeaders) {
    if ($header -notin $csvHeaders) {
        Write-Error "Missing required column: $header"
        exit 1
    }
}

Write-Host "`n=== Proxy Address Configuration Script ===" -ForegroundColor Cyan
Write-Host "This script will set proxy addresses for users based on the CSV data."
Write-Host "Primary address (ProxyAddress1) will be set as SMTP: (uppercase)"
Write-Host "Secondary address (ProxyAddress2) will be set as smtp: (lowercase)"
Write-Host "`nYou will be prompted to approve each change before it's committed.`n" -ForegroundColor Yellow

$successCount = 0
$errorCount = 0
$skippedCount = 0

foreach ($user in $users) {
    # Clean up any whitespace
    $samAccount = $user.SAMaccount.Trim()
    $proxyAddress1 = $user.ProxyAddress1.Trim()
    $proxyAddress2 = $user.ProxyAddress2.Trim()
    
    Write-Host "Processing user: $samAccount" -ForegroundColor White
    
    # Validate email addresses
    if ([string]::IsNullOrWhiteSpace($proxyAddress1) -or [string]::IsNullOrWhiteSpace($proxyAddress2)) {
        Write-Warning "Skipping $samAccount - Missing proxy address data"
        $skippedCount++
        continue
    }
    
    # Check if user exists in AD
    try {
        $adUser = Get-ADUser -Identity $samAccount -Properties ProxyAddresses -ErrorAction Stop
    }
    catch {
        Write-Error "User not found in AD: $samAccount"
        $errorCount++
        continue
    }
    
    # Display current proxy addresses
    Write-Host "  Current ProxyAddresses: " -NoNewline
    if ($adUser.ProxyAddresses) {
        Write-Host ($adUser.ProxyAddresses -join ', ') -ForegroundColor Gray
    } else {
        Write-Host "None" -ForegroundColor Gray
    }
    
    # Build new proxy addresses array
    $newProxyAddresses = @(
        "SMTP:$proxyAddress1",  # Primary (uppercase SMTP)
        "smtp:$proxyAddress2"   # Secondary (lowercase smtp)
    )
    
    Write-Host "  Proposed ProxyAddresses: " -NoNewline
    Write-Host ($newProxyAddresses -join ', ') -ForegroundColor Green
    
    # Prompt for confirmation
    do {
        $response = Read-Host "  Apply this change? (Y)es, (N)o, (A)ll remaining, (Q)uit"
        $response = $response.ToUpper()
    } while ($response -notin @('Y', 'N', 'A', 'Q'))
    
    switch ($response) {
        'Q' {
            Write-Host "Script terminated by user." -ForegroundColor Yellow
            break
        }
        'N' {
            Write-Host "  Skipped by user." -ForegroundColor Yellow
            $skippedCount++
            continue
        }
        'A' {
            Write-Host "  Applying change and all remaining..." -ForegroundColor Green
            $approveAll = $true
        }
        'Y' {
            Write-Host "  Applying change..." -ForegroundColor Green
        }
    }
    
    # Apply the change
    try {
        Set-ADUser -Identity $samAccount -Replace @{ProxyAddresses = $newProxyAddresses}
        Write-Host "  SUCCESS: Proxy addresses updated for $samAccount" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Error "  FAILED: Could not update $samAccount - $_"
        $errorCount++
    }
    
    Write-Host "" # Empty line for readability
    
    # If user chose 'Q', break out of the loop
    if ($response -eq 'Q') {
        break
    }
    
    # If 'A' was selected, continue without prompting
    if ($response -eq 'A') {
        # Continue the loop but skip prompting for remaining users
        foreach ($remainingUser in $users[($users.IndexOf($user) + 1)..($users.Count - 1)]) {
            $samAccount = $remainingUser.SAMaccount.Trim()
            $proxyAddress1 = $remainingUser.ProxyAddress1.Trim()
            $proxyAddress2 = $remainingUser.ProxyAddress2.Trim()
            
            Write-Host "Processing user: $samAccount" -ForegroundColor White
            
            if ([string]::IsNullOrWhiteSpace($proxyAddress1) -or [string]::IsNullOrWhiteSpace($proxyAddress2)) {
                Write-Warning "Skipping $samAccount - Missing proxy address data"
                $skippedCount++
                continue
            }
            
            try {
                $adUser = Get-ADUser -Identity $samAccount -Properties ProxyAddresses -ErrorAction Stop
            }
            catch {
                Write-Error "User not found in AD: $samAccount"
                $errorCount++
                continue
            }
            
            $newProxyAddresses = @(
                "SMTP:$proxyAddress1",
                "smtp:$proxyAddress2"
            )
            
            Write-Host "  Applying: " -NoNewline
            Write-Host ($newProxyAddresses -join ', ') -ForegroundColor Green
            
            try {
                Set-ADUser -Identity $samAccount -Replace @{ProxyAddresses = $newProxyAddresses}
                Write-Host "  SUCCESS: Proxy addresses updated for $samAccount" -ForegroundColor Green
                $successCount++
            }
            catch {
                Write-Error "  FAILED: Could not update $samAccount - $_"
                $errorCount++
            }
            
            Write-Host ""
        }
        break
    }
}

# Summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Successful updates: $successCount" -ForegroundColor Green
Write-Host "Errors: $errorCount" -ForegroundColor Red
Write-Host "Skipped: $skippedCount" -ForegroundColor Yellow
Write-Host "Total processed: $($successCount + $errorCount + $skippedCount)" -ForegroundColor White