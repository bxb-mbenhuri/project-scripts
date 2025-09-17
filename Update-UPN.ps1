# Update-UPN.ps1
# Script to update UPN for AD users from CSV

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
$requiredHeaders = @('SAMaccount', 'UPN', 'ProxyAddress1', 'ProxyAddress2', 'NewUPN')
$csvHeaders = $users[0].PSObject.Properties.Name

foreach ($header in $requiredHeaders) {
    if ($header -notin $csvHeaders) {
        Write-Error "Missing required column: $header"
        exit 1
    }
}

Write-Host "`n=== UPN Update Script ===" -ForegroundColor Cyan
Write-Host "This script will update UPN for users based on the CSV data."
Write-Host "Current UPN will be changed to NewUPN value from the CSV."
Write-Host "`nYou will be prompted to approve each change before it's committed.`n" -ForegroundColor Yellow

$successCount = 0
$errorCount = 0
$skippedCount = 0

foreach ($user in $users) {
    # Clean up any whitespace
    $samAccount = $user.SAMaccount.Trim()
    $currentUPN = $user.UPN.Trim()
    $newUPN = $user.NewUPN.Trim()
    
    Write-Host "Processing user: $samAccount" -ForegroundColor White
    
    # Validate UPN data
    if ([string]::IsNullOrWhiteSpace($newUPN)) {
        Write-Warning "Skipping $samAccount - Missing NewUPN data"
        $skippedCount++
        continue
    }
    
    # Check if user exists in AD
    try {
        $adUser = Get-ADUser -Identity $samAccount -Properties UserPrincipalName -ErrorAction Stop
    }
    catch {
        Write-Error "User not found in AD: $samAccount"
        $errorCount++
        continue
    }
    
    # Display current UPN
    Write-Host "  Current UPN in AD: " -NoNewline
    Write-Host $adUser.UserPrincipalName -ForegroundColor Gray
    
    # Check if current UPN matches CSV
    if ($adUser.UserPrincipalName -ne $currentUPN) {
        Write-Warning "  NOTE: Current UPN in AD ($($adUser.UserPrincipalName)) differs from CSV UPN ($currentUPN)"
    }
    
    # Check if UPN is already set to the new value
    if ($adUser.UserPrincipalName -eq $newUPN) {
        Write-Host "  UPN is already set to target value - skipping." -ForegroundColor Yellow
        $skippedCount++
        Write-Host ""
        continue
    }
    
    Write-Host "  Proposed new UPN: " -NoNewline
    Write-Host $newUPN -ForegroundColor Green
    
    # Prompt for confirmation
    do {
        $response = Read-Host "  Apply this UPN change? (Y)es, (N)o, (A)ll remaining, (Q)uit"
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
            Write-Host "  Applying UPN change..." -ForegroundColor Green
        }
    }
    
    # Apply the change
    try {
        Set-ADUser -Identity $samAccount -UserPrincipalName $newUPN
        Write-Host "  SUCCESS: UPN updated for $samAccount" -ForegroundColor Green
        Write-Host "    Old: $($adUser.UserPrincipalName)" -ForegroundColor Gray
        Write-Host "    New: $newUPN" -ForegroundColor Green
        $successCount++
    }
    catch {
        Write-Error "  FAILED: Could not update UPN for $samAccount - $_"
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
            $currentUPN = $remainingUser.UPN.Trim()
            $newUPN = $remainingUser.NewUPN.Trim()
            
            Write-Host "Processing user: $samAccount" -ForegroundColor White
            
            if ([string]::IsNullOrWhiteSpace($newUPN)) {
                Write-Warning "Skipping $samAccount - Missing NewUPN data"
                $skippedCount++
                continue
            }
            
            try {
                $adUser = Get-ADUser -Identity $samAccount -Properties UserPrincipalName -ErrorAction Stop
            }
            catch {
                Write-Error "User not found in AD: $samAccount"
                $errorCount++
                continue
            }
            
            # Skip if already set to target value
            if ($adUser.UserPrincipalName -eq $newUPN) {
                Write-Host "  UPN already set to target value - skipping." -ForegroundColor Yellow
                $skippedCount++
                Write-Host ""
                continue
            }
            
            Write-Host "  Changing UPN: " -NoNewline
            Write-Host "$($adUser.UserPrincipalName) -> $newUPN" -ForegroundColor Green
            
            try {
                Set-ADUser -Identity $samAccount -UserPrincipalName $newUPN
                Write-Host "  SUCCESS: UPN updated for $samAccount" -ForegroundColor Green
                $successCount++
            }
            catch {
                Write-Error "  FAILED: Could not update UPN for $samAccount - $_"
                $errorCount++
            }
            
            Write-Host ""
        }
        break
    }
}

# Summary
Write-Host "`n=== SUMMARY ===" -ForegroundColor Cyan
Write-Host "Successful UPN updates: $successCount" -ForegroundColor Green
Write-Host "Errors: $errorCount" -ForegroundColor Red
Write-Host "Skipped: $skippedCount" -ForegroundColor Yellow
Write-Host "Total processed: $($successCount + $errorCount + $skippedCount)" -ForegroundColor White

# Additional validation check
if ($successCount -gt 0) {
    Write-Host "`nRecommendation: Run a verification query to confirm changes:" -ForegroundColor Cyan
    Write-Host "Get-ADUser -Filter * -Properties UserPrincipalName | Where-Object {`$_.SamAccountName -in @('user1','user2')} | Select-Object SamAccountName,UserPrincipalName" -ForegroundColor Gray
}