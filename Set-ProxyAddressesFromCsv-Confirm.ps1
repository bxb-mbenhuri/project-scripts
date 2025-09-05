<#  Set-ProxyAddressesFromCsv-Confirm.ps1
    CSV columns: userprincipalname, primaryEmail, secondaryEmail
    Example run
    .\Set-ProxyAddressesFromCsv-Confirm.ps1 -CsvPath .\addresses.csv
    Optional
    Add -AutoAccept to skip prompts and apply changes automatically
#>

param(
  [Parameter(Mandatory = $true)]
  [string]$CsvPath,
  [switch]$AutoAccept
)

Import-Module ActiveDirectory

function Show-List {
  param([string]$Title, [string[]]$Items)
  Write-Host "$Title"
  if ($Items -and $Items.Count -gt 0) {
    $Items | ForEach-Object { Write-Host "  - $_" }
  } else {
    Write-Host "  - none"
  }
}

Import-Csv -Path $CsvPath | ForEach-Object {
  $upn = ($_.userprincipalname ?? '').Trim()
  $p   = ($_.primaryEmail     ?? '').Trim()
  $s   = ($_.secondaryEmail   ?? '').Trim()

  if (-not $upn) {
    Write-Host "Skipped row with empty userprincipalname"
    return
  }

  $user = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties proxyAddresses, mail, sAMAccountName
  if (-not $user) {
    Write-Host "User not found $upn"
    return
  }

  # Build proposed proxyAddresses
  $primaryEntry   = if ($p) { "SMTP:$p" } else { $null }
  $secondaryEntry = if ($s) { "smtp:$s" } else { $null }

  $current = @($user.proxyAddresses) | Where-Object { $_ }            # existing
  $other   = $current | Where-Object { $_ -notmatch '^(?i)smtp:' }    # keep non SMTP types
  $proposed = @($primaryEntry, $secondaryEntry) + $other |
              Where-Object { $_ } | Select-Object -Unique

  # Quick equality check ignoring order and case
  $same = (@($current)  | Sort-Object -Unique) -join '|' `
       -ieq (@($proposed) | Sort-Object -Unique) -join '|'

  Write-Host ""
  Write-Host "User $($user.SamAccountName)  UPN $upn"
  Show-List -Title "Current proxy addresses"  -Items $current
  Show-List -Title "Proposed proxy addresses" -Items $proposed
  if ($p) { Write-Host "Mail attribute will be set to $p" }

  if ($same -and ($user.mail -ieq $p)) {
    Write-Host "No change required"
    return
  }

  $doApply = $false
  if ($AutoAccept) {
    $doApply = $true
  } else {
    $answer = Read-Host "Type Y to apply or N to skip"
    if ($answer -match '^(?i)Y$') { $doApply = $true }
  }

  if (-not $doApply) {
    Write-Host "Skipped $upn"
    return
  }

  try {
    Set-ADUser -Identity $user -Replace @{ proxyAddresses = $proposed }
    if ($p) { Set-ADUser -Identity $user -EmailAddress $p }
    Write-Host "Applied changes for $upn"
  }
  catch {
    Write-Host "Error applying changes for $upn  $($_.Exception.Message)"
  }
}
