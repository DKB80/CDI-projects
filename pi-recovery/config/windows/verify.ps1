<#
.SYNOPSIS
  Read-only compliance check for the CDI Windows standard. Changes NOTHING.
  Prints PASS/FAIL per item and exits non-zero if anything fails, so it can
  gate a rollout. Run after configure.ps1 (or on its own to audit a site).
    powershell -ExecutionPolicy Bypass -File .\verify.ps1
#>
param([string]$ReportDir = $PSScriptRoot)

$Expect = @{                                   # mirror configure.ps1's DESIRED STATE
  TimeZone = 'W. Australia Standard Time'
  EnableRDP = $true
  FirewallProfilesOn = $true
  Packages = @('Tailscale.Tailscale')
  Services = @{ 'W32Time' = 'Automatic' }
}

$fail = 0
function Check($name, [bool]$ok, $detail='') {
  $tag = if ($ok) { 'PASS' } else { $script:fail++; 'FAIL' }
  "{0,-5} {1} {2}" -f $tag, $name, $detail
}

$out = @("CDI compliance - $env:COMPUTERNAME - $(Get-Date -Format s)", ('='*50))
$out += Check 'timezone'  ((Get-TimeZone).Id -eq $Expect.TimeZone) (Get-TimeZone).Id
$out += Check 'rdp'       (((Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections).fDenyTSConnections) -eq 0)
$fw = (Get-NetFirewallProfile | Where-Object Enabled -eq $true).Count
$out += Check 'firewall'  ($fw -eq 3) "$fw/3 profiles on"
foreach ($p in $Expect.Packages) {
  $have = [bool](winget list --id $p -e --accept-source-agreements 2>$null | Select-String $p)
  $out += Check "pkg:$p" $have
}
foreach ($s in $Expect.Services.Keys) {
  $svc = Get-Service $s -ErrorAction SilentlyContinue
  $out += Check "svc:$s" ($svc -and $svc.StartType -eq $Expect.Services[$s]) ($svc.StartType)
}

$rep = Join-Path $ReportDir ("cdi-verify_{0}_{1}.txt" -f $env:COMPUTERNAME, (Get-Date -Format 'yyyyMMdd-HHmmss'))
$out += ('='*50); $out += (if ($fail) { "RESULT: $fail FAILED" } else { 'RESULT: all PASS' })
$out | Tee-Object -FilePath $rep | Out-Host
exit ([int]($fail -gt 0))
