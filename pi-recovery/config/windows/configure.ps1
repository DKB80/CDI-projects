<#
.SYNOPSIS
  CDI standard-configuration applier for Windows (standalone / USB path).
  Idempotent: safe to re-run. Logs a transcript and writes a compliance report
  next to itself so you have PROOF each machine was configured.

.NOTES
  Run elevated:  right-click > Run with PowerShell (as Administrator), or:
    powershell -ExecutionPolicy Bypass -File .\configure.ps1
  Dry run (show what would change, change nothing):
    powershell -ExecutionPolicy Bypass -File .\configure.ps1 -WhatIf

  This is NOT keystroke injection. It is a transparent, reviewable script the
  operator runs deliberately on a machine they own/service. Edit the DESIRED
  STATE block for your standard, then it applies uniformly everywhere.
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
  [string]$ReportDir = $PSScriptRoot
)

# ===================== DESIRED STATE — EDIT FOR YOUR STANDARD =====================
$Desired = @{
  TimeZone      = 'W. Australia Standard Time'      # tzutil /l to list
  NtpServer     = 'au.pool.ntp.org'
  PowerNeverSleep = $true                            # never sleep/hibernate (kiosk/SCADA)
  EnableRDP     = $true
  FirewallOn    = $true
  # winget package IDs to ensure present (idempotent):
  Packages      = @('Tailscale.Tailscale')          # add e.g. 'Google.Chrome','VideoLAN.VLC'
  # services: name => desired StartupType ('Automatic','Manual','Disabled')
  Services      = @{ 'W32Time' = 'Automatic' }
  # scheduled tasks to ensure exist: name => @{ Exe; Args; Schedule }
  ScheduledTasks = @{
    # 'CDI-Poll' = @{ Exe='C:\CDI\poll.exe'; Args=''; AtStartup=$true }
  }
  # static IP (leave $null to skip — network changes are the riskiest, do carefully):
  StaticIP      = $null   # e.g. @{ Alias='Ethernet'; IP='192.168.1.50'; Prefix=24; Gateway='192.168.1.1'; DNS='192.168.1.1' }
}
# =================================================================================

$ErrorActionPreference = 'Stop'
$stamp   = Get-Date -Format 'yyyyMMdd-HHmmss'
$host_   = $env:COMPUTERNAME
$log     = Join-Path $ReportDir "cdi-config_${host_}_${stamp}.log"
$report  = Join-Path $ReportDir "cdi-config_${host_}_${stamp}.report.txt"
$results = [System.Collections.Generic.List[string]]::new()

function Note($area, $msg, $state) { $results.Add(("{0,-16} {1,-8} {2}" -f $area, $state, $msg)) }

# require admin
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
      ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
  Write-Error 'Must run as Administrator.'; exit 1
}

Start-Transcript -Path $log -Append | Out-Null
Write-Host "CDI configure.ps1 on $host_  ($stamp)  WhatIf=$($WhatIfPreference)" -ForegroundColor Cyan

try {
  # --- OS settings & policies ---
  if ($Desired.TimeZone) {
    if ((Get-TimeZone).Id -ne $Desired.TimeZone) {
      if ($PSCmdlet.ShouldProcess('TimeZone', "set $($Desired.TimeZone)")) { Set-TimeZone -Id $Desired.TimeZone }
      Note 'OS' "timezone -> $($Desired.TimeZone)" 'SET'
    } else { Note 'OS' 'timezone already correct' 'OK' }
  }
  if ($Desired.NtpServer) {
    if ($PSCmdlet.ShouldProcess('NTP', "set $($Desired.NtpServer)")) {
      w32tm /config /manualpeerlist:"$($Desired.NtpServer)" /syncfromflags:manual /update | Out-Null
      Start-Service W32Time -ErrorAction SilentlyContinue; w32tm /resync | Out-Null
    }
    Note 'OS' "ntp -> $($Desired.NtpServer)" 'SET'
  }
  if ($Desired.PowerNeverSleep) {
    if ($PSCmdlet.ShouldProcess('Power', 'never sleep')) {
      powercfg /change standby-timeout-ac 0; powercfg /change hibernate-timeout-ac 0
      powercfg /change monitor-timeout-ac 15
    }
    Note 'OS' 'power: never sleep' 'SET'
  }
  if ($Desired.EnableRDP) {
    $cur = (Get-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections).fDenyTSConnections
    if ($cur -ne 0) {
      if ($PSCmdlet.ShouldProcess('RDP', 'enable')) {
        Set-ItemProperty 'HKLM:\System\CurrentControlSet\Control\Terminal Server' -Name fDenyTSConnections -Value 0
        Enable-NetFirewallRule -DisplayGroup 'Remote Desktop' -ErrorAction SilentlyContinue
      }
      Note 'OS' 'RDP enabled' 'SET'
    } else { Note 'OS' 'RDP already enabled' 'OK' }
  }
  if ($Desired.FirewallOn) {
    if ($PSCmdlet.ShouldProcess('Firewall', 'enable all profiles')) {
      Set-NetFirewallProfile -Profile Domain,Public,Private -Enabled True
    }
    Note 'OS' 'firewall on' 'SET'
  }

  # --- Install / update software ---
  foreach ($pkg in $Desired.Packages) {
    $have = (winget list --id $pkg -e --accept-source-agreements 2>$null | Select-String $pkg)
    if (-not $have) {
      if ($PSCmdlet.ShouldProcess("winget:$pkg", 'install')) {
        winget install --id $pkg -e --silent --accept-package-agreements --accept-source-agreements | Out-Null
      }
      Note 'Software' "install $pkg" 'SET'
    } else { Note 'Software' "$pkg present" 'OK' }
  }

  # --- Services ---
  foreach ($svc in $Desired.Services.Keys) {
    $want = $Desired.Services[$svc]; $s = Get-Service $svc -ErrorAction SilentlyContinue
    if ($s -and $s.StartType -ne $want) {
      if ($PSCmdlet.ShouldProcess("svc:$svc", "startup $want")) { Set-Service $svc -StartupType $want }
      Note 'Service' "$svc -> $want" 'SET'
    } elseif ($s) { Note 'Service' "$svc already $want" 'OK' } else { Note 'Service' "$svc missing" 'WARN' }
  }

  # --- Scheduled tasks ---
  foreach ($tn in $Desired.ScheduledTasks.Keys) {
    $t = $Desired.ScheduledTasks[$tn]
    if (-not (Get-ScheduledTask -TaskName $tn -ErrorAction SilentlyContinue)) {
      if ($PSCmdlet.ShouldProcess("task:$tn", 'create')) {
        $act = New-ScheduledTaskAction -Execute $t.Exe -Argument ($t.Args)
        $trg = if ($t.AtStartup) { New-ScheduledTaskTrigger -AtStartup } else { New-ScheduledTaskTrigger -AtLogOn }
        Register-ScheduledTask -TaskName $tn -Action $act -Trigger $trg -RunLevel Highest -User 'SYSTEM' | Out-Null
      }
      Note 'Task' "create $tn" 'SET'
    } else { Note 'Task' "$tn exists" 'OK' }
  }

  # --- Network (guarded) ---
  if ($Desired.StaticIP) {
    $n = $Desired.StaticIP
    if ($PSCmdlet.ShouldProcess("net:$($n.Alias)", "static $($n.IP)/$($n.Prefix)")) {
      New-NetIPAddress -InterfaceAlias $n.Alias -IPAddress $n.IP -PrefixLength $n.Prefix -DefaultGateway $n.Gateway -ErrorAction SilentlyContinue
      Set-DnsClientServerAddress -InterfaceAlias $n.Alias -ServerAddresses $n.DNS
    }
    Note 'Network' "static $($n.IP) on $($n.Alias)" 'SET'
  }
}
catch { Note 'ERROR' $_.Exception.Message 'FAIL' }
finally {
  $summary = @("CDI configuration report - $host_ - $stamp", ('=' * 50)) + $results
  $summary | Tee-Object -FilePath $report | Out-Host
  Write-Host "`nReport: $report" -ForegroundColor Green
  Stop-Transcript | Out-Null
}
