<#
.SYNOPSIS
  One-time bootstrap so the Pi (Ansible) can manage this Windows box over the
  network. Enables WinRM over HTTPS with a self-signed cert and opens the
  firewall for it. Run once per machine, elevated, from the USB.
    powershell -ExecutionPolicy Bypass -File .\bootstrap-winrm.ps1

  After this, from the Pi:  ansible-playbook -i inventory.ini site.yml
  Only run on machines you own/service and intend to manage remotely.
#>
$ErrorActionPreference = 'Stop'
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
    ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { Write-Error 'Run as Administrator.'; exit 1 }

Write-Host 'Enabling WinRM (HTTPS)...' -ForegroundColor Cyan
Enable-PSRemoting -Force -SkipNetworkProfileCheck | Out-Null

# HTTPS listener with a self-signed cert (Ansible connects with validate_certs=no on a trusted LAN)
$cn = $env:COMPUTERNAME
if (-not (Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$cn" })) {
  $cert = New-SelfSignedCertificate -DnsName $cn -CertStoreLocation Cert:\LocalMachine\My
} else { $cert = Get-ChildItem Cert:\LocalMachine\My | Where-Object { $_.Subject -eq "CN=$cn" } | Select-Object -First 1 }

if (-not (Get-ChildItem WSMan:\localhost\Listener | Where-Object { $_.Keys -match 'Transport=HTTPS' })) {
  New-Item -Path WSMan:\localhost\Listener -Transport HTTPS -Address * -CertificateThumbPrint $cert.Thumbprint -Force | Out-Null
}
Set-Item WSMan:\localhost\Service\Auth\Basic $true
Set-Item WSMan:\localhost\Service\AllowUnencrypted $false

New-NetFirewallRule -DisplayName 'WinRM HTTPS (CDI)' -Direction Inbound -Protocol TCP -LocalPort 5986 -Action Allow -ErrorAction SilentlyContinue | Out-Null

Write-Host "WinRM HTTPS ready on $cn (port 5986). Add this host to the Pi's Ansible inventory." -ForegroundColor Green
