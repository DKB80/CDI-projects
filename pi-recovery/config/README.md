# CDI Site Configuration Toolkit

Standardise machine config across sites **uniformly and verifiably** — the job a
rubber-ducky can't do (no proof, no idempotency, breaks on state/timing). Two paths:

```
config/
  windows/  configure.ps1     standalone: apply the CDI standard, logged + reported
            verify.ps1        read-only compliance check (PASS/FAIL, exit code)
            bootstrap-winrm.ps1  one-time: enable WinRM so the Pi can manage it over the net
  mac/      configure.sh      standalone Mac mini config
  ansible/  site.yml          networked: push config to all hosts + verify
            inventory.example.ini, ansible.cfg
```

## Which path
- **Networked site** → Ansible from the Pi. One command configures every machine and
  reports what changed / what's compliant. Windows needs `bootstrap-winrm.ps1` run once
  first (from USB); Mac/Linux over SSH.
- **Standalone machine** (no network) → run `configure.ps1` (Windows) or `configure.sh`
  (Mac) from the Pi's USB, elevated. Then `verify.ps1` to prove it. All runs write a
  timestamped report next to the script.

## Define your standard once
Edit the **DESIRED STATE** block at the top of `configure.ps1` (and mirror it in
`verify.ps1` and `ansible/site.yml`): timezone/NTP, power, RDP, firewall, winget
packages, services, scheduled tasks, static IP. Then it applies the same everywhere.

## Networked workflow
```bash
# on the Pi (Ansible is installed at /usr/bin/ansible)
cd /opt/cdi-recovery/config/ansible
ansible-galaxy collection install ansible.windows community.windows community.general
cp inventory.example.ini inventory.ini      # fill in hosts; secrets via ansible-vault
ansible-playbook site.yml --check --diff    # dry run — shows changes, changes nothing
ansible-playbook site.yml                   # apply
ansible-playbook site.yml --tags verify --check   # compliance pass
```

## Standalone workflow (Windows)
1. Arm the Pi's USB drive (`cdi-gadget-massstorage.sh on`) or copy `config/windows/` to a USB stick.
2. On the target, elevated PowerShell:
   ```
   powershell -ExecutionPolicy Bypass -File .\configure.ps1 -WhatIf   # preview
   powershell -ExecutionPolicy Bypass -File .\configure.ps1           # apply
   powershell -ExecutionPolicy Bypass -File .\verify.ps1              # prove it
   ```
3. Collect the `*.report.txt` files as your per-site compliance record.

## Why this and not a rubber ducky
Keystroke injection types blindly into whatever's on screen — different machine
states make the "same" script land differently, it needs an already-unlocked session,
and it returns no proof. These scripts are transparent, idempotent, run deliberately by
the operator on machines you own/service, and **produce a compliance report** — which is
exactly "ensure they're all configured correctly."

> Scope: only run against equipment you own or are contracted to service.
