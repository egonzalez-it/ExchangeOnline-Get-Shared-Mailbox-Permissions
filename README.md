# 📧 Get Shared Mailbox Access (Exchange Online)

![PowerShell](https://img.shields.io/badge/PowerShell-Ready-blue)
![Exchange Online](https://img.shields.io/badge/Microsoft-Exchange%20Online-red)
![Microsoft 365](https://img.shields.io/badge/Microsoft-365-blue)
![License](https://img.shields.io/badge/License-MIT-yellow)

---

## 🚀 Overview

This PowerShell script identifies which **shared mailboxes** a user can access in **Exchange Online**.

It checks for the following permissions:

- Full Access  
- Send As  
- Send on Behalf

Designed for IT administrators who need a quick and reliable way to audit user access to shared mailboxes.

---

## 🎯 Key Features

- Checks all shared mailboxes in the tenant  
- Identifies **Full Access** permissions  
- Identifies **Send As** permissions  
- Identifies **Send on Behalf** permissions  
- Clean and structured output  
- Simple and reusable script  

---

## ⚙️ Requirements

- Exchange Online PowerShell module
- Active Exchange Online session
- Appropriate permissions to read mailbox delegation settings 

---

## 🔐 Connect to Exchange Online

Before running the script:

```powershell
Connect-ExchangeOnline
```
---

## Usage

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\Get-SharedMailboxAccess.ps1 -UserEmail user@company.com
```
---

## Example Output
```
SharedMailbox               FullAccess SendAs SendOnBehalf
-------------               ---------- ------ ------------
finance@company.com         True       False  True
support@company.com         True       True   False
info@company.com            False      True   False
```

---

## Notes
This script is read-only
It does not modify any Exchange configuration
Results depend on your current Exchange Online administrative permissions
In some environments, Send on Behalf values may be stored as recipient names or aliases rather than SMTP addresses