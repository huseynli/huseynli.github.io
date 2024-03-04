---
title: PowerShell - Dumping emails from an Outlook client
date: 2024-03-04 01:11:00 +0200
categories: [PowerShell]
tags: [powershell, outlook, registry]
img_path: /assets/img/posts/
image:
  path: warning-programmatic-access.png
  alt: PowerShell - Dumping emails from an Outlook client.
---

Say you can only access a machine via terminal and you need to dump Outlook emails via a script. It is possible, but there are some caveats.

## Dumping emails
First and foremost, it is possible to export emails via a PowerShell script like below: 
```powershell
$Outlook = New-Object -ComObject Outlook.Application
$NameSpace = $Outlook.GetNamespace('mapi')
$Mailbox = $NameSpace.Stores['<email_address>'].GetRootFolder()
$Inbox = $Mailbox.Folders['Inbox']
$Inbox.items | Select-Object -Property Subject, Body
```
You can customize the script to suite your needs. For example, specify the Outlook folder or choose the specific properties you want to dump.

## Security measures
Outlook, rightfully, has security measures in place to block such activities.

![Programmatic access warning](warning-programmatic-access.png)
_Screenshot from a [Rangeforce](https://www.rangeforce.com) module_

It expects a user to allow or deny this activity as it might be malicious. This is not a big deal if you have a GUI access.

## Bypassing security measures
This warning message becomes a problem if you only have a terminal access to the system, or you need to run your script unattended.

> This is NOT a good idea for real-world applications. It bypasses security measures put in place for a good reason. This is was ad-hoc solution for a specific task that has no real-world risks or implications.
{: .prompt-danger }

It is possible to allow programmatic access to Outlook via following registry keys:
```terminal
reg add HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\16.0\Outlook\Security /t REG_DWORD /v ObjectModelGuard /d 2
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\office\16.0\Outlook\Security /t REG_DWORD /v PromptOOMSend /d 2
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\office\16.0\Outlook\Security /t REG_DWORD /v AdminSecurityMode /d 3
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\office\16.0\Outlook\Security /t REG_DWORD /v promptoomaddressinformationaccess /d 2
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\office\16.0\Outlook\Security /t REG_DWORD /v promptoomaddressbookaccess /d 2
```
> Replace 16.0 with your Office version.
{: .prompt-info }

> Some of the above are **HKCU** keys. You need to add them with the same user it is intended to run as.
{: .prompt-info }

Sources:
* "Allow Programmatic Access to Outlook" by nice-automation.
