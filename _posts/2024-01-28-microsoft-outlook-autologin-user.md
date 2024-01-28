---
title: Microsoft Outlook - Log AD User in Automatically on Launch
date: 2024-01-28 22:39:00 +0200
categories: [Microsoft Office, Outlook]
tags: [microsoft, windows, outlook, registry, auto-login, office 2016]
img_path: /assets/img/posts/
---

Lets assume a situation where you want Outlook 2016 to login as the **currently logged in AD user** whenever Outlook is launched, including the first launch. Normally you would login to your email account after Outlook is installed and it should automatically login on consecutive launches. However, lets assume you are creating an environment for educational purposes where the virtual machine setup and outlook installation and deployment is scripted. You do not have the luxury of setting things up via gui and then distributing the snapshot of the environment.

> This method is NOT advised for real-world applications. This was an ad-hoc solution for a specific task that had no real-world risks or implications. No security or best-practice considerations were taken into account while doing this.
{: .prompt-warning }

The following registry keys should configure Outlook 2016 to log-in as the currently logged in AD user. Note that this was tested in an environment that also included Domain Controller with AD and an Exchange server. All on-premise installations.

```terminal
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ExcludeExplicitO365Endpoint /d 1
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ZeroConfigExchangeOnce /d 1
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ZeroConfigExchange /d 1
```

Depending on your Office 16 version, you might need another path. Try and see which one works (the path below includes `\Policies\`).
```terminal
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ExcludeExplicitO365Endpoint /d 1
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ZeroConfigExchangeOnce /d 1
reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Outlook\AutoDiscover /t REG_DWORD /v ZeroConfigExchange /d 1
```

`ExcludeExplicitO365Endpoint` should stop Outlook from considering/attempting to connect to O365 and instead use the on-premise Exchange server stated earlier.

Note that you are adding `HKCU` registry keys. These commands should be run under the same user you expect to auto-login to Outlook.

As stated in the warning above, this is not intended to be used in real-world applications. There are definitely better and more appropriate ways of achieving similar results.