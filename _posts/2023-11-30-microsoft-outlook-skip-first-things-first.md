---
title: Microsoft Outlook - Skipping the "First Things First" Pop-up
date: 2023-11-30 22:00:00 +0200
categories: [Microsoft Office, Outlook]
tags: [microsoft, windows, outlook, registry, pop-up, office 2016]
img_path: /assets/img/posts/
image:
  path: office_first_things_first.png
  alt: Microsoft Outlook - "First Things First" Pop-up.
---

During the first launch of **Microsoft Outlook**, after its installation, a **"First Things First"** window will pop-up. In my specific case (hands-on cybersecurity training module), I needed it gone. I needed it not to pop-up as the end user would be the first one to launch outlook and I needed it to look as if Outlook has been used before.

I needed to do it in a scripted manner. No GUI. This probably can be done in a better way, in a proper way, but in my case the following registry keys helped:

```terminal
reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\General /t REG_DWORD /v shownfirstrunoptin /d 1

reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Common\General /t REG_DWORD /v shownfirstrunoptin /d 1

reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\FirstRun /t REG_DWORD /v BootedRTM /d 1

reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\FirstRun /t REG_DWORD /v BootedRTM /d 1

reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Registration /t REG_DWORD /v AcceptAllEulas /d 1

reg add HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Registration /t REG_DWORD /v AcceptAllEulas /d 1
```

I had Office 2016. I am not sure which registry key or maybe combination of keys did the job, but adding these, solved the issue for me. Couple of notes:
* Note that the registry keys seem to be repeating. They are not. There's a difference in the paths `...\Software\Policies\Microsoft\...` vs `...\Software\Microsoft\...`. In one version of Office 2016 the ones with `Policies` did the job, in the other version it worked without. Try one of the paths, and figure out which one does it for you, or just add them all.
* These are `HKCU` registry keys. Meaning, the registry keys will be added for the user you are currently logged in as and will skip the pop-up for this user.