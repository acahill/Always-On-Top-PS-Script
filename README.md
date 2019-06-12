Always-On-Top-PS-Script
=======================

A PowerShell script that lets you select an open window (or windows) to give the "always on top" attribute.

I use it with Remote Desktop, since I have three monitors but the program can only fullscreen across one monitor or all of them. With this script I can make another application stay on top of my fullscreen RDP connection on the third monitor, even though my organisation's security policy blocks third-party programs.

Also great if you are watching a video and want to make sure it isn't covered by other windows as you move them around.

### Usage

To make this script easier to run you can create a shortcut to run it directly, with no console window in the background. For example, if the script is on the desktop use the Target:

```
C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -windowstyle hidden -command "& '%USERPROFILE%\Desktop\AlwaysOnTop.ps1' "
```

Also set the option "Run" to "Minimized".

### Credit

Forked from the original script https://github.com/bkfarnsworth/Always-On-Top-PS-Script
