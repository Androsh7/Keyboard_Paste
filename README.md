# Keyboard_Paste

Keyboard Paste is a utility that simulates keyboard input to paste clipboard contents as if they were typed manually.

This is especially useful for pasting text into applications or websites that block traditional clipboard pasting.

The macro can be customized to use any key combination for activation (default: Left Ctrl + Left Alt).

# Run on startup

To have this program run on startup open a shell in the folder containing the script and run the following command:

```
New-ItemProperty -Path "HKCU:\\Software\\Microsoft\\Windows\\CurrentVersion\\Run" -Name "Keyboard_Paste" -PropertyType String -Value "powershell.exe -executionpolicy bypass -file `"$((get-item -path "Keyboard_Paste.ps1").FullName)`""
```
