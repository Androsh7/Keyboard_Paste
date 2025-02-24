# title screen
Write-Host "-------------------- Keyboard Paste --------------------" -ForegroundColor Green

# this allows for keyboard emulation
Write-Host "Loading System.Windows.Forms Assembly" -ForegroundColor Yellow -nonewline
Add-Type -AssemblyName System.Windows.Forms
Write-Host " - Done" -ForegroundColor Yellow

# this is required to read keyboard asynchronously (see powershell keylogger)
Write-Host "Loading user32.dll Keypress API" -ForegroundColor Yellow -nonewline
$signature = @"
[DllImport("user32.dll", CharSet=CharSet.Auto, ExactSpelling=true)]
public static extern short GetAsyncKeyState(int virtualKeyCode);
"@
$API = Add-Type -MemberDefinition $signature -Name 'Keypress' -Namespace API -PassThru
Write-host " - Done" -ForegroundColor Yellow

# these are the keys that will be used to trigger the macro
$macro_keys_raw = @("LCTRL", "LALT")

# debug flag (This will show the clipboard in plaintext)
$debug = $false

# key tables
# https://docs.microsoft.com/en-us/windows/win32/inputdev/virtual-key-codes
$winkeylist = @{
    "`b" = 8; "`t" = 9; "`n" = 13; " " = 32;
    "0" = 48; "1" = 49; "2" = 50; "3" = 51;
    "4" = 52; "5" = 53; "6" = 54; "7" = 55;
    "8" = 56; "9" = 57; "a" = 65; "b" = 66;
    "c" = 67; "d" = 68; "e" = 69; "f" = 70;
    "g" = 71; "h" = 72; "i" = 73; "j" = 74;
    "k" = 75; "l" = 76; "m" = 77; "n" = 78;
    "o" = 79; "p" = 80; "q" = 81; "r" = 82;
    "s" = 83; "t" = 84; "u" = 85; "v" = 86;
    "w" = 87; "x" = 88; "y" = 89; "z" = 90;
    "LCTRL" = 162; "RCTRL" = 163; "LSHIFT" = 160;
    "RSHIFT" = 161; "LALT" = 164; "RALT" = 165;
    ";" = 186; "=" = 187; "," = 188; "-" = 189;
    "." = 190; "/" = 191; "`"" = 192; "[" = 219
    "`\" = 220; "]" = 221; "'" = 222
}

$altkeylist = @{
    ")" = 48; "!" = 49; "@" = 50; "#" = 51;
    "$" = 52; "%" = 53; "^" = 54; "&" = 55;
    "*" = 56; "(" = 57; "A" = 65; "B" = 66;
    "C" = 67; "D" = 68; "E" = 69; "F" = 70;
    "G" = 71; "H" = 72; "I" = 73; "J" = 74;
    "K" = 75; "L" = 76; "M" = 77; "N" = 78;
    "O" = 79; "P" = 80; "Q" = 81; "R" = 82;
    "S" = 83; "T" = 84; "U" = 85; "V" = 86;
    "W" = 87; "X" = 88; "Y" = 89; "Z" = 90;
    ":" = 186; "+" = 187; "<" = 188; "_" = 189;
    ">" = 190; "?" = 191; "~" = 192; "{" = 219;
    "|" = 220; "}" = 221; '"' = 222
}


function Get_Key {
    param (
        [Parameter(Mandatory=$true)]
        [string]$Search_Char
    )
    $reg_key = $winkeylist[$Search_Char]
    if ($null -ne $reg_key) {
        return @{
            "Capital" = $false
            "Key" = $reg_key
        }
    } 
    $alt_key = $altkeylist[$Search_Char]
    if ($null -ne $alt_key) {
        return @{
            "Capital" = $true
            "Key" = $alt_key
        }
    }
    return $null
}

# sanitizing for SendKeys
$chars_to_sanitize = @(
    [regex]'[{}]', 
    [regex]'[\[\]]',
    [regex]'[\(\)]',
    [regex]'[+]', 
    [regex]'[\^]', 
    [regex]'[%]', 
    [regex]'[~]'
)
function Sanitize_SendKeys {
    param (
        [Parameter(Mandatory=$true)]
        [string]$String
    )
    $sanitized_string = $String
    $chars_to_sanitize | ForEach-Object {
        $sanitized_string = [regex]::Replace($sanitized_string, $_, '{$&}')
    }
    return $sanitized_string
}

# convert the raw macro keys to their virtual key codes
Write-Host "Converting macro keys to key codes" -ForegroundColor Yellow
$macro_key_codes = @()
$macro_keys_raw | ForEach-Object {
    $macro_key_codes += Get_Key -Search_Char $_
    Write-Host "Macro Key $($macro_key_codes.Length): `"$_`" key code: $($macro_key_codes[-1].Key)" -ForegroundColor Green
}

$prev_state = $false

while ($true) {
    Start-Sleep -Milliseconds 40

    # check if all the macro keys are pressed
    $curr_state = $true
    $macro_key_codes | ForEach-Object {
        if ([API.Keypress]::GetAsyncKeyState($_.Key) -ge 0) {
            $curr_state = $false
            return
        }
    }

    if (-not $curr_state -and $prev_state) {
        start-sleep -Milliseconds 100
        Write-Host $macro_keys_raw[0] -ForegroundColor Green -nonewline
        1..$($macro_keys_raw.Length - 1) | ForEach-Object {
            Write-Host " + $($macro_keys_raw[$_])" -ForegroundColor Green -nonewline
        }
        Write-Host " pressed" -ForegroundColor Green
        
        # grab the user's clipboard
        $clipboard = Get-Clipboard
        if ($debug) { Write-Host "Clipboard: $clipboard" -ForegroundColor Yellow }

        # error handling for empty clipboards
        if ($null -eq $clipboard -or $clipboard.Length -lt 1) {
            Write-Host "No clipboard data found" -ForegroundColor Red

        # prints a single line
        } elseif ($clipboard.gettype().Name -eq "String") {
            Write-Host "Printing clipboard as string" -ForegroundColor Green
            $sanitized_string = Sanitize_SendKeys -String $clipboard
            if ($debug) { Write-Host "Sanitized string: $sanitized_string" -ForegroundColor Yellow }
            [System.Windows.Forms.SendKeys]::SendWait($sanitized_string)

        # prints multiple lines
        } elseif ($clipboard.gettype().Name -eq "Object[]") {
            Write-Host "Printing clipboard as array" -ForegroundColor Green
            for ($line = 0; $line -lt $clipboard.Length; $line++) {
                if ($debug) { Write-Host "Printing line #${line}: $($clipboard[$line])" }
                if ($clipboard[$line].length -eq 0) {
                    [System.Windows.Forms.SendKeys]::SendWait("~")
                    continue
                }
                $sanitized_string = Sanitize_SendKeys -String $clipboard[$line]
                if ($debug) { Write-Host "Sanitized string: $sanitized_string" -ForegroundColor Yellow }
                [System.Windows.Forms.SendKeys]::SendWait($sanitized_string)
                if ($line -lt $clipboard.Length - 1) {
                    [System.Windows.Forms.SendKeys]::SendWait("~")
                }
            }
        } 

        # error handling for bad clipboard data type
        else {
            Write-Host "Unknown clipboard type: $($clipboard.gettype().Name)" -ForegroundColor Red
        }
    }

    # ensures that the print only happens on the first key press
    $prev_state = $curr_state
}