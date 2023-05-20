Set WshShell = CreateObject("WScript.Shell")
WshShell.SendKeys(chr(&hAD))
WScript.sleep 500
do While True = True
     WScript.Sleep 600
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys(chr(&hAF))
loop
