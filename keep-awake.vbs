Set sh = WScript.CreateObject("WScript.Shell")
Do While True
  WScript.Sleep 60000
  sh.SendKeys "{CAPSLOCK}{CAPSLOCK}"
Loop
