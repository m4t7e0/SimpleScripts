Dim objResult
Set objShell = WScript.CreateObject("WScript.Shell")    

Do While True
  objResult = objShell.SendKeys("{NUMLOCK}{NUMLOCK}")
  WScript.Sleep (6000)
Loop

