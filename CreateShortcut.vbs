Set objshell = WScript.CreateObject("WScript.Shell")
CurDir = objshell.CurrentDirectory
Quote = """"
Set Link = objshell.CreateShortcut(".\Keygen.lnk")
    Link.TargetPath = CurDir+".\lib\pythonw.exe"
    Link.Arguments = Quote+CurDir+"\bin\keygen.pyw"+Quote
    Link.IconLocation = CurDir+"\etc\multi-res.ico"
    Link.Description = "Key Generator"
    Link.WorkingDirectory = CurDir+"\bin\"
    Link.Save