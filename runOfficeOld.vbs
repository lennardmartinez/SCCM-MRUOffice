Dim bExists, oFSO, oShell, strAppData, strComputer, strKey, strProgramFiles, StrOfficeApp

Sub subParse
 If oFSO.FileExists(oAppData & "\mru\" & strOfficeApp & ".reg") Then
 Else
  strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\15.0\" & strOfficeApp & "User MRU\"
  on error resume next
  set present = WshShell.RegRead(strKey)
  bExists=False   
  If err.number<>0 Then
   If right(present,1)="\" Then
    If InStr(1,err.description,ssig,1)<>0 Then
     bExists=True
    Else
     bExists=False
    End If
   Else
    bExists=False
   End If
   err.clear
  Else
   bExists=True
  End If
  on error GoTo 0
  If bExists=vbFalse Then
   WScript.Sleep 1000
   oShell.Run """" & strProgramFiles & strOfficeApp & ".exe""", 0
   WScript.Sleep 20000
   oShell.Run "taskkill /f /t /im " & strOfficeApp & ".exe", 0
   WScript.Sleep 1000
   oShell.Run("""\\wesc-fp-01.wesc.internal\IT Updates\Scripts\Office2013\Production\" & strOfficeApp & " MRU.bat"""), 0
  End If
 End If
End Sub

Function Ping(strHost)
 Dim objPing, oRetStatus, bReturn
 Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address='" & strHost & "'")
 For Each oRetStatus In objPing
 If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
  bReturn = False
 Else
  bReturn = True
 End If
 Set oRetStatus = Nothing
 Next
 Set objPing = Nothing
 Ping = bReturn
End Function

Set oFSO = CreateObject( "Scripting.FileSystemObject" )
Set oShell= Wscript.CreateObject("WScript.Shell")

strKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\14.0\"

on error resume next
present = WshShell.RegRead(strKey)
If err.number<>0 Then
 If right(strKey,1)="\" Then
  If InStr(1,err.description,ssig,1)<>0 Then
   bExists=True
  Else
   bExists=False
  End If
 Else
  bExists=False
 End If
 err.clear
Else
 bExists=True
End If
on error GoTo 0
If bExists=vbTrue Then
 If Ping("WESC-FP-01.wesc.internal") Then
  strProgramFiles = ""
  If oFSO.FileExists ("C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS") Then
   strProgramFiles =  "C:\Program Files (x86)\Microsoft Office\Office15\"
  ElseIf oFSO.FileExists ("C:\Program Files\Microsoft Office\Office15\OSPP.VBS") Then
   strProgramFiles = "C:\Program Files\Microsoft Office\Office15\"
  End If
  If strProgramFiles = "" Then
  Else
   oAppData = oShell.ExpandEnvironmentStrings("%appdata%")
   on error resume next
   oFSO.CreateFolder oAppData & "\mru"
   strOfficeApp = "Excel"
   subParse 
   strOfficeApp = "PowerPnt"
   subParse 
   strOfficeApp = "WinWord"
   subParse
  End If
 End If
End If
Wscript.Quit