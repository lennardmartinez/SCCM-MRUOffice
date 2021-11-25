Dim oFSO, oShell, oReg, strAppData, strKeyNew, strKeyOld, strProgramFiles, strOfficeApp, strOfficeReg, strRegFile
Const HKCU = &H80000001

Sub subParse
 strRegFile = strAppData & "\mru\" & strOfficeReg & ".reg"
 If oFSO.FileExists(strRegFile) Then
 Else
  strKeyNew = "Software\Microsoft\Office\15.0\" & strOfficeReg & "\User MRU\"
  If KeyExists(HKCU, strKeyNew) = True Then
  Else
   WScript.Sleep 1000
   oShell.Run """" & strProgramFiles & strOfficeApp & ".exe""", 0
   WScript.Sleep 20000
   oShell.Run "taskkill /f /t /im " & strOfficeApp & ".exe", 0
   WScript.Sleep 1000
   oReg.EnumKey HKCU, strKeyNew, arrSubKeys
   For Each strSubkey In arrSubKeys
    If Instr(strSubKey,"AD_")>0 then
     strKeyHash = strSubKey
    End If
   Next
   oShell.Run "reg export " & Chr(34) & "HKCU\" & strKeyOld & strOfficeReg & "\File MRU" & Chr(34) & " " & Chr(34) & strRegFile & Chr(34) & " /y", 0
   strOldText = strKeyOld & strOfficeReg
   strNewText = strKeyNew & strKeyHash
   WScript.Sleep 2000
   set oFile = oFSO.OpenTextFile(strRegFile, 1,,-1)
   strText = oFile.ReadAll
   oFile.Close
   strNewText = Replace(strText, strOldText, strNewText)
   Set oFile = oFSO.OpenTextFile(strRegFile, 2)
   oFile.Write strNewText
   oFile.Close
   oShell.Run "regedit.exe /s " & Chr(34) & strRegFile & Chr(34), 0, True
  End If
 End If
End Sub

Function KeyExists(Key, KeyPath)
 Dim oReg: Set oReg = GetObject("winmgmts:!root/default:StdRegProv")
 If oReg.EnumKey(Key, KeyPath, arrSubKeys) = 0 Then
  KeyExists = True
 Else
  KeyExists = False
 End If
End Function

Set oFSO = CreateObject( "Scripting.FileSystemObject" )
Set oShell= Wscript.CreateObject("WScript.Shell")
Set oReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
strKeyOld = "Software\Microsoft\Office\14.0\"
strAppData = oShell.ExpandEnvironmentStrings("%appdata%")

on error resume next
If KeyExists(HKCU, strKeyOld) = True Then
 strProgramFiles = ""
 If oFSO.FileExists ("C:\Program Files (x86)\Microsoft Office\Office15\OSPP.VBS") Then
  strProgramFiles =  "C:\Program Files (x86)\Microsoft Office\Office15\"
 ElseIf oFSO.FileExists ("C:\Program Files\Microsoft Office\Office15\OSPP.VBS") Then
  strProgramFiles = "C:\Program Files\Microsoft Office\Office15\"
 End If
 If strProgramFiles = "" Then
 Else
  on error resume next
  oFSO.CreateFolder strAppData & "\mru"
  strOfficeApp = "Excel"
  strOfficeReg = "Excel"
  subParse
  strOfficeApp = "PowerPnt"
  strOfficeReg = "PowerPoint"
  subParse
  strOfficeApp = "WinWord"
  strOfficeReg = "Word"
 subParse
 End If
End If
on error GoTo 0
Wscript.Quit