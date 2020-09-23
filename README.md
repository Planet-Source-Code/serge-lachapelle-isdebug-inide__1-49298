<div align="center">

## IsDebug \- InIDE


</div>

### Description

return TRUE if inside IDE, return FALSE if in compiled EXE, work in standard EXE and in DLL
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Serge Lachapelle](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/serge-lachapelle.md)
**Level**          |Intermediate
**User Rating**    |5.0 (10 globes from 2 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/serge-lachapelle-isdebug-inide__1-49298/archive/master.zip)





### Source Code

```
Private Type MODULEENTRY32
  dwSize As Long
  th32ModuleID As Long
  th32ProcessID As Long
  GlblcntUsage As Long
  ProccntUsage As Long
  modBaseAddr As Long
  modBaseSize As Long
  hModule As Long
  szModule As String * 256
  szExePath As String * 260
End Type
Private Const TH32CS_SNAPMODULE = &H8
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function Module32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function Module32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As MODULEENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Public Function IsDebug() As Boolean
 Dim qwe As String
 Dim hProcess As MODULEENTRY32, hMod&, hSnapshot&
 hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId)
 hProcess.dwSize = Len(hProcess)
 hMod = Module32First(hSnapshot, hProcess)
 qwe = Left$(hProcess.szExePath, InStr(hProcess.szExePath, vbNullChar) - 1)
 If LCase$(Right$(qwe, 4)) <> ".exe" Then
  Do
   hMod = Module32Next(hSnapshot, hProcess)
   qwe = Left$(hProcess.szExePath, InStr(hProcess.szExePath, vbNullChar) - 1)
  Loop Until (LCase$(Right$(qwe, 4)) = ".exe") Or (hMod = 0)
 End If
 IsDebug = LCase(hProcess.szExePath) Like "*vb#.exe*"
 CloseHandle hSnapshot
End Function
```

