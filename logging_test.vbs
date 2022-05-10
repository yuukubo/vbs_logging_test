Option Explicit

dim stage

''''''''''''''''''''''
stage = "1"

Dim objPath
Set objPath = CreateObject("Scripting.FileSystemObject").GetFolder(".")


''''''''''''''''''''''
stage = "2"

dim datetimeNow, now_YYYYMMDDhhmmss

datetimeNow = Now()

now_YYYYMMDDhhmmss= Year(datetimeNow)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Month(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Day(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Hour(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Minute(datetimeNow) , 2)
now_YYYYMMDDhhmmss= now_YYYYMMDDhhmmss & Right("0" & Second(datetimeNow) , 2)
'MsgBox now_YYYYMMDDhhmmss '20210802012136

''''''''''''''''''''''
stage = "3"
 
Dim fileName
Dim filePath
 
fileName = WScript.ScriptName
  
filePath = WScript.ScriptFullName
 
Dim objFileSys
 
Set objFileSys = CreateObject("Scripting.FileSystemObject")
 
'WScript.Echo objFileSys.getBaseName(filePath)

''''''''''''''''''''''
stage = "4"

dim fso
dim f

set fso = CreateObject("Scripting.FileSystemObject")
set f = fso.OpenTextFile(objPath & "\" & now_YYYYMMDDhhmmss &  "_" & objFileSys.getBaseName(filePath) & "_log.txt", 8, True)
f.WriteLine("start : " & Now())
f.Close


''''''''''''''''''''''
stage = "5"

dim strLogFile, objFSO, strLogMsg
strLogFile = objPath & "\" & now_YYYYMMDDhhmmss &  "_" & objFileSys.getBaseName(filePath) & "_log.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")

strLogMsg = "script start" : LogWrite(strLogMsg)
strLogMsg = ""

strLogMsg = "logging test Start" : LogWrite(strLogMsg)
strLogMsg = ""

strLogMsg = "logging test End" : LogWrite(strLogMsg)
strLogMsg = ""


''''''''''''''''''''''
stage = "6"

strLogMsg = "2 / 1 = " & 2 / 1 : LogWrite(strLogMsg)
strLogMsg = ""
'msgbox 2 / 1

On Error Resume Next
strLogMsg = "2 / 0 = " & 2 / 0 : LogWrite(strLogMsg)
strLogMsg = ""

If Err.Number <> 0 Then
    strLogMsg = "Err.Number : " & Err.Number & " Err.Description : " & Err.Description & " Err.Source : " & Err.Source & ", stage : " & stage : LogWrite(strLogMsg)
    strLogMsg = ""
End If

On Error Goto 0
Err.Clear

On Error Resume Next
stage = "7"

CInt("•¶Žš—ñ")
If Err.Number <> 0 Then
    strLogMsg = "Err.Number : " & Err.Number & " Err.Description : " & Err.Description & " Err.Source : " & Err.Source & ", stage : " & stage : LogWrite(strLogMsg)
    strLogMsg = ""
End If

On Error Goto 0
Err.Clear

strLogMsg = "script end" : LogWrite(strLogMsg)
strLogMsg = ""


''''''''''''''''''''''

WScript.Quit


''''''''''''''''''''''

Function LogWrite(strMsg)

  dim objFile

  strMsg = "[" & FormatDateTime(Now, 0) & "]" & Space(1) & strMsg
  Set objFile = objFSO.OpenTextFile(strLogFile, 8, True, 0)
  objFile.WriteLine strMsg
  objFile.Close

  LogWrite = 0

End Function
