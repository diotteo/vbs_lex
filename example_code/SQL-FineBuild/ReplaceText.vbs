''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  ReplaceText.vbs  
'  Copyright FineBuild Team © 2008 - 2015.  Distributed under Ms-Pl License
'
'  Purpose:      Find and Replace text strings in a file
'
'  Author:       Ed Vassie, based on a freeware script published by Microsoft
'
'  Date:         November 2006
'
'  Version:      1
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Const ForReading = 1
Const ForWriting = 2

Dim objFile
Dim objFSO
Dim strFileName
Dim strOldData
Dim strOldText
Dim strNewData
Dim strNewText

err.Number  = 0
strFileName = Wscript.Arguments(0)
strOldText  = Wscript.Arguments(1)
strNewText  = Wscript.Arguments(2)

Set objFSO  = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile(strFileName, ForReading)

strOldData  = objFile.ReadAll
objFile.Close
strNewData  = Replace(strOldData, strOldText, strNewText, 1, -1 , 1) ' 1=Start Pos;-1=Replace all;1=Ignore case

Set objFile = objFSO.OpenTextFile(strFileName, ForWriting)
objFile.WriteLine strNewData
objFile.Close

Select Case True
  Case err.Number = 0
    ' Nothing
  Case Else
    Wscript.Echo "Error " + err.Number + " occurred"
End Select
