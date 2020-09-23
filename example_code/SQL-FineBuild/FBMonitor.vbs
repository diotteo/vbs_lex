''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBMonitor.vbs  
'  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
'
'  Purpose:      Script to monitor if FineBuild has hung, and if so to trigger a reboot
'
'  Parameters:   ProcessId - Current Process Id
'                WaitTime  - Number of minutes to waid before triggering a reboot
'
'  Author:       Ed Vassie
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     13 Jan 2020  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colArgs
Dim objShell
Dim strPathFB, strProcessId, strStartTime, strWaitEnd, strWaitTime

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case Else
      Call ProcessMonitor()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case (err.Number = 3010) And (err.Description = "Reboot required")
      Call FBLog("***** Reboot in progress *****")
    Case Left(err.Description, 11) = "Stop forced"
      Call FBLog("***** " & err.Description & " *****")
    Case err.Number <> 0 
      Call FBLog("***** Error has occurred *****")
      If strProcessIdLabel <> "" Then
        Call FBLog(" Process    : " & strProcessIdLabel & ": " & strProcessIdDesc)
      End If
      If err.Number <> "" Then
        Call FBLog(" Error code : " & err.Number)
      End If
      If err.Source <> "" Then
        Call FBLog(" Source     : " & err.Source)
      End If
      If err.Description <> "" Then
        Call FBLog(" Description: " & err.Description)
      End If
      If strDebugDesc <> "" And strDebugDesc <> err.Description Then
        Call FBLog(" Last Action: " & strDebugDesc)
      End If
      If strDebugMsg1 <> "" Then
        Call FBLog(" " & strDebugMsg1)
      End If
      If strDebugMsg2 <> "" Then
        Call FBLog(" " & strDebugMsg2)
      End If
      Call FBLog(" FBMonitor failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation processing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBManageBoot.vbs"
  Include "FBUtils.vbs"
  Call SetProcessIdCode("FBMO")

  Set colArgs       = Wscript.Arguments.Named
  strStartTime      = Now()
  strProcessId      = GetParam(Null, "ProcessId", GetBuildfileValue("ProcessId"))
  strWaitTime       = GetParam(Null, "WaitTime",  "5")

  strWaitEnd        = DateAdd("N", strWaitTime, strStartTime)
  strWaitEnd        = GetStdDateTime(strWaitEnd)

End Sub


Function GetParam(colParam, strParam, strDefault) 
' Get parameter value
  Dim strValue

' Find parameter value in XML configuration file
  Select Case True
    Case IsNull(colParam)
      strValue      = strDefault
    Case IsNull(colParam.getAttribute(strParam))
      strValue      = strDefault
    Case Else
      strValue      = colParam.getAttribute(strParam)
  End Select

' Apply any parameter overide from CSCRIPT arguments
  Select Case True
    Case Not colArgs.Exists(strParam)
      ' Nothing
    Case Else
      strValue      = colArgs.Item(strParam)
  End Select

  GetParam          = strValue

End Function


Sub ProcessMonitor()
  Call DebugLog("ProcessMonitor:")
  Dim strMessage

  strMessage        = "FBMonitor: Waiting until " & strWaitEnd & " for completion of " & strProcessId & " " & GetBuildfileValue("ProcessIdDesc")
  Call DebugLog(strMessage)
  WScript.Echo strMessage

  Do While GetStdDateTime("") < strWaitEnd
    Wscript.Sleep 10000
    Call LinkBuildfile("")
    If GetBuildfileValue("ProcessId") <> strProcessId Then
      Exit Do
    End If
  Loop

  Call LinkBuildfile("")
  If GetBuildfileValue("ProcessId") = strProcessId Then
    Call SetupReboot(strProcessId, "FineBuild may be hanging")
  End If

  Call DebugLog("FBMonitor for " & strProcessId & " Ending")

End Sub


Function Include(strFile)
  Dim objFSO, objFile
  Dim strFilePath, strFileText

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      err.Raise 8, "", "ERROR: This process must be run by SQLFineBuild.bat"
    Case Else
      Set objFSO        = CreateObject("Scripting.FileSystemObject")
      strFilePath       = strPathFB & "Build Scripts\" & strFile
      Set objFile       = objFSO.OpenTextFile(strFilePath)
      strFileText       = objFile.ReadAll()
      objFile.Close 
      ExecuteGlobal strFileText
  End Select

End Function


End Class
