''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBConfigDiscover.vbs  
'  Copyright FineBuild Team © 2017 - 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Create a Parameter File to reproduce with FineBuild the SQL Instance that is being discovered
'
'  Author:       Ed Vassie
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     25 Nov 2016  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colArgs, colBuildfile, colPrcEnvVars
Dim interrSave, intIndex
Dim objApp, objAutoUpdate, objBuildfile, objFSO, objReportFile, objShell, objWMI, objWMIREG
Dim strAnyKey, strBuildFile, strClusterName, strDebug, strDebugDesc, strDebugMsg1, strDebugMsg2, strEdition, strEdType, strFileArc, strInstance, strInstNode, strInstSQL
Dim strHKLM, strHKLMSQL
Dim strMainInstance, strMsgError, strMsgInfo, strMsgWarning, strOSLevel, strOSName, strOSType, strOSVersion
Dim strPathFBScripts, strProcessId, strProcessIdDesc, strProcessIdLabel, strRebootStatus, strReportFile, strReportOnly
Dim strServer, strSQLVersion, strStatusComplete, strStatusBypassed, strStopAt, strType
Dim strUserName, strValidateError, strVersionFB

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case Else
      Call ProcessDiscovery()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strProcessId > "7ZZ"
      ' Nothing
    Case err.Number = 0 
      ' Nothing
    Case (err.Number = 3010) And (err.Description = "Reboot required")
      Call FBLog("***** Reboot in progress *****")
    Case Left(err.Description, 11) = "Stop forced"
      Call FBLog("***** " & err.Description & " *****")
    Case Else
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
      Call FBLog(" FineBuild Instance Discovery failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation ()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strBuildFile      = objShell.ExpandEnvironmentStrings("%SQLLOGTXT%")
  If strBuildFile = "%SQLLOGTXT%" Then
    err.Raise 8, "", "ERROR: This process must be run by SQLFineBuild.bat"
  End If

  Set objApp        = CreateObject ("Shell.Application")
  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objBuildfile  = CreateObject ("Microsoft.XMLDOM")  
  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colArgs       = Wscript.Arguments.Named
  Set colPrcEnvVars = objShell.Environment("Process")

  strBuildFile      = Mid(strBuildFile, 2, Len(strBuildFile) - 6) & ".xml"
  objBuildFile.async = False
  objBuildfile.load(strBuildFile)
  Set colBuildFile  = objBuildfile.documentElement.selectSingleNode("BuildFile")
  strDebug          = GetBuildfileValue("Debug")

  strHKLM           = &H80000002
  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strClusterName    = GetBuildfileValue("ClusterName")
  strEdition        = GetBuildfileValue("AuditEdition")
  strPathFBScripts  = GetBuildfileValue("PathFBScripts")
  strEdType         = GetBuildfileValue("EdType")
  strFileArc        = GetBuildfileValue("FileArc")
  strInstance       = GetBuildfileValue("Instance")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strMainInstance   = GetBuildfileValue("MainInstance")
  strMsgError       = GetBuildfileValue("MsgError")
  strMsgInfo        = GetBuildfileValue("MsgInfo")
  strMsgWarning     = GetBuildfileValue("MsgWarning")
  strOSLevel        = GetBuildfileValue("OSLevel")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strProcessId      = GetBuildfileValue("ProcessId")
  strProcessIdLabel = GetBuildfileValue("ProcessId")
  strRebootStatus   = GetBuildfileValue("RebootStatus")
  strReportFile     = GetBuildfileValue("ReportFile")
  strReportOnly     = GetBuildfileValue("ReportOnly")
  strServer         = GetBuildfileValue("AuditServer")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strStatusComplete = GetBuildfileValue("StatusComplete")
  strStatusBypassed = GetBuildfileValue("StatusBypassed")
  strStopAt         = GetBuildfileValue("StopAt")
  strType           = GetBuildfileValue("Type")
  strUserName       = GetBuildfileValue("AuditUser")
  strValidateError  = GetBuildfileValue("ValidateError")
  strVersionFB      = GetBuildfileValue("VersionFB")

End Sub


Sub ProcessDiscovery()
  Call FBLog("FineBuild Instance Discovery processing (FBConfigDiscover.vbs)")

  Select Case True
    Case objFSO.FileExists(strPathFBScripts & "FBConfigDiscover.ps1")
      strCmd        = "POWERSHELL -ExecutionPolicy Bypass -File """ & strPathFBScripts & "Get-FBConfig.ps1"" " & strGroupDBA
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
    Case Else
      Call SetBuildMessage(strMsgError, "FineBuild Instance Discovery is not supported in this version of SQL FineBuild")
      Exit Sub
  End Select

  Call FBLog(" FineBuild Instance Discovery processing" & strStatusComplete)

End Sub


Sub SetBuildMessage (strType, strMessage)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935
  Dim colMessage
  Dim objAttribute
  Dim intBuildMsg
  Dim strMessageText

  Select Case True
    Case strMessage = ""
      Exit Sub
  End Select

  strMessageText    = strMessage
  strMessageText    = HidePassword(strMessageText, "Password")
  strMessageText    = HidePassword(strMessageText, "Pwd")

  intBuildMsg       = GetBuildfileValue("BuildMsg")
  If intBuildMsg = "" Then
    intBuildMsg     = 0
  End If

  intBuildMsg       = intBuildMsg + 1
  Set objAttribute  = objBuildFile.createAttribute("Msg" & CStr(intBuildMsg))
  Set colMessage    = objBuildfile.documentElement.selectSingleNode("Message")
  objAttribute.Text = Ucase(strType) & ": " & strMessageText
  colMessage.Attributes.setNamedItem objAttribute
  objBuildFile.documentElement.appendChild colMessage

  objBuildFile.save strBuildFile
  Call SetBuildfileValue("BuildMsg",           intBuildMsg)

  Select Case True
    Case strType = "ERROR" 
      Call FBLog(" ")
      Call FBLog(" " & strType & ": " & strMessageText)
      err.Raise 8, "", strType & ": " & strMessageText
    Case strType = "WARNING" 
      Call FBLog(" ")
      Call FBLog(" " & strType & ": " & strMessageText)
    Case Else
      Call FBLog(" " & strMessageText)
  End Select

End Sub


Function GetBuildfileValue(strParam) 
' Get value from Buildfile
Dim strValue

  Select Case True
    Case IsNull(colBuildfile.getAttribute(strParam))
      strValue      = ""
    Case Else
      strValue      = colBuildfile.getAttribute(strParam)
  End Select

  GetBuildfileValue = strValue

End Function


Sub SetBuildfileValue (strName, strValue)
  Call DebugLog("Add Buildfile value " & strName & ": " & strValue)
  ' Code based on http://www.vbforums.com/showthread.php?t=480935
  Dim objAttribute

  Select Case True
    Case Not IsNull(colBuildfile.getAttribute(strName))
      colBuildfile.setAttribute strName, strValue
    Case Else
      Set objAttribute  = objBuildFile.createAttribute(strName)
      objAttribute.Text = strValue
      colBuildFile.Attributes.setNamedItem objAttribute
      objBuildFile.documentElement.appendChild colBuildfile
  End Select

  objBuildFile.save strBuildFile

End Sub


Sub Util_RegWrite(strRegKey, strRegValue, strRegType)
  Call DebugLog("(" & strProcessIdLabel & ") " & "Write " & strRegKey)

  err.Number        = objShell.RegWrite(strRegKey, strRegValue, strRegType)
  intErrSave        = err.Number
  strErrSave        = err.Description
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case Else
      err.Raise intErrSave, "", strErrSave
  End Select

End Sub


Sub Util_RunExec(strCmd, strMessage, strResponse, intOK)
  Call DebugLog("(" & strProcessIdLabel & ") Exec " & strCmd) 
  Dim objCmd
  Dim strBox1, strBox2, strStdOut

  On Error Resume Next
  strBox1           = "[" & strResponseYes & "/" & strResponseNo & "]"
  strBox2           = "(" & strResponseYes & "/" & strResponseNo & ")?"
  Set objCmd        = objShell.Exec(strCmd)
  Select Case True
    Case Not IsObject(objCmd)
      intErrSave    = 8
      strErrSave    = "Command not recognised"
    Case Else
      Select Case True
        Case strMessage = "EOF"
          objCmd.StdIn.Close
        Case Left(strCmd, 11) = "POWERSHELL "
          objCmd.StdIn.Close
      End Select
      While Not objCmd.StdOut.AtEndOfStream
        strStdOut       = objCmd.StdOut.ReadLine()
        Select Case True
          Case Right(strStdOut, Len(strBox1)) = strBox1
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Right(strStdOut, Len(strBox2)) = strBox2
            objCmd.StdIn.Write strResponse & vbCrLf
          Case Left(strStdOut, Len(strAnyKey)) = strAnyKey
            objCmd.StdIn.Write strResponse & vbCrLf
          Case strMessage = ""
            ' Nothing
          Case Right(strStdOut, Len(strMessage)) = strMessage
            objCmd.StdIn.Write strResponse & vbCrLf
        End Select
      Wend
      While objCmd.Status = 0
        Wscript.Sleep 100
      WEnd
      intErrsave    = objCmd.ExitCode
      strErrSave    = err.Description
  End Select

  On Error Goto 0
  Select Case True
    Case intErrSave = 0 
      ' Nothing
    Case intErrSave = intOK
      ' Nothing
    Case intOK      = -1
      Call DebugLog("Command ended with code: " & intErrSave)
    Case Else
      err.Raise intErrSave, "", strErrSave
  End Select
  err.Clear

End Sub


Sub FBLog(strText)
  Dim strLogText

  strLogText        = strText
  
  strLogText        = CStr(Date()) & " " & CStr(Time()) & " FBCD " & strLogText
  Wscript.Echo strLogText

End Sub


Sub DebugLog(strDebugText)

  strDebugDesc      = strDebugText

  If strDebug = "YES" Then
    Call FBLog(" >" & strDebugText)
  End If

  strDebugMsg1      = ""
  strDebugMsg2      = ""

End Sub


Function HidePassword(strText, strKeyword)
  ' Change any passwords to ********
  Dim intIdx, intFound
  Dim strLogText

  strLogText        = strText
  intIdx = Instr(1, strLogText, strKeyword, vbTextCompare)
  While intIdx > 0
    intFound        = 0
    intIdx          = intIdx + Len(strKeyword)
    While Instr(""":='", Mid(strLogText, intIdx, 1)) > 0 
      intIdx        = intIdx + 1
      intFound      = 1
    Wend
    While Instr(""",/' ", Mid(strLogText, intIdx, 1)) = 0 And IntFound > 0
      strLogText    = Left(strLogText, intIdx - 1) & Chr(01) & Mid(strLogText, intIdx + 1)
      intIdx        = intIdx + 1
    Wend
    intIdx          = Instr(intIdx, strLogText, strKeyword, vbTextCompare)
  WEnd
  While Instr(strLogText, Chr(01) & Chr(01)) > 0
    strLogText      = Replace(Replace(Replace(strLogText, Chr(01) & Chr(01) & Chr(01) & Chr(01), Chr(01)), Chr(01) & Chr(01) & Chr(01), Chr(01)), Chr(01) & Chr(01), Chr(01))
  Wend
  strLogText        = Replace(strLogText, Chr(01), "**********")
  HidePassword      = strLogText

End Function


End Class