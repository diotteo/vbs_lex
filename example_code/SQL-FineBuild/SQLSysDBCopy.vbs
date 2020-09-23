''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  SqlSysDBCopy.vbs  
'  Copyright FineBuild Team © 2008 - 2017.  Distributed under Ms-Pl License
'
'  Purpose:      Copies critical system DB files to a backup location
'
'  Author:       Ed Vassie
'
'  Change History
'  Version  Author        Date         Description
'  2.0      Ed Vassie     01 Oct 208   SQL Server 2008 version
'  1.01     Ed Vassie     10 Feb 2008  Fixed parameter handling 
'  1.0      Ed Vassie     02 Feb 2008  Initial version for FineBuild v1.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim objFile, objFSO, objShell, objWMI, objWMIReg
Dim strCmd, strDirSystemDataBackup, strDirProg, strEdition, strHKLM, strHKLMSQL, strInstAgent, strInstance, strInstNode, strInstRegSQL, StrInstSQL, strPath, strPathFB, strPathNew, strSQLVersion
Dim strServer, strServInst, strType

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Call FBLog("Run at " & cStr(Time()) & " on " & cStr(Date()) & " server " & strServer)

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strType = "CLIENT" 
      ' Nothing
    Case Else
      Call SysDBCopy ()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case err.Number = 0 
      ' Nothing
    Case Else
      Call FBLog("***** Error has occurred *****")
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
      Call FBLog("Copy of System Databases failed")
  End Select

  Wscript.Quit(err.Number)

End Sub


Sub Initialisation ()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"

  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strSQLVersion     = GetBuildfileValue("SQLVersion")
  strHKLM           = &H80000002
  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strDebug          = GetBuildfileValue("Debug")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup") & "\"
  strDirProg        = GetBuildfileValue("DirProg")
  strEdition        = GetBuildfileValue("AuditEdition")
  strInstance       = GetBuildfileValue("Instance")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strServInst       = GetBuildfileValue("ServInst")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strServer         = GetBuildfileValue("AuditServer")
  strType           = GetBuildfileValue("Type")

End Sub


Sub SysDBCopy ()
  Call FBLog("Starting SQLSysDBCopy (FBSC - SQLSysDBCopy.vbs)")
  Dim strPathOld, strPathNew

  Call FBLog("Backup system databases to: " & strDirSystemDataBackup)
  Call FBLog("Copying master Database files")

  strPathOld        = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg0")
  strPathOld        = Mid(strPathOld, 3, (Len(strPathOld) - 12))
  strDebugMsg1      = "Source: " & strPathOld

  Set objFile       = objFSO.GetFile(strPathOld & "master.mdf")
  strPathNew        = strDirSystemDataBackup & objFile.Name
  strDebugMsg2      = "Target: " & strPathNew
  objFile.Copy strPathNew, True

  Set objFile       = objFSO.GetFile(strPathOld & "mastlog.ldf")
  strPathNew        = strDirSystemDataBackup & objFile.Name
  strDebugMsg2      = "Target: " & strPathNew
  objFile.Copy strPathNew, True

  If objFSO.FileExists(strPathOld & "distmdl.mdf") Then
    Call FBLog("Copying Distmdl Database files")
    Set objFile     = objFSO.GetFile(strPathOld & "distmdl.mdf")
    strPathNew      = strDirSystemDataBackup & objFile.Name
    strDebugMsg2    = "Target: " & strPathNew
    objFile.Copy strPathNew, True
    Set objFile     = objFSO.GetFile(strPathOld & "distmdl.ldf")
    strPathNew      = strDirSystemDataBackup & objFile.Name
    strDebugMsg2    = "Target: " & strPathNew
    objFile.Copy strPathNew, True
  End If

  Call FBLog("Copying mssqlsystemresource Database files")

  If strSQLVersion <> "SQL2005" Then
    strPathOld      = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLBinRoot")
    strDebugMsg1    = "Source: " & strPathOld
  End If

  Set objFile       = objFSO.GetFile(strPathOld & "\mssqlsystemresource.mdf")
  strPathNew        = strDirSystemDataBackup & objFile.Name
  strDebugMsg2      = "Target: " & strPathNew
  objFile.Copy strPathNew, True

  Set objFile       = objFSO.GetFile(strPathOld & "\mssqlsystemresource.ldf")
  strPathNew        = strDirSystemDataBackup & objFile.Name
  strDebugMsg2      = "Target: " & strPathNew
  objFile.Copy strPathNew, True

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
