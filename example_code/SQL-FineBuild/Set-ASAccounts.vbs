''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Set-ASAccounts.vbs  
'  Copyright FineBuild Team © 2008 - 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Configures Analysis Services Administration accounts
'
'  Parameters    /Domain:domain /GroupDBA:DBA Sysadmin group /ServAcnt:SQL Server service account /ASServAcnt:AS Service Account
'
'  Author:       Ed Vassie
'
'  Date:         October 2006
'
'  Version:      2.0
'  Version       3.0 All required parameters now taken from Config file
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim objAMOServer, objAssembly, objFile, objFSO, objRole, objRoleMember, objShell, objWMI, objWMIReg
Dim intMemberAdded
Dim strAgtAccount, strAsAccount, strCmd
Dim strDirSystemDataBackup, strDirProg, strDomain, strEdition, strFileASDLL
Dim strGroupDBA, strHKLM, strHKLMSQL, strInstAgent, strInstASCon, strInstASSQL, strInstance, strInstNode
Dim strPath, strPathFB, strPathNew, strServer, strSqlAccount, strSQLVersion, strType

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strType = "CLIENT" 
      ' Nothing
    Case Else
      Call SetupASAccounts(strInstASCon)
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
  strAgtAccount     = GetBuildfileValue("AgtAccount")
  strAsAccount      = GetBuildfileValue("AsAccount")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup")
  strDirProg        = GetBuildfileValue("DirProg")
  strDomain         = GetBuildfileValue("Domain")
  strEdition        = GetBuildfileValue("AuditEdition")
  strFileASDLL      = GetBuildfileValue("FileASDLL")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strInstance       = GetBuildfileValue("Instance")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstASCon      = GetBuildfileValue("InstASCon")
  strInstASSQL      = GetBuildfileValue("InstASSQL")
  strServer         = GetBuildfileValue("AuditServer")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strType           = GetBuildfileValue("Type")

End Sub


Sub SetupASAccounts(strInstASCon)
  Call FBLog("SetupASAccounts: " & strInstASCon & " (Set-ASAccounts.vbs)")

  Select Case True
    Case strSQLVersion <= "SQL2014"
      Set objAMOServer  = CreateObject ("Microsoft.AnalysisServices.Server")
    Case Else
      Set objAMOServer  = CreateObject ("Microsoft.AnalysisServices.Core.Server")
  End Select

  intMemberAdded    = False
  objAMOServer.Connect("Data source =" & strInstASCon)
  For Each objRole In objAMOServer.Roles
    Select Case True
      Case objRole.Name <> "Administrators"
        ' Nop
      Case Else
        Select Case True
          Case strSQLVersion > "SQL2005"
            ' Nothing
          Case Ucase(strAsAccount) = "LOCALSYSTEM" 
            Call AddRoleMember("SYSTEM")
          Case Else
            Call AddRoleMember(strAsAccount)
        End Select
        Select Case True
          Case strSQLVersion > "SQL2005"
            ' Nothing
          Case strGroupDBA = ""
            ' Nothing
          Case Else
            Call AddRoleMember(strGroupDBA) ' Add DBA Group as AS Administrator
        End Select
        Select Case True
          Case strSqlAccount = ""
            ' Nothing
          Case strSqlAccount = strAsAccount
            ' Nothing
          Case Else
            Call AddRoleMember(strSqlAccount) ' Add SQL Server service account as AS Administrator
        End Select
        Select Case True
          Case strAgtAccount = ""
            ' Nothing
          Case strAgtAccount = strAsAccount
            ' Nothing
          Case strAgtAccount = strSqlAccount
            ' Nothing
          Case Else
            Call AddRoleMember(strAgtAccount) ' Add SQL Agent service account as AS Administrator
        End Select
      If intMemberAdded = True Then
        objRole.Update()
      End If
    End Select
  Next

  objAMOServer.Disconnect()
  Set objAMOServer  = Nothing

End Sub


Sub AddRoleMember(strAccount)

  Set objRoleMember  = CreateObject("Microsoft.AnalysisServices.RoleMember")
  objRoleMember.Name = strAccount
  objRole.Members.Add(objRoleMember)
  intMemberAdded     = True

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