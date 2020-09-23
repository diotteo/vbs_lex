''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild3InstallFixes.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      SQL SERVER Service Pack and Cumulative Hotfix Install 
'
'  Author:       Ed Vassie
'
'  Date:         28 Dec 2007
'
'  Change History
'  Version  Author        Date         Description
'  2.2.0    Ed Vassie     15 Jul 2011  Created common code base for all SQL versions
'  2.1.0    Ed Vassie     18 Jun 2010  Initial version for SQL Srver 2008 R2
'  2.0.0    Ed Vassie     04 Aug 2008  Improved error handling
'                                      Initial version for SQL Server 2008
'  1.02     Ed Vassie     18 Feb 2008  Move display of ProcessId labels to assist debugging
'                                      All processing wrapped in SEE_MASK_NOZONECHECKS
'  1.01     Ed Vassie     10 Feb 2008  Convert Instance name to Upper case for comparisoms
'  1.0      Ed Vassie     02 Feb 2008  Initial version for FineBuild v1.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colBuild, colFolders, colPrcEnvVars
Dim objDrive, objFile, objFSO, objFolder, objShell, objWMI, objWMIReg
Dim strAction, strActionSQLDB, strAdminPassword, strAltFile, strAnyKey
Dim strClusterGroupAS, strClusterGroupSQL, strClusterAction, strClusterName, strClusterPassive, strCmd
Dim strDomain, strEdition, strFileArc, strHKLMSQL, strInstance, strInstAgent, strInstAS, strInstNode, strInstParm, strInstRegSQL, strInstSQL
Dim strIAcceptLicenseTerms, strInstLog, strMainInstance, strMode, strOSLanguage, strOSName, strOSType, strOSVersion
Dim strPath, strPathAddComp, strPathFB, strPathFBScripts, strPathNew, strPathLog, strPathTemp, strResSuffix
Dim strSetupSQLASCluster, strSetupSQLAS, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLTools
Dim strServer, strSetupLog, strSPLevel, strSPCULevel, strSQLLanguage, strSQLMedia, strPCUSource, strCUSource, strSQLSPMedia
Dim strSQLVersion, strSQLVersionFull, strSQLVersionNum, strStopAt, strType, strUserName, strWaitLong, strWaitShort, strWOWX86

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "3Z"
      ' Nothing
    Case Else
      Call SetupFixes()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "3Z"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("SQL Server Fixes Install complete", 2, "SQL Server Fixes Install" ,64)
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
      Call FBLog(" SQL Server Fixes install failed")
    End Select

  Wscript.Quit(err.Number)

End Sub


Sub Initialisation ()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBUtils.vbs"
  Include "FBManageBoot.vbs"
  Include "FBManageInstall.vbs"
  Include "FBManageService.vbs"
  Call SetProcessIdCode("FB3F")

  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")

  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAction         = GetBuildfileValue("Action")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strAdminPassword  = GetBuildfileValue("AdminPassword")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strClusterGroupAS = GetBuildfileValue("ClusterGroupAS")
  strClusterGroupSQL  = GetBuildfileValue("ClusterGroupSQL")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterPassive = GetBuildfileValue("ClusterPassive")
  strCUSource       = GetBuildfileValue("CUSource")
  strDomain         = GetBuildfileValue("Domain")
  strEdition        = GetBuildfileValue("AuditEdition")
  strFileArc        = GetBuildfileValue("FileArc")
  strIAcceptLicenseTerms = GetBuildfileValue("IAcceptLicenseTerms")
  strInstance       = GetBuildfileValue("Instance")
  strInstAgent      = GetBuildfileValue("InstAgent")
  strInstAS         = GetBuildfileValue("InstAS")
  strInstLog        = GetBuildfileValue("InstLog")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strMainInstance   = GetBuildfileValue("MainInstance")
  strMode           = GetBuildfileValue("Mode")
  strOSLanguage     = GetBuildfileValue("OSLanguage")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathAddComp    = FormatFolder("PathAddComp")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPCUSource      = GetBuildfileValue("PCUSource")
  strProcessId      = GetBuildfileValue("ProcessId")
  strProcessIdLabel = GetBuildfileValue("ProcessId")
  strResSuffix      = GetBuildfileValue("ResSuffix")
  strServer         = GetBuildfileValue("AuditServer")
  strSetupLog       = Ucase(objShell.ExpandEnvironmentStrings("%SQLLOGTXT%"))
  strSetupLog       = Left(strSetupLog, InStrRev(strSetupLog, "\"))
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLASCluster = GetBuildfileValue("SetupSQLASCluster")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strSQLLanguage    = GetBuildfileValue("SQLLanguage")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSPLevel        = GetBuildfileValue("SPLevel")
  strSPCULevel      = GetBuildfileValue("SPCULevel")
  strSQLMedia       = GetBuildfileValue("PathSQLMedia")
  strSQLSPMedia     = FormatFolder("PathSQLSP")
  strStopAt         = GetBuildfileValue("StopAt")
  strType           = GetBuildfileValue("Type")
  strUserName       = GetBuildfileValue("AuditUser")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")
  strWOWX86         = GetBuildfileValue("WOWX86")

  Select Case True
    Case strType = "CLIENT"
      strInstParm   = ""
    Case GetBuildfileValue("InstParm") = ""
      strInstParm   = "/ALLINSTANCES"
    Case strMainInstance = "YES"
      strInstParm   = "/ALLINSTANCES"
    Case Else
      strInstParm   = "/INSTANCENAME=" & strInstance
  End Select

  If strClusterPassive = "YES" Then
    strInstParm     = strInstParm & " /CLUSTERPASSIVE"
  End If

End Sub


Sub SetupFixes()
  Call SetProcessId("3", strSQLVersion & " Fixes processing (FineBuild3InstallFixes.vbs)")
  Dim strSPInclude

  Call SetUpdate("ON")
  Call SetSQLCompatFlags()
  strSPInclude      = GetBuildfileValue("SPInclude")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3A"
      ' Nothing
    Case GetBuildfileValue("SetupSP") <> "YES"
      ' Nothing
    Case strSPInclude <> ""
      Call SetBuildfileValue("SetupSPStatus", strStatusBypassed & strSPInclude)
    Case Else
      Call SetupServicePack()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3BA"
      ' Nothing
    Case GetBuildfileValue("SetupSPCU") <> "YES"
      ' Nothing
    Case (strSQLVersion = "SQL2005") And (strEdition = "EXPRESS")
      Call SetBuildfileValue("SetupSPCUStatus", strStatusBypassed)
    Case strSPInclude <> ""
      Call SetBuildfileValue("SetupSPCUStatus", strStatusBypassed & strSPInclude)
    Case Else
      Call SetupCumulativeUpdate()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3BB"
      ' Nothing
    Case GetBuildfileValue("SetupSPCUSNAC") <> "YES"
      ' Nothing
    Case (strSQLVersion = "SQL2008") And (strSPInclude <> "")
      Call SetBuildfileValue("SetupSPCUSNACStatus", strStatusBypassed & strSPInclude)
    Case Else
      Call SetupCumulativeUpdateSNAC()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3CZ"
      ' Nothing
    Case Else
      Call SetupBOL()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3DZ"
      ' Nothing
    Case Else
      Call PostFixTasks()
  End Select

  Call SetUpdate("OFF")
  Call SetProcessId("3Z", "SQL Fixes processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetSQLCompatFlags()
  Call DebugLog("SetSQLCompatFlags:")
  Dim arrProducts
  Dim objProduct
  Dim strPathReg, strProduct

  Select Case True
    Case strOSVersion >= "6.0"
      ' Nothing
    Case Else ' KB2918614 has broken 'minor upgrade' on W2003
      strPathReg    = "Installer\Products"
      objWMIReg.EnumKey strHKCR,strPathReg,arrProducts
      For Each objProduct In arrProducts
        strPath     = "HKCR\" & strPathReg & "\" & objProduct
        strProduct  = objShell.RegRead(strPath & "\ProductName")
        If Left(strProduct, 40) = "Microsoft SQL Server Setup Support Files" Then
          strCMD    = "REG DELETE """ & strPath & """ /f"
          Call Util_RunExec(strCmd, "", strResponseYes, 0)
        End If
      Next
  End Select

End Sub


Sub SetupServicePack()
  Call SetProcessId("3A", "Install SQL Service Pack " & strSPLevel)
  Dim objInstParm
  Dim strParmXtra, strParmRetry, strParmSilent

  If strActionSQLDB <> "ADDNODE" Then
    Call StopSQLServer()
  End If

  Select Case True
    Case strSQLVersion <= "SQL2008"
      strParmXtra   = strInstParm
    Case strIAcceptLicenseTerms = "YES"
      strParmXtra   = strInstParm & " /IAcceptSQLServerLicenseTerms"
  End Select
  
  strParmRetry      = "29534"
  If strOSVersion >= "6.1" Then
    strParmRetry    = strParmRetry & " 5"
  End If

  Select Case True
    Case strMode = "ACTIVE"
      strParmSilent = ""
    Case strSQLVersion = "SQL2005" And strEdition = "EXPRESS"
      strParmSilent = "/QB" 
    Case strSQLVersion = "SQL2005"
      strParmSilent = "/quiet" 
    Case Else
      strParmSilent = "/QS" 
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot",    "YES")
  Call SetXMLParm(objInstParm, "ParmXtra",     strParmXtra)
  Call SetXMLParm(objInstParm, "PathMain",     strSQLSPMedia & strSPLevel)
  Call SetXMLParm(objInstParm, "ParmLog",      "")
  Call SetXMLParm(objInstParm, "ParmReboot",   "")
  Call SetXMLParm(objInstParm, "ParmRetry",    strParmRetry)
  Call SetXMLParm(objInstParm, "ParmSilent",   strParmSilent)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("SP", GetBuildfileValue("SPFile"), objInstParm)

  If GetBuildfileValue("SetupSPStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Performing Post-SP actions")

  Select Case True
    Case strSQLVersion = "SQL2005"
      strCmd            = strHKLMSQL & strInstRegSQL & "\MSSQLServer\BackupDirectory"
      Call Util_RegWrite(strCmd, GetBuildfileValue("DirBackup"), "REG_SZ") 
  End Select

  Call SetSQLLogShortcut (strInstLog & strProcessIdLabel & " " & strProcessIdDesc)

  Call SetBuildfileValue("SetupSPStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupCumulativeUpdate()
  Call SetProcessId("3BA", "Install SQL Cumulative Update " & strSPLevel & " " & strSPCULevel)
  Dim objInstParm
  Dim strCUFile, strParmXtra, strParmRetry, strParmSilent

  If strActionSQLDB <> "ADDNODE" Then
    Call StopSQLServer()
  End If

  strCUFile         = GetBuildfileValue("SPFile")
  If strWOWX86 = "TRUE" Then
    strCUFile       = Replace(strCUFile, "-x64-", "-x86-") ' x86 SQL installed on x64 server
  End If

  Select Case True
    Case strSQLVersion <= "SQL2008"
      strParmXtra   = strInstParm
    Case strIAcceptLicenseTerms = "YES"
      strParmXtra   = strInstParm & " /IAcceptSQLServerLicenseTerms"
  End Select
  
  strParmRetry      = "29534"
  If strOSVersion >= "6.1" Then
    strParmRetry    = strParmRetry & " 5"
  End If

  Select Case True
    Case strMode = "ACTIVE"
      strParmSilent = ""
    Case strSQLVersion = "SQL2005" And strEdition = "EXPRESS"
      strParmSilent = "/QB" 
    Case strSQLVersion = "SQL2005"
      strParmSilent = "/quiet" 
    Case Else
      strParmSilent = "/QS" 
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot",    "YES")
  Call SetXMLParm(objInstParm, "ParmXtra",     strParmXtra)
  Call SetXMLParm(objInstParm, "PathMain",     strSQLSPMedia & strSPLevel)
  Call SetXMLParm(objInstParm, "ParmLog",      "")
  Call SetXMLParm(objInstParm, "ParmReboot",   "")
  Call SetXMLParm(objInstParm, "ParmRetry",    strParmRetry)
  Call SetXMLParm(objInstParm, "ParmSilent",   strParmSilent)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("SPCU", strCUFile, objInstParm)

  If GetBuildfileValue("SetupSPCUStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Performing Post-CU actions")

  Select Case True
    Case strSQLVersion = "SQL2005"
      strCmd            = strHKLMSQL & strInstRegSQL & "\MSSQLServer\BackupDirectory"
      Call Util_RegWrite(strCmd, GetBuildfileValue("DirBackup"), "REG_SZ") 
  End Select

  Call SetSQLLogShortcut (strInstLog & strProcessIdLabel & " " & strProcessIdDesc)

  Call SetBuildfileValue("SetupSPCUStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupCumulativeUpdateSNAC()
  Call SetProcessId("3BB", "Install SQL Cumulative Update " & strSPLevel & " " & strSPCULevel & " for SNAC")
  Dim objInstParm
  Dim strSNACDir, strSNACFile, strPathInst

  strSNACFile       = GetBuildfileValue("SNACFile")
  strPathInst       = GetPathInst(strSnacFile, strSQLSPMedia & strSPLevel & "\", "")
  If strPathInst    = "" Then
    Call SetBuildfileValue("SetupSPCUSNACStatus", strStatusBypassed & ", no media")
    Call FBLog(" " & strProcessIdDesc & strStatusBypassed & ", no media")
    Exit Sub
  End If

  strSNACDir        = strPathTemp & "\SNAC"
  Set objFile       = objFSO.GetFile(strPathInst)
  If Not objFSO.FolderExists(strSNACDir) Then
    objFSO.CreateFolder(strSNACDir)
    Wscript.Sleep strWaitShort
  End If
  strSNACFile       = "sqlncli." & Right(strPathInst, 3)
  objFile.Copy strSNACDir & "\" & strSNACFile

  Call SetXMLParm(objInstParm, "PathMain",     strSNACDir)
  Call RunInstall("SPCUSNAC",  strSNACFile,    objInstParm)

End Sub


Sub SetupBOL()
  Call SetProcessId("3C", "BOL Update processing")

  Dim strSetupBOL
  strSetupBOL       = GetBuildfileValue("SetupBOL")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3CA"
      ' Nothing
    Case strSetupBOL <> "YES"
      ' Nothing
    Case strSQLVersion >= "SQL2012"
      ' Nothing
    Case Else
      Call SetupBOLUpdate()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3CB"
      ' Nothing
    Case strSetupBOL <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strSQLVersion >= "SQL2016"
      ' Nothing
    Case Else
      Call SetupBOLViewer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3CC"
      ' Nothing
    Case strSetupBOL <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case Else
'      Call SetupHelpCatalog()
  End Select

  Call SetProcessId("3CZ", "BOL Update processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupBOLUpdate()
  Call SetProcessId("3CA", "Install SQL Books Online update")
  Dim objInstParm

  Call RunInstall("BOL", GetBuildfileValue("BOLmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupBOLViewer()
  Call SetProcessId("3CB", "Install SQL Books Online update")
  Dim strBOLdir, strBolexe, strHelpexe, strHelpPath, strRegBase

  strBOLdir         = strPathTemp & "SQLBOL"
  strBolexe         = GetBuildfileValue("BOLexe")
  strPath           = strPathAddComp & strBolexe 
  Select Case True
    Case objFSO.FileExists(strPath)
      Call SetTrustedZone(strPath)
    Case Else
      Call SetBuildfileValue("SetupBOLStatus", strStatusBypassed)
      Call FBLog(" " & strProcessIdDesc & strStatusBypassed)
      Exit Sub
  End Select

  Call DebugLog("Extract BOL files")
  strRegBase        = "HKLM\SOFTWARE\Microsoft\Help\v1.0\"
  strCmd            = """" & strPath & """ /Auto """ & strBOLdir & "\"""
  Call Util_RunExec(strCmd, "", "", 0)

  Call DebugLog("Build Help library")
  strPath           = strRegBase & "LocalStore"
  strHelpPath       = objShell.RegRead(strPath)
  If Not objFSO.FolderExists(strHelpPath) Then
    objFSO.CreateFolder(strHelpPath)
  End If
  strPath           = strRegBase & "AppRoot"
  strHelpexe        = objShell.RegRead(strPath) & "HelpLibManager.exe"
  strPath           = strBOLdir & "\" & Left(strBolexe, InstrRev(strBolexe, ".") - 1)
  If Not objFSO.FolderExists(strPath) Then
    strPath         = strBOLdir
  End If
  strCmd            = " /product SQLSERVER /version " & strSQLVersionNum & " /locale en-us /silent /sourceMedia " & strPath & "\helpcontentsetup.msha "
  Call Util_RunExec("""" & strHelpexe & """" & strCmd, "", "", 0)

  Call DebugLog("Remove temporary install media folder")
  Set objFolder     = objFSO.GetFolder(strBOLdir)
  objFolder.Delete(1)

  Call SetBuildfileValue("SetupBOLStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupHelpCatalog()
  Call SetProcessId("3CB", "Setup SQL Books Online catalog")
' Code based on http://ariely.info/Blog/tabid/83/EntryId/176/Installing-Microsoft-SQL-Server-2016-Book-online-for-offline-use.aspx
  Dim strRegBase

  Call DebugLog("Build Help library")
  strRegBase        = "HKLM\SOFTWARE\Microsoft\Help\v2.2\"
  strPath           = strRegBase & "Catalogs\SSMS16\LocationPath"

' "C:\Program Files (x86)\Microsoft Help Viewer\v2.2\HlpViewer.exe" /catalogName SSMS16

  Call SetBuildfileValue("SetupBOLStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub PostFixTasks()
  Call SetProcessId("3D", "Post Fix Tasks")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3DA"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strType = "CLIENT"
      ' Nothing
    Case Else
      Call RunUpdates()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3DB"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strType = "CLIENT"
      ' Nothing
    Case Else
      Call RunBackup()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3DC"
      ' Nothing
    Case strType = "CLIENT"
      Call GetSQLVersion()
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call GetSQLVersion()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "3DD"
      ' Nothing
    Case strType <> "FIX"
      ' Nothing
    Case Else
      Call CloseFixFrocess()
  End Select

  Call SetProcessId("3DZ", "Post Fix Tasks" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub RunUpdates()
  Call SetProcessId("3DA", "Run Update Scripts")

  Call StartSQL()
  Wscript.Sleep strWaitLong ' Wait to allow time for update scripts to run
  Call StartSSAS()

  Call ProcessEnd(strStatusComplete)

End Sub


Sub RunBackup()
  Call SetProcessId("3DB", "Copy System Database Files")

  Call StopSQLServer()

  Call DebugLog("Backing up System Database files")
  strPathLog        = GetPathLog("")
  strCmd            = "%COMSPEC% /D /C CSCRIPT.EXE """ & strPathFBScripts & "SqlSysDBCopy.vbs"" >> " & strPathLog
  Call Util_RunExec(strCmd, "", "", 0)

  If strActionSQLDB <> "ADDNODE" Then
    Call StartSQL()
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub GetSQLVersion()
  Call SetProcessId("3DC", "Get SQL patch level")

  strSQLVersionFull = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\Version")
  Call SetBuildfileValue("SQLVersionFull", strSQLVersionFull)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CloseFixFrocess()
  Call SetProcessId("3DD", "Close Fix Frocess")

  Call SetBuildfileValue("FineBuildStatus", strstatusComplete)

  Call ProcessEnd(strStatusComplete)

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


Sub StopSQLServer()
  Call DebugLog("StopSQLServer:")

  Call StopSQL()

  strCmd            = """" & strPathFBScripts & "SqlServiceStop.bat"""
  If strInstance <> "MSSQLSERVER" Then
    strCmd          = strCmd & " " & strInstance
  End If
  Call Util_RunExec(strCmd, "", "", 2)

End Sub


Sub SetSQLLogShortcut(strDescription)
  Call DebugLog("SetSQLLogShortcut:" & strDescription)
  Dim objLogFolder, objShortcut
  Dim strSQLLog, strPathDest

  strPath           = objShell.RegRead(strHKLMSQL & strSQLVersionNum & "\Bootstrap\BootstrapDir") & "Log"
  Set objFolder     = objFSO.GetFolder(strPath)
  Set colFolders    = objFolder.Subfolders

  Call DebugLog("Find the most recent SQL install log folder")
  strSQLLog         = ""
  For Each objLogFolder In colFolders
    If objLogFolder.name > strSQLLog Then
      strSQLLog     = objLogFolder.name
    End If
  Next

  Select Case True
    Case strSQLLog = "" 
      strSQLLog     = "Files"
    Case Else
      Call DebugLog("Copy SQL Support log to main SQL log folder")
      strPathDest   = strPath & "\" & strSQLLog & "\"
      strPathLog    = strPathTemp & "SqlSetup.log"
      If objFSO.FileExists(strPathLog) Then
        Call objFSO.CopyFile(strPathLog, strPathDest, True)
        Call objFSO.DeleteFile(strPathLog, True)
      End If
      strPathLog    = strPathTemp & "SqlSetup_Local.log"
      If objFSO.FileExists(strPathLog) Then
        Call objFSO.CopyFile(strPathLog, strPathDest, True)
        Call objFSO.DeleteFile(strPathLog, True)
      End If
  End Select

  Call FBLog(" " & strDescription & " log files located in: " & strPath & "\" & strSQLLog)
  Set objShortcut   = objShell.CreateShortcut(Mid(strSetupLog, 2) & strDescription & ".lnk")
  objShortcut.TargetPath  = strPath & "\" & strSQLLog
  objShortcut.Description = strSQLLog
  objShortcut.Save()

End Sub


End Class