''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild4InstallXtras.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      SQL Server Extra Components Install 
'
'  Author:       Ed Vassie
'
'  Date:         02 Jul 2008
'
'  Change History
'  Version  Author        Date         Description
'  2.1      Ed Vassie     18 Jun 2010  Initial SQL Server 2008 R2 version
'  2.0      Ed Vassie     02 Jul 2008  Initial SQL Server 2008 version for FineBuild v2.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colPrcEnvVars
Dim intIndex
Dim objAutoUpdate, objFile, objFolder, objFSO, objShell, objWMI, objWMIReg
Dim strAction, strActionAO, strActionSQLAS, strActionSQLDB, strActionSQLRS, strAdminPassword, strAGDagName, strAGPrimary, strAllUserProf, strAltFile, strAnyKey
Dim strCatalogInstance, strCatalogServerName, strClusterAction, strClusterHost, strClusterName, strClusterGroupDTC, strClusterNameSQL, strClusterGroupSQL, strCmd, strCmdPS, strCmdRS, strCmdSQL, strCompatFlags
Dim strDirBackup, strDirData, strDirDBA, strDirLog, strDirProg, strDirProgX86, strDirProgSys, strDirProgSysX86, strDirSystemDataPrimary, strDirSystemDataBackup, strDirSystemDataShared, strDirSQL, strDirSys, strDirSysData, strDirSysWOW
Dim strDomain, strDrive, strDQSInstall
Dim strEdition, strEditionEnt, strFileArc, strGroupAO, strGroupDBA, strGroupDBAAlt, strGroupDBANonSa, strGroupDBANonSaAlt, strGroupDistComUsers, strGroupIISIUsers, strIISRoot, strInstLog, strInstRegRS, strInstRegSQL, strInstRS, strInstRSDir, strInstRSHost, strInstRSSQL, strInstRSURL
Dim strHKLMSQL, strHTTP, strInstance, strInstNode, strInstParm, strInstSQL, strLocalAdmin
Dim strMailServer, strMDSAccount, strMenuAdminTools, strMenuConfigTools, strMenuPerfTools, strMenuPrograms, strMenuSQL, strMenuSQL2005, strMenuSQLAS, strMenuSQLDocs, strMenuSQLIS, strMenuSQLNS, strMenuSQLRS
Dim strMode, strNTAuthEveryone, strOSLanguage, strOSName, strOSType, strOSVersion
Dim strPath, strPathAddComp, strPathAlt, strPathBOL, strPathCmdSQL, strPathExe, strPathFB, strPathFBScripts, strPathLog, strPathMDS, strPathNew, strPathNLS, strPathOld, strPathSSRS, strPathTemp, strPathSSMS, strPathSSMSX86, strPathVS, strPID, strProcArc, strPSInstall
Dim strReboot, strRsAccount, strRSAlias, strRsPassword, strRSDBName, strRsExecAccount, strRsExecPassword, strRSInstallMode, strRSURLSuffix, strsaPwd, strServer, strServerAO, strServInst, strSetupLog
Dim strSetupAOAlias, strSetupAPCluster, strSetupIIS, strSetupPowerBI, strSetupRSDB, strSetupSQLAS, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLRSCluster, strSetupSQLIS, strSetupSQLTools
Dim strSetupSQLRS, strPathSQLMedia, strSQLAccount, strSQLBinRoot, strSqlBrowserStartup, strSQLLanguage, strSQLProgDir, strSQLRSExe, strSQLVersion, strSQLVersionNum, strSQLVersionFull, strSQLVersionWMI
Dim strVersionNet3, strVolSys, strVSVersionNum, strStopAt, strTCPPortAO, strTCPPortRS, strType, strUserDNSDomain, strUserName, strUserDTop, strUserProf, strWaitLong, strWaitMed, strWaitShort, strWOWX86


Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "4Z"
      ' Nothing
    Case Else
      Call ProcessXtras()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "4Z"
      ' Nothing
    Case err.Number = 0 
      Call  objShell.Popup("SQL Server Xtra Components Install complete", 2, "SQL Server Xtra Components Install" ,64)
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
      Call FBLog(" Xtra Components Install failed")
  End Select

  Wscript.Quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Call SetProcessIdCode("FB4X")
  Include "FBUtils.vbs"
  Include "FBManageAccount.vbs"
  Include "FBManageBoot.vbs"
  Include "FBManageCluster.vbs"
  Include "FBManageInstall.vbs"
  Include "FBManageRSWMI.vbs"
  Include "FBManageService.vbs"

  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")

  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAction         = GetBuildfileValue("Action")
  strActionAO       = GetBuildfileValue("ActionAO")
  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strActionSQLRS    = GetBuildfileValue("ActionSQLRS")
  strAdminPassword  = GetBuildfileValue("AdminPassword")
  strAGDagName      = GetBuildfileValue("AGDagName")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strAllUserProf    = GetBuildfileValue("AllUserProf")
  strCatalogInstance   = GetBuildfileValue("CatalogInstance")
  strCatalogServerName = GetBuildfileValue("CatalogServerName")
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterGroupDTC  = GetBuildfileValue("ClusterGroupDTC")
  strClusterGroupSQL  = GetBuildfileValue("ClusterGroupSQL")
  strClusterNameSQL = GetBuildfileValue("ClusterNameSQL")
  strCmdPS          = GetBuildfileValue("CmdPS")
  strCmdRS          = GetBuildfileValue("CmdRS")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strCompatFlags    = GetBuildfileValue("CompatFlags")
  strDirData        = GetBuildfileValue("DirData")
  strDirDBA         = GetBuildfileValue("DirDBA")
  strDirLog         = GetBuildfileValue("DirLog")
  strDirProg        = GetBuildfileValue("DirProg")
  strDirProgX86     = GetBuildfileValue("DirProgX86")
  strDirProgSys     = GetBuildfileValue("DirProgSys")
  strDirProgSysX86  = GetBuildfileValue("DirProgSysX86")
  strDirSys         = GetBuildfileValue("DirSys")
  strDirSysData     = GetBuildfileValue("DirSysData")
  strDirSystemDataPrimary  = GetBuildfileValue("DirSystemDataPrimary")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup")
  strDirSystemDataShared = GetBuildfileValue("DirSystemDataShared")
  strDirSysWOW      = GetBuildfileValue("DirSysWOW")
  strDomain         = GetBuildfileValue("Domain")
  strEdition        = GetBuildfileValue("AuditEdition")
  strEditionEnt     = GetBuildfileValue("EditionEnt")
  strFileArc        = GetBuildfileValue("FileArc")
  strGroupAO        = GetBuildfileValue("GroupAO")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBAAlt    = GetBuildfileValue("GroupDBAAlt")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strGroupDBANonSAAlt  = GetBuildfileValue("GroupDBANonSAAlt")
  strGroupDistComUsers = GetBuildfileValue("GroupDistComUsers")
  strGroupIISIUsers = GetBuildfileValue("GroupIISIUsers")
  strHTTP           = GetBuildfileValue("HTTP")
  strIISRoot        = GetBuildfileValue("IISRoot")
  strInstance       = GetBuildfileValue("Instance")
  strInstLog        = GetBuildfileValue("InstLog")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strInstRegRS      = GetBuildFileValue("InstRegRS")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstRSDir      = GetBuildfileValue("InstRSDir")
  strInstRSHost     = GetBuildfileValue("InstRSHost")
  strInstRSSQL      = GetBuildfileValue("InstRSSQL")
  strInstRSURL      = GetBuildfileValue("InstRSURL")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strMailServer     = GetBuildfileValue("MailServer")
  strMDSAccount     = GetBuildfileValue("MDSAccount")
  strMenuAdminTools = GetBuildfileValue("MenuAdminTools")
  strMenuConfigTools  = GetBuildfileValue("MenuConfigTools")
  strMenuPerfTools  = GetBuildfileValue("MenuPerfTools")
  strMenuPrograms   = GetBuildfileValue("MenuPrograms")
  strMenuSQL        = GetBuildfileValue("MenuSQL")
  strMenuSQL2005    = GetBuildfileValue("MenuSQL2005")
  strMenuSQLAS      = GetBuildfileValue("MenuSQLAS")
  strMenuSQLDocs    = GetBuildfileValue("MenuSQLDocs")
  strMenuSQLIS      = GetBuildfileValue("MenuSQLIS")
  strMenuSQLNS      = GetBuildfileValue("MenuSQLNS")
  strMenuSQLRS      = GetBuildfileValue("MenuSQLRS")
  strMode           = GetBuildfileValue("Mode")
  strNTAuthEveryone = GetBuildfileValue("NTAuthEveryone")
  strOSLanguage     = GetBuildfileValue("OSLanguage")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathAddComp    = FormatFolder("PathAddComp")
  strPathBOL        = GetBuildfileValue("PathBOL")
  strPathCmdSQL     = GetBuildfileValue("PathCmdSQL")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  strPathMDS        = GetBuildfileValue("PathMDS")
  strPathSQLMedia   = FormatFolder("PathSQLMedia")
  strPathSSMS       = GetBuildfileValue("PathSSMS")
  strPathSSMSx86    = GetBuildfileValue("PathSSMSX86")
  strPathSSRS       = GetBuildfileValue("PathSSRS")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPathVS         = GetBuildfileValue("PathVS")
  strPID            = GetBuildfileValue("PID")
  strProcArc        = GetBuildfileValue("ProcArc")
  strPSInstall      = GetBuildfileValue("PSInstall")
  strReboot         = GetBuildfileValue("RebootStatus")
  strRsAccount      = GetBuildfileValue("RsAccount")
  strRSAlias        = GetBuildfileValue("RSAlias")
  strRsPassword     = GetBuildfileValue("RsPassword")
  strRSDBName       = GetBuildfileValue("RSDBName")
  strRsExecAccount  = GetBuildfileValue("RsExecAccount")
  strRsExecPassword = GetBuildfileValue("RsExecPassword")
  strRSInstallMode  = GetBuildfileValue("RSInstallMode")
  strRSURLSuffix    = GetBuildfileValue("RSURLSuffix")
  strsaPwd          = GetBuildfileValue("saPwd")
  strServer         = GetBuildfileValue("AuditServer")
  strServerAO       = GetBuildfileValue("ServerAO")
  strServInst       = GetBuildfileValue("ServInst")
  strSetupAOAlias   = GetBuildfileValue("SetupAOAlias")
  strSetupAPCluster = GetBuildfileValue("SetupAPCluster")
  strSetupIIS       = GetBuildfileValue("SetupIIS")
  strSetupLog       = Ucase(objShell.ExpandEnvironmentStrings("%SQLLOGTXT%"))
  strSetupLog       = Left(strSetupLog, InStrRev(strSetupLog, "\"))
  strSetupPowerBI   = GetBuildfileValue("SetupPowerBI")
  strSetupRSDB      = GetBuildfileValue("SetupRSDB")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSQLAccount     = GetBuildfileValue("SqlAccount")
  strSQLBinRoot     = GetBuildfileValue("SQLBinRoot")
  strSqlBrowserStartup  = GetBuildfileValue("SqlBrowserStartup")
  strSQLLanguage    = GetBuildfileValue("SQLLanguage")
  strSQLProgDir     = GetBuildfileValue("SQLProgDir")
  strSQLRSExe       = GetBuildfileValue("SQLRSexe")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strSQLVersionFull = GetBuildfileValue("SQLVersionFull")
  strSQLVersionWMI  = GetBuildfileValue("SQLVersionWMI")
  strStopAt         = GetBuildfileValue("StopAt")
  strTCPPortAO      = GetBuildfileValue("TCPPortAO")
  strTCPPortRS      = GetBuildfileValue("TCPPortRS")
  strType           = GetBuildfileValue("Type")
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  strUserDTop       = GetBuildfileValue("UserDTop")
  strUserName       = GetBuildfileValue("AuditUser")
  strUserProf       = GetBuildfileValue("UserProf")
  strVolSys         = GetBuildfileValue("VolSys")
  strVSVersionNum   = GetBuildfileValue("VSVersionNum")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitMed        = GetBuildfileValue("WaitMed")
  strWaitShort      = GetBuildfileValue("WaitShort")
  strWOWX86         = GetBuildfileValue("WOWX86")
  strPath           = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\"
  objWMIReg.GetStringValue strHKLM,strPath,"Version",strVersionNet3

  strAGPrimary      = GetStatefileValue(strGroupAO)

  If Right(strPathTemp, 1) <> "\" Then
    strPathTemp     = strPathTemp & "\"
  End If

End Sub


Sub ProcessXtras()
  Call SetProcessId("4", strSQLVersion & " Xtras processing (FineBuild4InstallXtras.vbs)")

  Call SetUpdate("ON")
  strReboot         = GetBuildfileValue("RebootStatus")

  Select Case True
    Case strType = "CLIENT" 
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call StartSQL()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AZ"
      ' Nothing
    Case Else
      Call SetupPreReqXtras()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4BZ"
      ' Nothing
    Case Else
      Call SetupBIXtras()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4IZ"
      ' Nothing
    Case Else
      Call SetupISXtras()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RZ"
      ' Nothing
    Case Else
      Call SetupReportXtras()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SZ"
      ' Nothing
    Case Else
      Call SetupSQLXtras()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TZ"
      ' Nothing
    Case Else
      Call SetupToolsXtras()
  End Select

  If strSetupSQLTools = "YES" Then
    Call SetBuildfileValue("SetupSQLToolsStatus", strStatusComplete)
  End If
  Call SetUpdate("OFF")

  Call CheckReboot()
  strReboot         = GetBuildfileValue("RebootStatus")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4X"
      ' Nothing
    Case (strReboot = "Pending") Or (strProcessId = "4X")
      Call SetupReboot("4Y", "End of Xtras")
    Case Else
      ' Nothing
  End Select

  Call SetProcessId("4Z", "Xtras processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupPreReqXtras()
  Call SetProcessId("4A", "Pre-Requisite Components")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AA"
      ' Nothing
    Case GetBuildfileValue("SetupSQLPowershell") <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLPowershell()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ABZ"
      ' Nothing
    Case strSQLVersion > "SQL2008R2"
      ' Nothing
    Case Else
      Call CheckVS2005Fixes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ACZ"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strSQLVersion > "SQL2016"
      ' Nothing
    Case Else
      Call CheckVS2010Fixes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strprocessId > "4AD"
      ' Nothing
    Case GetBuildfileValue("SetupMBCA") <> "YES"
      ' Nothing
    Case Else
      Call SetupMBCA()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AE"
      ' Nothing
    Case GetBuildfileValue("SetupReportViewer") <> "YES"
      ' Nothing
    Case Else
      Call SetupReportViewer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AF"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case GetBuildfileValue("SetupSQLBC") <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLBC()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AGZ"
      ' Nothing
    Case GetBuildfileValue("SetupSQLCE") <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLCE()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AH"
      ' Nothing
    Case GetBuildfileValue("SetupSSMS") <> "YES"
      ' Nothing
    Case GetBuildfileValue("UseFreeSSMS") <> "YES"
      ' Nothing
    Case Else
      Call SetupSSMS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AI"
      ' Nothing
    Case GetBuildfileValue("SetupKB2854082") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2854082()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AJ"
      ' Nothing
    Case GetBuildfileValue("SetupKB2862966") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2862966()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4AK"
      ' Nothing
    Case GetBuildfileValue("SetupVC2010") <> "YES"
      ' Nothing
    Case Else
      Call SetupVC2010()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALZ"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case Else
      Call SetupAlwaysOn()
  End Select

  Call SetProcessId("4AZ", "Pre-Requisite Components" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSQLPowershell()
  Call SetProcessId("4AA", "Setup SQL Powershell")
  Dim colFolders
  Dim objFolder
  Dim strPathPSInst, strPathPS

  Call SetBuildfileValue("SetupSQLPowershellStatus", strStatusProgress)
  strPathPS         = GetBuildfileValue("PathPS")

  strPathPSInst     = strPathAddComp & "WindowsPowerShell"
  If objFSO.FolderExists(strPathPSInst) Then
    Set colFolders  = objFSO.GetFolder(strPathPSInst).SubFolders
    For Each objFolder In colFolders
      Call SetupPSModule(strPathPS, strPathPSInst, objFolder.Name)
    Next
  End If

  Call SetBuildfileValue("SetupSQLPowershellStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPSModule(strPathPS, strPathPSInst, strPSName)
  Call DebugLog("SetupPSModule: " & strPSName)

  Call SetBuildfileValue("SetupSQLPowershell" & strPSName & "Status", strStatusProgress)

  Set objFolder     = objFSO.GetFolder(strPathPSInst & "\" & strPSName)
  strDebugMsg1      = "Source: " & strPathPSInst & "\" & strPSName
  strDebugMsg2      = "Target: " & strPathPS
  objFolder.Copy strPathPS & "\" & strPSName
  Call SetBuildfileValue("SetupSQLPowershell" & UCase(strPSName) & "Status", strStatusComplete)

  Call SetBuildfileValue("SQLPowershellList", GetBuildfileValue("SQLPowershellList") & strPSName & " ")

End Sub


Sub CheckVS2005Fixes()
  Call SetProcessId("4AB","Check VS2005 Fixes")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ABA"
      ' Nothing
    Case GetBuildfileValue("SetupVS2005SP1") <> "YES"
      ' Nothing
    Case Else
      Call SetupVS2005SP1()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ABB"
      ' Nothing
    Case GetBuildfileValue("SetupKB932232") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB932232()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ABC"
      ' Nothing
    Case GetBuildfileValue("SetupKB954961") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB954961()
  End Select

  Call SetProcessId("4ABZ", " Check VS2005 Fixes" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupVS2005SP1()
  Call SetProcessId("4ABA", "Installing VS 2005 SP1")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmSilent", "/q")
  Call SetXMLParm(objInstParm, "ParmLog",    "/log")
  Call RunInstall("VS2005SP1", GetBuildfileValue("VS2005SP1exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB932232()
  Call SetProcessId("4ABB", "Installing VS 2005 SP1 fix for compatibiity with Vista and above KB932232")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmSilent", "/q")
  Call SetXMLParm(objInstParm, "ParmLog",    "/log")
  Call RunInstall("KB932232", GetBuildfileValue("KB932232exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB954961()
  Call SetProcessId("4ABC", "Installing VS 2005 SP1 fix for compatibiity with SQL 2008 and above KB954961")
  Dim objInstParm

'  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB954961~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.1.6001.18242")
'  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB954961", GetBuildfileValue("KB954961exe"), "")

  Call ProcessEnd("")

End Sub


Sub CheckVS2010Fixes()
  Call SetProcessId("4AC","Check VS2010 Fixes")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ACA"
      ' Nothing
    Case GetBuildfileValue("SetupVS2010SP1") <> "YES"
      ' Nothing
    Case Else
      Call SetupVS2010SP1()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ACB"
      ' Nothing
    Case GetBuildfileValue("SetupKB2549864") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2549864()
  End Select

  Call SetProcessId("4ACZ", " Check VS2010 Fixes" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupVS2010SP1()
  Call SetProcessId("4ACA", "Installing VS 2010 SP1")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",   strPathAddComp & "VS2010SP1\")
  Call RunInstall("VS2010SP1", GetBuildfileValue("VS2010SP1exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB2549864()
  Call SetProcessId("4ACB", "Installing VS 2010 SP1 KB2549864 fix")
  Dim objInstParm

'  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB2549864~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.1.6001.18242")
'  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB2549864", GetBuildfileValue("KB2549864exe"), "")

  Call ProcessEnd("")

End Sub


Sub SetupMBCA()
  Call SetProcessId("4AD", "Install Baseline Configuration Analyzer")

  Call RunInstall("MBCA", GetBuildfileValue("MBCAmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupReportViewer()
  Call SetProcessId("4AE", "Install Report Viewer")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConType",  "FILE")
  Call SetXMLParm(objInstParm, "PreConValue", GetBuildfileValue("ReportViewerVersion"))
  Call SetXMLParm(objInstParm, "ParmLog",     "/l")
  Call SetXMLParm(objInstParm, "ParmRetry",   "1618")
  Call SetXMLParm(objInstParm, "ParmSilent",  "/q")
  Call RunInstall("ReportViewer", GetBuildfileValue("ReportViewerexe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSQLBC()
  Call SetProcessId("4AF", "Install SQL Backwards Compatibility")
  Dim objInstParm

  Select Case True
    Case strSQLVersion = "SQL2008R2"
      Call SetXMLParm(objInstParm, "PathAlt", strPathSQLMedia & "1033_ENU_LP\x86\Setup\x86\")
    Case Else
      Call SetXMLParm(objInstParm, "PathAlt", strPathSQLMedia & "x86\Setup\x86\" )
  End Select
  Call RunInstall("SQLBC", GetBuildfileValue("SQLBCmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSQLCE()
  Call SetProcessId("4AG", "Install SQL Compact Edition")

  Select Case True
    Case strSQLVersion <= "SQL2008R2"
      Call SetBuildfileValue("SetupSQLCEStatus", strStatusBypassed)
    Case strSQLVersion >= "SQL2012"
      Call SetupSQLCE40()
  End Select

  Call SetProcessId("4AGZ", " Install SQL Compact Edition" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSQLCE40()
  Call SetProcessId("4AGA", "Install SQL Compact Edition 4.0")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "SetupOption", "Extract")
  Call SetXMLParm(objInstParm, "InstFile",    "SSCERuntime_" & strFileArc & "-" & strOSLanguage & ".msi")
  Call RunInstall("SQLCE", GetBuildfileValue("SQLCEexe"), objInstParm)

End Sub


Sub SetupSSMS()
  Call SetProcessId("4AH", "Installing SSMS")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("SSMS", GetBuildfileValue("SSMSexe"), objInstParm)

  If GetBuildfileValue("SetupSSMSStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Save path to SSMS")
  strPath           = GetBuildfileValue("RegTools") & "SQLPath"
  strPathSSMS       = objShell.RegRead(strPath) & "\"
  Call SetBuildfileValue("PathSSMS", strPathSSMS)

  Call SetBuildfileValue("SetupSSMSStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupKB2854082()
  Call SetProcessId("4AI", "Installing KB2854082")
  Dim objInstParm

  Call RunInstall("KB2854082", GetBuildfileValue("KB2854082File"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB2862966()
  Call SetProcessId("4AJ", "Installing KB2862966")
  Dim objInstParm

  Call RunInstall("KB2862966", "windows8-rt-kb2854082-x64_f6f19ddae1dd7d15b21aee336ac01de9becb41c6.msu", objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupVC2010()
  Call SetProcessId("4AK", "Installing VC2010")
  Dim objInstParm

  Call RunInstall("VC2010", "vcredist_x86.exe", objInstParm)

  If strFileArc = "X64" Then
    Call RunInstall("VC2010X64", "vcredist_x64.exe", objInstParm)
  End If

  Call ProcessEnd("")

End Sub


Sub SetupAlwaysOn()
  Call SetProcessId("4AL", "Setup Always On")
  Dim strSetupAlwaysOn

  strSetupAlwaysOn  = GetBuildfileValue("SetupAlwaysOn")
  If strSetupAlwaysOn = "YES" Then
    Call SetBuildfileValue("SetupAlwaysOnStatus", strStatusProgress)
  End If

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALA"
      ' Nothing
    Case strType = "CLIENT" 
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case Else
      Call EnableAlwaysOn()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALB"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else
      Call EnableAOEndpoint()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALC"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else
      Call CreateAOGroup()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALD"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else
      Call ConfigureAOGroup()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ALE"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case GetBuildfileValue("ActionDAG") = ""
      ' Nothing
    Case Else
      Call ConfigureDag()
  End Select

  If strSetupAlwaysOn = "YES" Then
    Call SetBuildfileValue("SetupAlwaysOnStatus", strStatusComplete)
  End If
  Call SetProcessId("4ALZ", " Setup Always On" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub EnableAlwaysOn()
  Call SetProcessId("4ALA", "Enable Always On Pre-requisites")
  Dim objSQLManagement, objHADRService

'  Select Case true
'    Case strActionSQLDB = "ADDNODE"
'      strPath       = strHKLMSQL & strInstRegSQL & "\MSSQLServer\HADR\"
'      Call Util_RegWrite(strPath & "HADR_Enabled", "1", "REG_DWORD")
'    Case Else
      Call StartSQL()
      Set objSQLManagement = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\SqlServer\ComputerManagement" & strSQLVersionWMI)
      Set objHADRService   = objSQLManagement.Get("HADRServiceSettings.CanSetHADRService=True,InstanceName=""" & strInstSQL & """")
      objHADRService.ChangeHADRService True
      Call StopSQLServer()
      Call StartSQL()
'  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub EnableAOEndpoint()
  Call SetProcessId("4ALB", "Enable Mirroring Endpoint")

  strCmd            = "CREATE ENDPOINT [" & strServerAO & "] AS TCP (LISTENER_PORT=" & strTCPPortAO & ", LISTENER_IP=ALL) "
  strCmd            = strCmd & "FOR DATA_MIRRORING (ROLE=ALL, AUTHENTICATION=WINDOWS NEGOTIATE, ENCRYPTION=REQUIRED ALGORITHM " & GetBuildfileValue("EncryptAO") & ")"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)
  WScript.Sleep strWaitMed
  strCmd            = "ALTER  ENDPOINT [" & strServerAO & "] STATE=STARTED"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 0)
  strCmd            = "ALTER AUTHORIZATION ON ENDPOINT::" & strServerAO & " TO [sa]"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 0)
  strCmd            = "GRANT CONNECT ON ENDPOINT::" & strServerAO & " TO [" & strSqlAccount & "]"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 0)

  WScript.Sleep strWaitLong
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreateAOGroup()
  Call SetProcessId("4ALC", "Create Database Availability Group")

  Select Case True
    Case strActionAO = "ADDNODE"
      Call CreateSecondaryAO()
    Case Else 
      Call CreatePrimaryAO()
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreatePrimaryAO()
  Call DebugLog("CreatePrimaryAO:")

  Call Util_RunExec("REG DELETE ""HKEY_LOCAL_MACHINE\CLUSTER\HadrAgNameToldMap\" & strGroupAO & """ /f",   "", strResponseYes, -1) ' To avoid possible 41042 error

  strCmd            = "CREATE AVAILABILITY GROUP [" & strGroupAO & "]"
  If strClusterHost <> "YES" Then
    strCmd          = strCmd & " WITH (CLUSTER_TYPE = NONE)"
  End If
  strCmd            = strCmd & " FOR REPLICA ON "
  strCmd            = strCmd & " '" & strServerAO & "' WITH (ENDPOINT_URL = 'TCP://" & strServerAO & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
  Select Case True
    Case strEditionEnt = "YES"
      strCmd        = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT"
    Case Else
      strCmd        = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = SYNCHRONOUS_COMMIT"
  End Select
  strCmd            = strCmd & ");"
  Call Util_ExecSQL(strCmdSQL & "-V 11 -Q", """" & strCmd & """", -1)
  Wscript.Sleep strWaitLong

' Experience is that first AG created after enabling Always On is not stable, but deleting and re-creating it does leave it stable
  Call Util_ExecSQL(strCmdSQL & "-V 11 -Q", """DROP AVAILABILITY GROUP [" & strGroupAO & "]""", -1)
  Call Util_ExecSQL(strCmdSQL & "-V 11 -Q", """" & strCmd & """", 1)
  WScript.Sleep strWaitLong

End Sub


Sub CreateSecondaryAO()
  Call DebugLog("CreateSecondaryAO:")

  Select Case True
    Case strSetupSQLDBCluster = "YES" 
      Call AddChildNode("AO", strGroupAO)
    Case Else
      strCmd        = "CREATE AVAILABILITY GROUP [" & strGroupAO & "]"
      If strClusterHost <> "YES" Then
        strCmd      = strCmd & " WITH (CLUSTER_TYPE = NONE)"
      End If
      strCmd        = strCmd & " FOR REPLICA ON "
      strCmd        = strCmd & " '" & strAGPrimary & "' WITH (ENDPOINT_URL = 'TCP://" & strAGPrimary & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
      Select Case True
        Case strEditionEnt = "YES"
          strCmd    = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT"
        Case Else
          strCmd    = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = SYNCHRONOUS_COMMIT"
      End Select
      strCmd        = strCmd & ")"
      strCmd        = strCmd & ",'" & strServerAO  & "' WITH (ENDPOINT_URL = 'TCP://" & strServerAO  & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
      Select Case True
        Case strEditionEnt = "YES"
          strCmd    = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT"
        Case Else
          strCmd    = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = SYNCHRONOUS_COMMIT"
      End Select
      strCmd        = strCmd & ");"
      Call Util_ExecSQL(strCmdSQL & "-V 11 -Q", """" & strCmd & """", -1)
  End Select

End Sub


Sub ConfigureAOGroup()
  Call SetProcessId("4ALD", "Configure Database Availability Group")

  Call DebugLog("AG Primary Node: " & strAGPrimary)
  Select Case True
    Case strActionAO = "ADDNODE"
      Call ConfigureSecondaryAO()
    Case (strAGPrimary <> "") And (strAGPrimary <> strServInst)
      Call ConfigureSecondaryAO()
    Case Else 
      Call ConfigurePrimaryAO()
  End Select

  Call ConfigureAOXtras()

  Call ConfigureLinkedServers(strServer, strAGPrimary, strGroupAO)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigurePrimaryAO()
  Call DebugLog("ConfigurePrimaryAO:")

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] SET ("
  Select Case True
    Case strEditionEnt <> "YES" 
      strCmd        = strCmd & "AUTOMATED_BACKUP_PREFERENCE = PRIMARY"
    Case strSetupAPCluster = "YES"
      strCmd        = strCmd & "AUTOMATED_BACKUP_PREFERENCE = PRIMARY"
    Case Else
      strCmd        = strCmd & "AUTOMATED_BACKUP_PREFERENCE = SECONDARY"
  End Select
  strCmd            = strCmd & ");"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  If strSQLVersion >= "SQL2016" Then
    strCmd          = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] SET (DTC_SUPPORT=PER_DB);"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)
  End If

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] MODIFY REPLICA ON '" & strServerAO & "' "
  strCmd            = strCmd & "WITH (SESSION_TIMEOUT = 10);"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] MODIFY REPLICA ON '" & strServerAO & "' "
  strCmd            = strCmd & "WITH (BACKUP_PRIORITY = 50);"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] MODIFY REPLICA ON '" & strServerAO & "' "
  strCmd            = strCmd & "WITH (PRIMARY_ROLE(ALLOW_CONNECTIONS = ALL));"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] MODIFY REPLICA ON '" & strServerAO & "' "
  strCmd            = strCmd & "WITH ("
  Select Case True
    Case strEditionEnt = "YES" 
      strCmd        = strCmd & "SECONDARY_ROLE(ALLOW_CONNECTIONS = ALL)"
    Case Else
      strCmd        = strCmd & "SECONDARY_ROLE(ALLOW_CONNECTIONS = NO)"
  End Select
  strCmd            = strCmd & ");"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  If strSQLVersion >= "SQL2017" Then
    strCmd          = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] MODIFY REPLICA ON '" & strServerAO & "' "
    strCmd          = strCmd & "WITH (SEEDING_MODE=AUTOMATIC);"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)
  End If

  If strClusterHost = "YES" Then
    strCmd          = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] ADD LISTENER '" & strGroupAO & "' (WITH IP (" & GetClusterIPAddresses(strGroupAO, "AO", "SET") & "), PORT = 1433);"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
    WScript.Sleep strWaitShort
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strGroupAO & "_" & strGroupAO & """ /PRIV RegisterAllProvidersIP = 0"
    Call Util_RunExec(strCmd, "", strResponseYes, 5024)
    strCmd          = "CLUSTER """ & strClusterName & """ RESOURCE """ & strGroupAO & "_" & strGroupAO & """ /PRIV HostRecordTTL = 300"
    Call Util_RunExec(strCmd, "", strResponseYes, 5024)
  End If

  Call SetStatefileValue(strGroupAO, strServerAO)

End Sub


Sub ConfigureSecondaryAO()
  Call DebugLog("ConfigureSecondaryAO:")

  If strSetupSQLDBCluster <> "YES" Then
    strCmd          = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] "
    strCmd          = strCmd & "ADD REPLICA ON '" & strServerAO & "' "
    strCmd          = strCmd & "WITH (ENDPOINT_URL = 'TCP://" & strServerAO & "." & strUserDNSDomain & ":" & strTCPPortAO & "', SESSION_TIMEOUT = 10, BACKUP_PRIORITY = 50, PRIMARY_ROLE(ALLOW_CONNECTIONS = ALL)"
    Select Case True
      Case strEditionEnt = "YES"
        strCmd      = strCmd & ", SECONDARY_ROLE(ALLOW_CONNECTIONS = ALL)"
      Case Else
        strCmd      = strCmd & ", SECONDARY_ROLE(ALLOW_CONNECTIONS = NO)"
    End Select
    Select Case True
      Case strEditionEnt = "YES"
        strCmd      = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT"
      Case Else
        strCmd      = strCmd & ", FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = SYNCHRONOUS_COMMIT"
    End Select
    If strSQLVersion >= "SQL2017" Then
      strCmd        = strCmd & ", SEEDING_MODE=AUTOMATIC"
    End If
    strCmd          = strCmd & ")"
    Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strGroupAO & """   -E -b -e " & "-Q", """" & strCmd & ";""", 1)
  End If

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] JOIN"
  If strSetupSQLDBCluster <> "YES" Then
    strCmd          = strCmd & " WITH (CLUSTER_TYPE = NONE)"
  End If
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)

End Sub


Sub ConfigureAOXtras()
  Call DebugLog("ConfigureAOXtras:")

  If strSQLVersion >= "SQL2017" Then
    strCmd          = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] GRANT CREATE ANY DATABASE;"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)
  End If

  strCmd            = "ALTER EVENT SESSION AlwaysOn_health ON SERVER WITH (STARTUP_STATE=ON);"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  Select Case True
    Case strSetupAOAlias <> "YES"
      Call SetBuildMessage(strMsgInfo, "Recommendation: Create DNS Alias called " & strGroupAO & " to point to " & strServerAO)
    Case GetBuildfileValue("SetupAOAliasStatus") = strStatusProgress
      Call SetBuildfileValue("SetupAOAliasStatus", strStatusComplete)
  End Select

End Sub


Sub ConfigureDag()
  Call SetProcessId("4ALE", "Configure Distributed Availability Group")
  Dim strActionDAG

  strActionDAG      = GetBuildfileValue("ActionDAG")
  Select Case True
    Case strActionDAG = "ADDNODE"
      Call ConfigureSecondaryDag()
    Case Else 
      Call ConfigurePrimaryDag()
  End Select

  Call ConfigureLinkedServers(strGroupAO, strAGPrimary, strAGDagName)

  Call ProcessEnd(strStatusBypassed)

End Sub


Sub ConfigurePrimaryDag()
  Call DebugLog("ConfigurePrimaryDag:")

  strCmd            = "CREATE AVAILABILITY GROUP [" & strAGDagName & "]"
  strCmd            = strCmd & " WITH (DISTRIBUTED)"
  strCmd            = strCmd & " AVAILABILITY GROUP ON "
  strCmd            = strCmd & " '" & strGroupAO & "' WITH ("
  strCmd            = strCmd & "  LISTENER_URL = 'TCP://" & strServerAO & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
  strCmd            = strCmd & "  , FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT, SEEDING_MODE = AUTOMATIC"
  strCmd            = strCmd & "  );"
  Call Util_ExecSQL(strCmdSQL & "-V 11 -Q", """" & strCmd & """", -1)

  Select Case True
    Case strSetupAOAlias = "YES" 
      Call SetBuildfileValue("SetupAOAliasStatus", strStatusComplete)
    Case Else
      Call SetBuildMessage(strMsgInfo, "Recommendation: Create DNS Alias called " & strAGDagName & " to point to " & strGroupAO)
  End Select

  Call SetStatefileValue(strAGDagName, strGroupAO)
  Call SetStatefileValue(strAGDagName & "Server", strServerAO)

End Sub


Sub ConfigureSecondaryDag()
  Call DebugLog("ConfigureSecondaryDag:")
  Dim objSQL, objSQLData
  Dim strAGPrimaryServer, strAGDagNodes

  strAGPrimaryServer = GetStatefileValue(strAGPrimary & "Server")
  strAGDagNodes      = "0" & GetBuildfileValue("AGDagNodes")

  If strAGDagNodes < 2 Then
    strCmd            = "DROP AVAILABILITY GROUP [" & strAGDagName & "];"
    Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strAGPrimary & """ -E -b -e " & "-Q", """" & strCmd & ";""", -1)
    strCmd          = "CREATE AVAILABILITY GROUP [" & strAGDagName & "]"
    strCmd          = strCmd & " WITH (DISTRIBUTED)"
    strCmd          = strCmd & " AVAILABILITY GROUP ON "
    strCmd          = strCmd & " '" & strAGPrimary & "' WITH ("
    strCmd          = strCmd & "  LISTENER_URL = 'TCP://" & strAGPrimaryServer & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
    strCmd          = strCmd & "  , FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT, SEEDING_MODE = AUTOMATIC"
    strCmd          = strCmd & "  ),"
    strCmd          = strCmd & " '" & strGroupAO & "' WITH ("
    strCmd          = strCmd & "  LISTENER_URL = 'TCP://" & strServerAO & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
    strCmd          = strCmd & "  , FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT, SEEDING_MODE = AUTOMATIC"
    strCmd          = strCmd & "  );"
    Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strAGPrimary & """ -E -b -e " & "-Q", """" & strCmd & ";""", 1)
  End If

  strCmd            = "ALTER AVAILABILITY GROUP [" & strAGDagName & "] "
  strCmd            = strCmd & " JOIN AVAILABILITY GROUP ON "
  strCmd            = strCmd & " '" & strAGPrimary & "' WITH ("
  strCmd            = strCmd & "  LISTENER_URL= 'TCP://" & strAGPrimaryServer & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "  
  strCmd            = strCmd & "  , FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT, SEEDING_MODE = AUTOMATIC"
  strCmd            = strCmd & "  ),"
  strCmd            = strCmd & " '" & strGroupAO & "' WITH ("
  strCmd            = strCmd & "  LISTENER_URL = 'TCP://" & strServerAO & "." & strUserDNSDomain & ":" & strTCPPortAO & "' "
  strCmd            = strCmd & "  , FAILOVER_MODE = MANUAL, AVAILABILITY_MODE = ASYNCHRONOUS_COMMIT, SEEDING_MODE = AUTOMATIC"
  strCmd            = strCmd & "  );"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)

End Sub


Sub ConfigureLinkedServers(strServer, strAGPrimary, strGroupAO)
  Call DebugLog("ConfigureLinkedServers:")

  Call CreateLinkedServer(strGroupAO, strServer)

  Select Case True
    Case strAGPrimary = ""
      ' Nothing
    Case strServer = strAGPrimary
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      ' Nothing
    Case Else
      Call CreateLinkedServer(strAGPrimary, strServer)
      Call CreateLinkedServer(strServer,    strAGPrimary)
  End Select

End Sub


Sub CreateLinkedServer(strName, strLocation)
  Call DebugLog("CreateLinkedServer: " & strName & " at " & strLocation)

  strCmd            = "EXEC master.dbo.sp_addlinkedserver   @server =   N'" & strName & "', @srvproduct=N'SQL Server'"
  Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strLocation & """ -E -b -e " & "-Q", """" & strCmd & ";""", 1)
  strCmd            = "EXEC master.dbo.sp_addlinkedsrvlogin @rmtsrvname=N'" & strName & "', @useself=N'True', @locallogin=NULL, @rmtuser=NULL, @rmtpassword=NULL"
  Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strLocation & """ -E -b -e " & "-Q", """" & strCmd & ";""", 1)
  strCmd            = "EXEC master.dbo.sp_serveroption      @server =   N'" & strName & "', @optname=N'rpc out', @optvalue=N'true'"
  Call Util_ExecSQL("""" & strPathCmdSQL & """ -S """ & strLocation & """ -E -b -e " & "-Q", """" & strCmd & ";""", 1)

End Sub


Sub SetupBIXtras()
  Call SetProcessId("4B", "Business Intelligence Extras")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4BA"
      ' Nothing
    Case GetBuildfileValue("SetupSSDTBI") <> "YES"
      ' Nothing
    Case Else
      Call SetupSSDTBI()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4BB"
      ' Nothing
    Case GetBuildfileValue("SetupMDXStudio") <> "YES"
      ' Nothing
    Case Else
      Call SetupMDXStudio()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4BC"
      ' Nothing
    Case GetBuildfileValue("SetupBIDSHelper") <> "YES"
      ' Nothing
    Case Else
      Call SetupBIDSHelper()
  End Select

  Call SetProcessId("4BZ", "Business Intelligence Extras" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSSDTBI()
  Call SetProcessId("4BA", "Install SSDT-BI")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",  strPathAddComp & "SSDTBI")
  Call RunInstall("SSDTBI", GetBuildfileValue("SSDTBIexe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupMDXStudio()
  Call SetProcessId("4BB", "Install MDX Studio")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSysX86)
  Call SetXMLParm(objInstParm, "InstFile",   GetBuildfileValue("MDXexe"))
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "MDX Studio")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLAS)
  Call RunInstall("MDXStudio", GetBuildfileValue("MDXZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupBIDSHelper()
  Call SetProcessId("4BC", "Install BIDS Helper")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmLog",    "")
  Call SetXMLParm(objInstParm, "ParmReboot", "")
  Call SetXMLParm(objInstParm, "ParmSilent", "/S")
  Call RunInstall("BIDSHelper", GetBuildfileValue("BIDSexe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupISXtras()
  Call SetProcessId("4I", "Integration Services Extras")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4IA"
      ' Nothing
    Case GetBuildfileValue("SetupDTSDesigner") <> "YES"
      ' Nothing
    Case Else
      Call SetupDTSDesigner()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4IB"
      ' Nothing
    Case GetBuildfileValue("SetupDTSBackup") <> "YES"
      ' Nothing
    Case Else
      Call SetupDTSBackup()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4IC"
      ' Nothing
    Case GetBuildfileValue("SetupDimensionSCD") <> "YES"
      ' Nothing
    Case strVersionNet3 < "3.5.30729.01"
      Call SetBuildfileValue("SetupDimensionSCDStatus", strStatusBypassed)
    Case Else
      Call SetupDimensionSCD()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4ID"
      ' Nothing
    Case GetBuildfileValue("SetupRawReader") <> "YES"
      ' Nothing
    Case Else
      Call SetupRawReader()
  End Select

  Call SetProcessId("4IZ", " Integration Services Extras" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupDTSDesigner()
  Call SetProcessId("4IA", "Install DTS Designer")
  Dim objInstParm
  Dim strDTSFix, strPathDTS, strPathSSIS

  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("DTSDesigner", GetBuildfileValue("DTSmsi"), objInstParm)

  If GetBuildfileValue("SetupDTSDesignerStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("DTS Designer Post-install processing")
  strDTSFix         = GetBuildfileValue("DTSFix")
  strPathDTS        = strDirProgSysX86 & "\" & strSQLProgDir & "\80\Tools\Binn\"
  strPathSSIS       = strPathSSMSx86 & "binn\VSShell\Common7\IDE\"
  Select Case True
    Case GetBuildfileValue("SetupSSMS") <> "YES" 
      ' Nothing
    Case strSQLVersion = "SQL2005" and strEdition = "EXPRESS"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      Call DebugLog("Applying KB 917406 fix for DTS Designer")
      strDebugMsg1  = "Source: " & strPathSSIS
      Set objFile   = objFSO.GetFile(strPathSSIS & "sqlwb.exe.manifest")
      strPathNew    = strPathSSIS & "sqlwb.exe.manifest.original"
      strDebugMsg2  = "Target: " & strPathNew
      objFile.Copy strPathNew, True
      strPathOld    = strPathFBScripts
      strDebugMsg1  = "Source: " & strPathOld
      Set objFile   = objFSO.GetFile(strPathOld & strDTSFix)
      strPathNew    = strPathSSIS & "sqlwb.exe.manifest"
      strDebugMsg2  = "Target: " & strPathNew
      objFile.Copy strPathNew, True
    Case Else
      Call DebugLog("Copying DTS files for SSMS integration")
      strPathNew    = strPathSSMSX86 & "Binn\VSShell\Common7\IDE\Resources\1033"
      Call SetupFolder(strPathNew)
      Call CopyFile(strPathDTS & "semsfc.dll",                     strPathSSMSX86 & "Binn\VSShell\Common7\IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "semsfc.rll", strPathSSMSX86 & "Binn\VSShell\Common7\IDE\Resources\1033\")
      Call CopyFile(strPathDTS & "sqlgui.dll",                     strPathSSMSX86 & "Binn\VSShell\Common7\IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "sqlgui.rll", strPathSSMSX86 & "Binn\VSShell\Common7\IDE\Resources\1033\")
      Call CopyFile(strPathDTS & "sqlsvc.dll",                     strPathSSMSX86 & "Binn\VSShell\Common7\IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "sqlsvc.rll", strPathSSMSX86 & "Binn\VSShell\Common7\IDE\Resources\1033\")
      WScript.Sleep strWaitShort
  End Select

  Select Case True
    Case GetBuildfileValue("SetupBIDS") <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case Else
      Call DebugLog("Copying DTS files for BIDS integration")
      strPathNew    = strPathVS & "IDE\Resources\1033"
      Call SetupFolder(strPathNew)
      Call CopyFile(strPathDTS & "semsfc.dll",                     strPathVS & "IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "semsfc.rll", strPathVS & "IDE\Resources\1033\")
      Call CopyFile(strPathDTS & "sqlgui.dll",                     strPathVS & "IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "sqlgui.rll", strPathVS & "IDE\Resources\1033\")
      Call CopyFile(strPathDTS & "sqlsvc.dll",                     strPathVS & "IDE\")
      Call CopyFile(strPathDTS & "Resources\1033\" & "sqlsvc.rll", strPathVS & "IDE\Resources\1033\")
  End Select

  Call SetBuildfileValue("SetupDTSDesignerStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDTSBackup()
  Call SetProcessId("4IB", "Install DTS Backup 2000")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "MenuOption", "Move")
  Call SetXMLParm(objInstParm, "MenuSource", strUserProf & "\" & strMenuPrograms & "\DTSBackup 2000")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLIS)
  Call RunInstall("DTSBackup", GetBuildfileValue("DTSBackupmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupDimensionSCD()
  Call SetProcessId("4IC", "Install SSIS Dimension Merge SCD")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstFile", GetBuildfileValue("DimensionSCDmsi"))
  Call RunInstall("DimensionSCD", GetBuildfileValue("DimensionSCDZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupRawReader()
  Call SetProcessId("4ID", "Install SSIS Raw File Reader")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "SetupOption", "Copy")
  Call SetXMLParm(objInstParm, "InstOption",  "Menu")
  Call SetXMLParm(objInstParm, "InstTarget",  strPathVS & "Tools")
  Call SetXMLParm(objInstParm, "MenuOption",  "Build")
  Call SetXMLParm(objInstParm, "MenuName",    "SSIS Raw File Reader")
  Call SetXMLParm(objInstParm, "MenuPath",    strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLIS)
  Call RunInstall("RawReader", GetBuildfileValue("RawReaderexe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupReportXtras()
  Call SetProcessId("4R", "Report Services Extras")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAZ"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case Else
      Call SetupRS() 
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RB"
      ' Nothing
    Case GetBuildfileValue("SetupRptTaskPad") <> "YES"
      ' Nothing
    Case Else
      Call SetupRptTaskPad()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RC"
      ' Nothing
    Case GetBuildfileValue("SetupRSScripter") <> "YES"
      ' Nothing
    Case Else
      Call SetupRSSCripter()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RD"
      ' Nothing
    Case GetBuildfileValue("SetupRSLinkGen") <> "YES"
      ' Nothing
    Case Else
      Call SetupRSLinkGen()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RE"
      ' Nothing
    Case GetBuildfileValue("SetupPowerBIDesktop") <> "YES"
      ' Nothing
    Case Else
      Call SetupPowerBIDesktop()
  End Select

  Call SetProcessId("4RZ", " Report Services Extras" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupRS()
  Call SetProcessId("4RA", "Configure RS Service")
  Dim strRSActualMode

  Call SetBuildfileValue("Setup" & GetRSProcess() & "Status", strStatusProgress)
  strRSActualMode   = GetBuildfileValue("RSActualMode")
  If strRSActualMode = "" Then
    strRSActualMode = strRSInstallMode
  End If

  If strSetupSQLDBCluster = "YES" Then
    Call MoveToNode(strClusterGroupSQL, GetPrimaryNode(strClusterNameSQL))
  End If

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAA"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      Call InstallRS()
    Case strSetupPowerBI = "YES"
      Call InstallRS()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAB"
      ' Nothing
    Case Else
      Call CheckRSService()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RACZ"
      ' Nothing
    Case Else
      Call ResetRSPerms()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAD"
      ' Nothing
    Case strActionSQLRS = "ADDNODE"
      ' Nothing
    Case strSetupRSDB <> "YES"
      ' Nothing
    Case Else
      Call SetupRSDB()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAE"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      Call SetRSServiceAC()
    Case strSetupPowerBI = "YES"
      Call SetRSServiceAC()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAF"
      ' Nothing
    Case Else
      Call SetupRSRights()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAG"
      ' Nothing
    Case strSetupRSDB <> "YES"
      ' Nothing
    Case GetBuildfileValue("SetupRSIndexes") <> "YES"
      ' Nothing
    Case strActionSQLRS = "ADDNODE"
      Call SetBuildFileValue("SetupRSIndexesStatus", strStatusPreConfig)
    Case Else
      Call SetupRSIndexes()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAH"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      Call ConnectRSDB()
    Case strSetupPowerBI = "YES"
      Call ConnectRSDB()
    Case UCase(Left(strRSActualMode, 9)) <> UCase("FilesOnly")
      ' Nothing
    Case Else
      Call ConnectRSDB()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAI"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      Call SetRSDirectories()
    Case strSetupPowerBI = "YES"
      Call SetRSDirectories()
    Case UCase(Left(strRSActualMode, 9)) <> UCase("FilesOnly")
      ' Nothing
    Case Else
      Call SetRSDirectories()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAJ"
      ' Nothing
    Case strSetupPowerBI <> "YES"
      ' Nothing
    Case Else
      Call SetupPBUrlacl()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAK"
      ' Nothing
    Case Else
      Call SetupRSKey()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAL"
      ' Nothing
    Case GetBuildfileValue("SetupRSAdmin") <> "YES"
      ' Nothing
    Case GetBuildfileValue("ActionDAG") = "ADDNODE"
      ' Nothing
    Case strActionSQLRS = "ADDNODE"
      Call SetBuildfileValue("SetupRSAdminStatus", strStatusPreConfig)
    Case strSetupSQLRSCluster = "YES"
      Call ConfigRSAdmin()
    Case UCase(Left(strRSActualMode, 9)) = UCase("FilesOnly")
      Call SetBuildfileValue("SetupRSAdminStatus", strStatusBypassed)
    Case Else
      Call ConfigRSAdmin()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAM"
      ' Nothing
    Case GetBuildfileValue("SetupRSExec") <> "YES"
      ' Nothing
    Case Else
      Call ConfigRSExec()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAN"
      ' Nothing
    Case Else
      Call SetupRSResources()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAO"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      Call ConnectRSFarm()
    Case strSetupPowerBI = "YES"
      Call ConnectRSFarm()
    Case UCase(Left(strRSActualMode, 9)) <> UCase("FilesOnly")
      ' Nothing
    Case Else
      Call ConnectRSFarm()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAP"
      ' Nothing
    Case strSetupSQLRSCluster <> "YES"
      ' Nothing
    Case Else
      Call SetupRSCluster()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAQ"
      ' Nothing
    Case GetBuildfileValue("SetupRSAlias") <> "YES"
      ' Nothing
    Case Else
      Call ConfigRSAlias()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAR"
      ' Nothing
    Case Else
      Call SetupRSServiceDep()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAS"
      ' Nothing
    Case strSQLVersion > "SQL2005"
      ' Nothing
    Case Else
      Call ProcessSetRSIECompatibility()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RAT"
      ' Nothing
    Case GetBuildfileValue("SetupRSKeepAlive") <> "YES"
      ' Nothing
    Case Else
      Call SetupRSKeepAlive()
  End Select

  If strSetupPowerBI = "YES" Then
    Call SetBuildfileValue("SetupPowerBIStatus", strStatusComplete)
  End If
  Call SetBuildfileValue("SetupSQLRSStatus", strStatusComplete)
  Call SetProcessId("4RAZ", " Check SSRS Configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Function GetRSProcess()
  Dim strRSProcess

  Select Case True
    Case strSetupPowerBI = "YES"
      strRSProcess  = "PowerBI"
    Case Else
      strRSProcess  = "SQLRS"
  End Select

  GetRSProcess      = strRSProcess

End Function


Sub InstallRS()
  Call SetProcessId("4RAA", "Install RS for " & strSQLRSExe)
  Dim objInstParm
  Dim strPowerBIPID, strRSProcess
  
  strPowerBIPID     = GetBuildfileValue("PowerBIPID")
  strRSProcess      = GetRSProcess()
  Select Case True
    Case (strSetupPowerBI = "YES") And (strPowerBIPID <> "")
      Call SetXMLParm(objInstParm, "ParmXtra", "/IAcceptLicenseTerms /PID=" & strPowerBIPID)
    Case (strSetupPowerBI <> "YES") And (strPID <> "")
      Call SetXMLParm(objInstParm, "ParmXtra", "/IAcceptLicenseTerms /PID=" & strPID)
    Case Else
      Call SetXMLParm(objInstParm, "ParmXtra", "/IAcceptLicenseTerms /Edition=Dev")
  End Select
  Call SetXMLParm(objInstParm,  "CleanBoot",    "YES")
  Call SetXMLParm(objInstParm,  "PreConKey",    Mid(strHKLMSQL, 6) & "Instance Names\RS\" & strInstRSSQL)
  Call SetXMLParm(objInstParm,  "ParmLog",      "/Log")
  Call SetXMLParm(objInstParm,  "StatusOption", strStatusProgress)
  Call RunInstall(strRSProcess, strSQLRSexe,    objInstParm)

  Call ProcessEnd("")

End Sub


Sub CheckRSService()
  Call SetProcessId("4RAB", "Check RS Service")

  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstRS & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strPathNew > ""
      Call SetBuildfileValue("SetupSQLRSStatus", strStatusProgress)
    Case (GetBuildfileValue("RSInstallRetry") = "") And (GetBuildfileValue("SQLRSexe") <> "") ' Retry install one more if RS not fully installed
      Call SetBuildfileValue("RSInstallRetry", "Y")
      Call SetupReboot("4RAA", "Retry RS Install")
    Case Else
      strDebugMsg1  = "RS Executable: " & GetBuildfileValue("SQLRSexe")
      Call SetBuildMessage(strMsgError, "Reporting Services was not installed")
  End Select 

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case Else
      strPath       = strHKLMSQL & "Instance Names\RS\"
      strInstRegRS  = objShell.RegRead(strPath & strInstRSSQL)
      Call SetBuildfileValue("InstRegRS",      strInstRegRS)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ResetRSPerms()
  Call SetProcessId("4RAC", "Reset RS Permissions")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RACA"
      ' Nothing
    Case Else
      Call SetupIISServiceAc()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RACB"
      ' Nothing
    Case Else
      Call SetupRSCommand()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RACC"
      ' Nothing
    Case Else
      Call SetupRSFolderPerm()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4RACD"
      ' Nothing
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case Else
      Call SetupRSRegistry()
  End Select

  Call SetProcessId("4RACZ", " Reset RS Permissions" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupIISServiceAc()
  Call SetProcessId("4RACA", "Setup IIS Service Account")
  Dim strIISAccount

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strPath       = strDirSys & "\MICROSOFT.Net\Framework\v2.0.50727\Temporary ASP.Net Files"
      strIISAccount = GetBuildfileValue("NTAuth") & "\" & GetBuildfileValue("NTAuthNetwork")
    Case strSQLVersion <= "SQL2008R2"
      strPath       = strDirSys & "\MICROSOFT.Net\Framework\v2.0.50727\Temporary ASP.Net Files"
      strIISAccount = strRsAccount
    Case Else
      strPath       = strDirSys & "\MICROSOFT.Net\Framework\v4.0.30319\Temporary ASP.Net Files"
      strIISAccount = strRsAccount
  End Select
  strCmd            = "NET LOCALGROUP """ & strGroupIISIUsers & """ """ & FormatAccount(strIISAccount) & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  If Not objFSO.FolderExists(strPath) Then
    Call SetupFolder(strPath)
    strCmd        = """" & strPath & """ /T /C /E /G """ & FormatAccount(strGroupIISIUsers) & """:F"
    Call RunCacls(strCmd)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSCommand()
  Call SetProcessId("4RACB", "Setup RS Command")
  Dim strItemReg

  Select Case True
    Case (strSQLVersion >= "SQL2017") Or (strSetupPowerBI = "YES")
      strPath       = strHKLMSQL & strInstRegRS & strInstRSDir
      strItemReg    = Right(strPath, Len(strPath) -  InstrRev(strPath, "\"))
      strPath       = Left(strPath, InstrRev(strPath, "\"))
      objWMIReg.GetStringValue strHKLM,Mid(strPath, 6),strItemReg,strPathNew
      strPathNew    = strPathNew & "\Shared Tools\"
    Case Else
      strPath       = "SOFTWARE\Wow6432Node\Microsoft\Microsoft SQL Server\" & strSQLVersionNum & "\Tools\ClientSetup\"
      objWMIReg.GetStringValue strHKLM,strPath,"Path",strPathNew
      If IsNull(strPathNew) Then
        strPath     = "SOFTWARE\Microsoft\Microsoft SQL Server\" & strSQLVersionNum & "\Tools\ClientSetup\"
        objWMIReg.GetStringValue strHKLM,strPath,"Path",strPathNew
      End If
  End Select
  If Right(strPathNew, 1) <> "\" Then
    strPathNew      = strPathNew & "\"
  End If
  strCmdRS          = strPathNew
  Call SetBuildfileValue("CmdRS",  strCmdRS)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSFolderPerm()
  Call SetProcessId("4RACC", "Setup RS Folder Permissions")
  Dim strDirRS

  strDirRS          = objShell.RegRead(strHKLMSQL & strInstRegRS & strInstRSDir)
  If Right(strDirRS, 1) <> "\" Then
    strDirRS        = strDirRS & "\"
  End If

  Select Case True
    Case strSetupPowerBI = "YES"
      strPathSSRS   = strDirProgSys & "\Microsoft Power BI Report Server\" & strInstRSSQL
    Case Else
      strPathSSRS   = strDirProg & "\" & strInstRegRS & "\Reporting Services"
  End Select
  Call SetBuildfileValue("PathSSRS",  strPathSSRS)

  Call ResetDBAFilePerm(strDirProg)
  Call ResetFilePerm(strDirProg, strRsAccount)

  Call ResetDBAFilePerm(strDirRS) 
  Call ResetFilePerm(strDirRS, strRsAccount)

  If strActionSQLRS <> "ADDNODE" Then
    Call ResetDBAFilePerm(strDirRS & strInstRegRS)  
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSRegistry()
  Call SetProcessId("4RACD", "Setup RS Registry Permissions")

  Call ProcessUser("4RACD", "Process User RS Registry", "ProcessUserRSRegistry")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSDB()
  Call SetProcessId("4RAD", "Setup RS DB")
  Dim objInstParm
  Dim strFile, strFunction, strRSProcess, strRSTemp

  Call StartSQL()
  Call StartSSRS("FORCE")
  strRSProcess      = GetRSProcess()
  WScript.Sleep strWaitLong
  WScript.Sleep strWaitLong
  WScript.Sleep strWaitLong

  Call SetXMLParm(objInstParm, "PathMain",      strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",      "-v strRSDBName=""" & strRSDBName & """")
  Call SetXMLParm(objInstParm, "LogXtra",       "Set-RSDBs")
  Call RunInstall("RSDB",      "Set-RSDBs.sql", objInstParm)

  strRSTemp         = GetRSTemp()
  strFunction       = SetRSInParam("GenerateDatabaseCreationScript")
  objRSInParam.Properties_.Item("DatabaseName")     = strRSDBName
  objRSInParam.Properties_.Item("Lcid")             = intRSLcid
  Select Case True
    Case strSQLVersion <= "SQL2005"
      ' Nothing
    Case Else
      objRSInParam.Properties_.Item("IsSharePointMode") = False
  End Select

  If RunRSWMI(strFunction, "") = 0 Then
    strFile         = "Set-RSSchema.sql"
    strPath         = strRSTemp & strFile
    strDebugMsg1    = "Path: " & strPath
    Call WriteFile(strPath, objRSOutParam.Script)
    Call SetXMLParm(objInstParm,  "PathMain",      strRSTemp)
    Call SetXMLParm(objInstParm,  "ParmXtra",      "-d """ & strRSDBName & """")
    Call SetXMLParm(objInstParm,  "LogXtra",       strFile)
    Call RunInstall(strRSProcess, strFile,         objInstParm)
  End If

  strFunction       = SetRSInParam("GenerateDatabaseRightsScript")
  objRSInParam.Properties_.Item("DatabaseName")     = strRSDBName
  objRSInParam.Properties_.Item("IsRemote")         = False
  objRSInParam.Properties_.Item("IsWindowsUser")    = True
  objRSInParam.Properties_.Item("UserName")         = FormatAccount(strRsAccount)
  If RunRSWMI(strFunction, "") = 0 Then
    strFile         = "Set-RSRights.sql"
    strPath         = strRSTemp & strFile
    strDebugMsg1    = "Path: " & strPath
    Call WriteFile(strPath, objRSOutParam.Script)
    Call SetXMLParm(objInstParm,  "PathMain",      strRSTemp)
    Call SetXMLParm(objInstParm,  "ParmXtra",      "-d """ & strRSDBName & """")
    Call SetXMLParm(objInstParm,  "LogXtra",       strFile)
    Call RunInstall(strRSProcess, strFile,         objInstParm)
  End If

  strFunction       = SetRSInParam("GenerateDatabaseUpgradeScript")
  objRSInParam.Properties_.Item("DatabaseName")     = strRSDBName
  objRSInParam.Properties_.Item("ServerVersion")    = "C.0.9.45"
  If RunRSWMI(strFunction, "-2147220938 -2147220960") = 0 Then ' Exclude when no update required
    strFile         = "Set-RSUpgrade.sql"
    strPath         = strRSTemp & strFile
    strDebugMsg1    = "Path: " & strPath
    Call WriteFile(strPath, objRSOutParam.Script)
    Call SetXMLParm(objInstParm,  "PathMain",      strRSTemp)
    Call SetXMLParm(objInstParm,  "ParmXtra",      "-d """ & strRSDBName & """")
    Call SetXMLParm(objInstParm,  "LogXtra",       strFile)
    Call RunInstall(strRSProcess, strFile,         objInstParm)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Function GetRSTemp()
  Dim strRSTemp

  Select Case True
    Case strClusterNameSQL = ""
      strRSTemp     = strPathTemp
    Case GetBuildfileValue("VolTempSource") = "C"
      strRSTemp     = strPathTemp
    Case Else
      strRSTemp     = strDirData
  End Select

  If Right(strRSTemp, 1) <> "\" Then
    strRSTemp       = strRSTemp & "\"
  End If

  GetRSTemp         = strRSTemp

End Function


Sub WriteFile(strFile, strText)

  Set objFile       = objFSO.CreateTextFile(strFile, True)
  Wscript.Sleep strWaitShort
  objFile.Write strText
  objFile.Close

End Sub


Sub SetRSServiceAC()
  Call SetProcessId("4RAE", "Set RS Service Account")
  Dim strFunction

  Call StopSSRS()

  strFunction       = SetRSInParam("SetWindowsServiceIdentity")
  objRSInParam.Properties_.Item("UseBuiltInAccount") = False
  objRSInParam.Properties_.Item("Account")           = strRsAccount
  objRSInParam.Properties_.Item("Password")          = strRsPassword
  Select Case True
    Case strOSVersion <> "6.2"
      Call RunRSWMI(strFunction, "")
    Case strSetupPowerBI <> "YES"
      Call RunRSWMI(strFunction, "")
    Case Else
      ' Nothing
  End Select

  Select Case true
    Case strSetupRSDB <> "YES"
      ' Nothing
    Case strActionSQLRS = "ADDNODE"
      ' Nothing
    Case Else
      strCmd        = "CREATE LOGIN [" & strRsAccount & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-r -Q", """" & strCmd & """", 1)
      strCmd        = "CREATE USER [" & strRsAccount & "] FOR LOGIN [" & strRsAccount & "];"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & """ -Q", """" & strCmd & """", 1)
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & "TempDB"" -Q", """" & strCmd & """", 1)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='db_owner', @MEMBERNAME='" & strRsAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & """ -Q", """" & strCmd & """", 0)
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & "TempDB"" -Q", """" & strCmd & """", 1)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='RSExecRole', @MEMBERNAME='" & strRsAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & """ -Q", """" & strCmd & """", 0)
      Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & "TempDB"" -Q", """" & strCmd & """", 1)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSRights()
  Call SetProcessId("4RAF", "Setup RS DB Rights")
  Dim objInstParm
  Dim strFile, strFunction, strRSProcess, strRSTemp

  Call StartSQL()

  strRSProcess      = GetRSProcess()
  strRSTemp         = GetRSTemp()
  Select Case True
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("SetupAlwaysOn") <> "YES"
      ' Nothing
    Case strActionSQLRS <> "ADDNODE"
      ' Nothing
    Case Else
      strFile       = "Role MasterRsExecRole.sql"
      strPath       = """" & strRSTemp & strFile & """"
      strCmd        = "EXEC " & strGroupAO & ".master.dbo.sp_ScriptRoles @Role='RSExecRole' -o " & strPath
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)
      WScript.Sleep strWaitShort
      Call SetXMLParm(objInstParm,  "PathMain",      strRSTemp)
      Call SetXMLParm(objInstParm,  "ParmXtra",      "-d ""master""")
      Call SetXMLParm(objInstParm,  "LogXtra",       strFile)
      Call RunInstall(strRSProcess, strFile,         objInstParm)
      strFile       = "Role MsdbRsExecRole.sql"
      strPath       = """" & strRSTemp & strFile & """"
      strCmd        = "EXEC " & strGroupAO & ".msdb.dbo.sp_ScriptRoles @Role='RSExecRole' -o " & strPath
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)
      WScript.Sleep strWaitShort
      Call SetXMLParm(objInstParm,  "PathMain",      strRSTemp)
      Call SetXMLParm(objInstParm,  "ParmXtra",      "-d ""msdb""")
      Call SetXMLParm(objInstParm,  "LogXtra",       strFile)
      Call RunInstall(strRSProcess, strFile,         objInstParm)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSIndexes()
  Call SetProcessId("4RAG", "Setup RS DB Indexes and options")
  Dim objInstParm

  Call StartSQL()

  Call SetXMLParm(objInstParm, "PathMain",            strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",            "-v strRSDBName=""" & strRSDBName & """")
  Call RunInstall("RSIndexes", "Set-RSDBOptions.sql", objInstParm)

  Call ProcessEnd("")

End Sub


Sub ConnectRSDB()
  Call SetProcessId("4RAH", "Connect RS to Database")
  Dim strFunction, strServer

  Call StartSQL()
  Call StartSSRS("FORCE")

  Select Case True
    Case strCatalogInstance = ""
      strServer     = strCatalogServerName
    Case strCatalogInstance = "MSSQLSERVER"
      strServer     = strCatalogServerName
    Case Else
      strServer     = strCatalogServerName & "\" & strCatalogInstance
  End Select

  Call DebugLog("Connecting to " & strRSDBName & " on "  & strServer)
  strFunction       = SetRSInParam("SetDatabaseConnection")
  objRSInParam.Properties_.Item("Server")            = strServer
  objRSInParam.Properties_.Item("DatabaseName")      = strRSDBName
  objRSInParam.Properties_.Item("CredentialsType")   = 2 ' Use Windows Service account
  objRSInParam.Properties_.Item("UserName")          = ""
  objRSInParam.Properties_.Item("Password")          = ""
  Call RunRSWMI(strFunction, "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetRSDirectories()
  Call SetProcessId("4RAI", "Set RS Directories")

  Call StopSSRS()
  Call StartSQL()
  Call StartSSRS("FORCE")

  Call SetRSDirectory("ReportServerWebService", "ReportServer" & strRSURLSuffix)

  Select Case True
    Case strSQLVersion >= "SQL2016"
      Call SetRSDirectory("ReportServerWebApp", "Reports" & strRSURLSuffix)
    Case strSetupPowerBI = "YES"
      Call SetRSDirectory("ReportServerWebApp", "Reports" & strRSURLSuffix)
    Case Else
      Call SetRSDirectory("ReportManager",      "Reports" & strRSURLSuffix)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub



Sub SetupPBUrlacl()
  Call SetProcessId("4RAJ", "Setup PowerBI HTTP Reservations")

  Select Case True
    Case strRSAlias <> "" 
      strCmd        = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & strRSAlias & ":" & strTCPPortRS & "/PowerBI/"" USER=""NT SERVICE\PowerBIReportServer"" "
      Call Util_RunExec(strCmd, "", "", -1)
      strCmd        = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & strRSAlias & ":" & strTCPPortRS & "/wopi/""    USER=""NT SERVICE\PowerBIReportServer"" "
      Call Util_RunExec(strCmd, "", "", -1)
    Case strSetupSQLRSCluster = "YES"
      strCmd        = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & GetBuildfileValue("ClusterGroupRS") & ":" & strTCPPortRS & "/PowerBI/"" USER=""NT SERVICE\PowerBIReportServer"" "
      Call Util_RunExec(strCmd, "", "", -1)
      strCmd        = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & GetBuildfileValue("ClusterGroupRS") & ":" & strTCPPortRS & "/wopi/""    USER=""NT SERVICE\PowerBIReportServer"" "
      Call Util_RunExec(strCmd, "", "", -1)
  End Select

  strCmd            = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & strServer & ":" & strTCPPortRS & "/PowerBI/"" USER=""NT SERVICE\PowerBIReportServer"" "
  Call Util_RunExec(strCmd, "", "", -1)
  strCmd            = "NETSH HTTP ADD URLACL URL=""" & strHTTP & "://" & strServer & ":" & strTCPPortRS & "/wopi/""    USER=""NT SERVICE\PowerBIReportServer"" "
  Call Util_RunExec(strCmd, "", "", -1)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSKey()
  Call SetProcessId("4RAK", "Setup RS Encryption Key")
  Dim strRSKeyAction, strRSPwd

  Call StopSSRS()
  Call StartSQL()
  Call StartSSRS("FORCE")

  strRSKeyAction    = ""
  Select Case True
    Case (strEdition = "EXPRESS") And (UCase(strSQLRSExe) = UCase("PowerBIReportServer.exe"))
      ' Nothing
    Case GetBuildfileValue("ActionDAG") = "ADDNODE"
      strRSKeyAction = "ADDNODE"
    Case Left(strActionSQLRS, 7) = "INSTALL"
      strRSKeyAction = "INSTALL"
    Case strActionSQLRS <> "ADDNODE"
      ' Nothing
    Case (GetBuildfileValue("VolBackupSource") = "D") And (GetBuildfileValue("VolBackupType") <> "L")
      ' Nothing - Backup volume offline to this node
    Case Else
      strRSKeyAction = "ADDNODE"
  End Select

  Select Case True
    Case strRsExecPassword <> ""
      strRSPwd      = strRsExecPassword
    Case Else
      strRSPwd      = strsaPwd
  End Select

  strPath           = strDirSystemDataPrimary & "\RSEncryptionKey.snk"
  Select Case True
    Case strRSKeyAction = "INSTALL"
      If objFSO.FileExists(strPath) Then
        Call objFSO.DeleteFile(strPath, True)
        Wscript.Sleep strWaitShort ' Wait for NTFS Cache to catch up to avoid Permissions error
      End If
      strCmd        = """" & strCmdRS & "RSKEYMGMT.EXE"" -i """ & strInstRSSQL & """ -e -f """ & strPath & """ -p """ & strRSPwd & """ "
      Call Util_RunExec("%COMSPEC% /D /C Echo " & strResponseYes & "| " & strCmd, "", strResponseYes, "0 1")
      If intErrSave = 1 Then ' Wait and retry if first attempt fails
        WScript.Sleep strWaitLong
        WScript.Sleep strWaitLong
        WScript.Sleep strWaitLong
        Call Util_RunExec("%COMSPEC% /D /C Echo " & strResponseYes & "| " & strCmd, "", strResponseYes, 0)
      End If
    Case strRSKeyAction = "ADDNODE"
      strCmd        = """" & strCmdRS & "RSKEYMGMT.EXE"" -i """ & strInstRSSQL & """ -a -f """ & strPath & """ -p """ & strRSPwd & """ "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigRSAdmin()
  Call SetProcessId("4RAL", "Reporting Services Administration Accounts")
  Dim strPolicy

  Call StartSQL()
  Call StartSSRS("FORCE")

  If strType = "WORKSTATION" Then
    strPolicy       = " -v keepCurrentPolicy=""True"""
  Else
    strPolicy       = " -v keepCurrentPolicy=""False"""
  End If

  Call DebugLog("Set Server Administration Account security")
  If strGroupDBA <> strLocalAdmin Then
    strCmd          = """" & strCmdRS & "RS.EXE"" -i """ & strPathFBScripts & "Set-RSSecurity.rss"" -s """ & strHTTP & "://" & strInstRSHost & "/" & strInstRSURL & """ -v userName=""" & strGroupDBAAlt & """ -v roleName=""System Administrator""" & strPolicy
    Call Util_RunExec(strCmd, "", "", 1)
  End If

  If strGroupDBANonSA <> "" Then
    strCmd          = """" & strCmdRS & "RS.EXE"" -i """ & strPathFBScripts & "Set-RSSecurity.rss"" -s """ & strHTTP & "://" & strInstRSHost & "/" & strInstRSURL & """ -v userName=""" & strGroupDBANonSAAlt & """ -v roleName=""System User"" -v keepCurrentPolicy=""True"""
    Call Util_RunExec(strCmd, "", "", 1)
  End If

  Call DebugLog("Set Report Account security")
  If strGroupDBA <> strLocalAdmin Then
    strCmd          = """" & strCmdRS & "RS.EXE"" -i """ & strPathFBScripts & "Set-RSSecurity.rss"" -s """ & strHTTP & "://" & strInstRSHost & "/" & strInstRSURL & """ -v userName=""" & strGroupDBAAlt & """ -v roleName=""Content Manager""" & strPolicy
    Call Util_RunExec(strCmd, "", "", 1)
  End If

  If strGroupDBANonSA <> "" Then
    strCmd          = """" & strCmdRS & "RS.EXE"" -i """ & strPathFBScripts & "Set-RSSecurity.rss"" -s """ & strHTTP & "://" & strInstRSHost & "/" & strInstRSURL & """ -v userName=""" & strGroupDBANonSAAlt & """ -v roleName=""Content Manager"" -v keepCurrentPolicy=""True"""
    Call Util_RunExec(strCmd, "", "", 1)
  End If

  Call SetBuildfileValue("SetupRSAdminStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigRSExec()
  Call SetProcessId("4RAM", "Reporting Services Execution Account")
  Dim colRSConf
  Dim objRSConf
  Dim strFunction, strRSEmail, strRsShareAccount, strRsShareDomain, strRsSharePassword, strRSVersion

  Call StartSQL()
  Call StartSSRS("FORCE")

  strRSEmail        = GetBuildfileValue("RSEmail")
  strRSVersion      = GetBuildfileValue("RSVersion")

  Select Case True
    Case strRsExecAccount = ""
      ' Nothing   
    Case strRsExecPassword = ""
     ' Nothing
    Case Else
      strFunction   = SetRSInParam("SetUnattendedExecutionAccount")
      objRSInParam.Properties_.Item("UserName")            = strRsExecAccount
      objRSInParam.Properties_.Item("Password")            = strRsExecPassword
      Call RunRSWMI(strFunction, "")
  End Select

  Select Case True
    Case strMailServer = ""
      ' Nothing
    Case strRSEmail = ""
      ' Nothing
    Case Else
      strFunction   = SetRSInParam("SetEmailConfiguration")
      objRSInParam.Properties_.Item("SendUsingSMTPServer") = True
      objRSInParam.Properties_.Item("SMTPServer")          = strMailServer
      objRSInParam.Properties_.Item("SenderEmailAddress")  = strRSEmail
      Call RunRSWMI(strFunction, "")
  End Select

  strRsShareAccount  = GetBuildfileValue("RsShareAccount")
  strRsSharePassword = GetBuildfileValue("RsSharePassword")
  Select Case True
    Case strRsShareAccount = ""
      ' Nothing
    Case strRsSharePassword = ""
      ' Nothing
    Case (strSQLVersion >= "SQL2016") Or (strSetupPowerBI = "YES")
      strFunction   = SetRSInParam("SetFileShareAccount")
      objRSInParam.Properties_.Item("Account")             = strRsShareAccount
      objRSInParam.Properties_.Item("Password")            = strRsSharePassword
      Call RunRSWMI(strFunction, "")
  End Select

  Call SetBuildfileValue("SetupRSExecStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSResources()
  Call SetProcessId("4RAN", "Setup RS Resource Options")
  Dim objNode, objRSConfig, objRSNode
  Dim strName, strRSConfig, strRSName, strQueryLen

  Set objRSConfig   = CreateObject("Microsoft.XMLDOM")
  objRSConfig.async = "false"

  Call DebugLog("Set web.config Options")
  strRSConfig       = strPathSSRS & "\ReportServer\web.config"
  strDebugMsg1      = "Source: " & strRSConfig
  objFSO.CopyFile strRSConfig, strRSConfig & ".original"
  objRSConfig.load(strRSConfig)

  strQueryLen       = GetBuildfileValue("SetHeaderLength")
  Select Case True
    Case strOSVersion < "6"
      ' Nothing
    Case Else
      Call SetXMLConfigValue(objRSConfig, "system.web/httpRuntime", "maxQueryStringLength", strQueryLen, "A")
  End Select
  objRSConfig.save strRSConfig

  strPath           = "HKLM\SYSTEM\CurrentControlSet\Services\HTTP\Parameters\"
  Call Util_RegWrite(strPath & "MaxFieldLength",  strQueryLen,           "REG_DWORD")
  Call Util_RegWrite(strPath & "MaxRequestBytes", CStr(2 * strQueryLen), "REG_DWORD")

  Call DebugLog("Set rsreportserver.config Options")
  strRSConfig       = strPathSSRS & "\ReportServer\rsreportserver.config"
  strDebugMsg1      = "Source: " & strRSConfig
  objFSO.CopyFile strRSConfig, strRSConfig & ".original"
  objRSConfig.load(strRSConfig)

  Select Case True
    Case strSQLVersion <= "SQL2005"
      ' Nothing
    Case strRSAlias <> ""
      Call SetXMLConfigValue(objRSConfig, "Service", "UrlRoot",         strHTTP & "://" & strRSAlias & "/ReportServer", "")
      Call SetXMLConfigValue(objRSConfig, "UI",      "ReportServerUrl", strHTTP & "://" & strRSAlias & "/ReportServer", "")
    Case strSetupSQLRSCluster = "YES"
      Call SetXMLConfigValue(objRSConfig, "Service", "UrlRoot",         strHTTP & "://" & GetBuildfileValue("ClusterGroupRS") & "/ReportServer", "")
      Call SetXMLConfigValue(objRSConfig, "UI",      "ReportServerUrl", strHTTP & "://" & GetBuildfileValue("ClusterGroupRS") & "/ReportServer", "")
    Case Else
      Call SetXMLConfigValue(objRSConfig, "Service", "UrlRoot",         strHTTP & "://" & strServer & "/ReportServer", "")
      Call SetXMLConfigValue(objRSConfig, "UI",      "ReportServerUrl", strHTTP & "://" & strServer & "/ReportServer", "")
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      If GetBuildfileValue("SetWorkingSetMaximum") > "0" Then
'        Call SetXMLConfigValue(objRSConfig, "Service", "MaximumMemoryLimit", GetBuildfileValue("SetWorkingSetMaximum"), "")
      End If
    Case Else
      If GetBuildfileValue("SetWorkingSetMaximum") > "0" Then
        Call SetXMLConfigValue(objRSConfig, "Service", "WorkingSetMaximum", GetBuildfileValue("SetWorkingSetMaximum"), "")
      End If
      Call DebugLog("Enable Kerberos Security")
      Call SetXMLConfigValue(objRSConfig, "Authentication/AuthenticationTypes", "RSWindowsNegotiate", "", "")
      Call SetXMLConfigValue(objRSConfig, "Authentication/AuthenticationTypes", "RSWindowsKerberos",  "", "")
      Call SetXMLConfigValue(objRSConfig, "Authentication/AuthenticationTypes", "RSWindowsNTLM",      "", "")
  End Select

  objRSConfig.save strRSConfig

  Call DebugLog("Set Site Name")
  strName           = RTrim(" " & GetBuildfileValue("RSName"))
  Select Case True
    Case strRSAlias <> ""
      strCmd        = "UPDATE ConfigurationInfo SET SiteName = '" & strRSAlias & strName & "';"
    Case strInstRSSQL = "MSSQLSERVER"
      strCmd        = "UPDATE ConfigurationInfo SET SiteName = '" & strServer & strName & "';"
    Case strInstRSSQL = "SSRS"
      strCmd        = "UPDATE ConfigurationInfo SET SiteName = '" & strServer & strName & "';"
    Case Else 
      strCmd        = "UPDATE ConfigurationInfo SET SiteName = '" & strInstRSSQL & strName & "';"
  End Select
  Call Util_ExecSQL(strCmdSQL & "-d """ & strRSDBName & """ -Q", """" & strCmd & """", 1)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConnectRSFarm()
  Call SetProcessId("4RAO", "Connect SSRS Catalog to Farm")

  Call StartSQL()
  Call StartSSRS("FORCE")

  strCmd            = """" & strCmdRS & "RSKEYMGMT.EXE"" -i """ & strInstRSSQL & """ -j -m """ & strCatalogServerName & """ -n """ & strInstRSSQL & """ "
  Select Case True
    Case (strActionSQLRS = "ADDNODE") And (GetBuildfileValue("VolBackupSource") = "D") And (GetBuildfileValue("VolBackupType") <> "L")
      ' Nothing - Backup volume offline to this node
    Case strActionSQLRS = "ADDNODE"
      Call Util_RunExec(strCmd, "", strResponseYes, 1)
    Case strEditionEnt = "YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case Else
      ' Nothing - not needed for Standard Edition
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSCluster()
  Call SetProcessId("4RAP", "Setup RS Cluster")

  Select Case True
    Case strActionSQLRS = "ADDNODE"
      Call StopSSRS()
      Call AddOwner(GetBuildfileValue("ClusterNameRS"))
    Case Else
      Call BuildCluster("RS", GetBuildfileValue("ClusterGroupRS"), "Reporting Services", GetBuildfileValue("ClusterNameRS"), "Generic Service", strInstRS, "SQL Server Reporting Services", "SOFTWARE\Microsoft\Microsoft SQL Server\" & strInstRegRS & "\MSSQLServer", "", "L")
  End Select

  Call SetBuildfileValue("SetupSQLRSClusterStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigRSAlias()
  Call SetProcessId("4RAQ", "Setup IIS Alias for SSRS")
  Dim objFile
  Dim strHost

  Select Case True
    Case strRSAlias <> ""
      strHost       = strRSAlias
    Case GetBuildfileValue("ClusterIPV4RS") <> ""
      strHost       = GetBuildfileValue("ClusterIPV4RS")
    Case GetBuildfileValue("ClusterIPV6RS") <> ""
      strHost       = GetBuildfileValue("ClusterIPV6RS")
    Case Else
      strHost       = strInstRSHost
  End Select

  If strSetupIIS = "YES" Then
    strHost         = RTrim(Replace(strHost, """", "")) & ":" & strTCPPortRS
    strPath         = strIISRoot & "\DEFAULT.HTM"
    Set objFile     = objFSO.CreateTextFile(strPath, True)
    Wscript.Sleep strWaitShort
    Select Case True
      Case strInstRSURL = "ReportServer"
        strCmd      = "<script>window.location='" & strHTTP & "://" & strHost & "/reports'</script>"
      Case Else
        strCmd      = "<script>window.location='" & strHTTP & "://" & strHost & "/" & strInstRSURL & "/reports'</script>"
    End Select
    objFile.WriteLine strCmd
    objFile.Close
  End If

  Select Case True
    Case strRSAlias = ""
      ' Nothing
    Case Else
      strPath       = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\" & strRSAlias & "\"
      Call Util_RegWrite(strPath & strHTTP, "1", "REG_DWORD") ' Add to Intranet
      strPath       = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\" & strRSAlias & "\"
      Call Util_RegWrite(strPath & strHTTP, "1", "REG_DWORD")
      strPath       = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\az416426.vo.msecnd.net\" ' MS Application Insights
      Call Util_RegWrite(strPath & "https", "4", "REG_DWORD") ' Add to Blocked
      strPath       = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\EscDomains\az416426.vo.msecnd.net\"
      Call Util_RegWrite(strPath & "https", "4", "REG_DWORD")
  End Select

  Call SetBuildfileValue("SetupRSAliasStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSServiceDep()
  Call SetProcessId("4RAR", "Setup RS Service Dependency")

  If strSetupSQLDB = "YES" Then
    strCmd          = "SC CONFIG "  & strInstRS & " DEPEND= " & strInstSQL
    Call Util_RunExec(strCmd, "", "", 2)
  End If

End Sub


Sub ProcessSetRSIECompatibility()
  Call SetProcessId("4RAS", "Setup SSRS IE Compatibility")
  Dim strSSRSPath, strSSRSFile, strOldText, strNewText

  strSSRSPath       = "ReportServer\Pages\"
  strSSRSFile       = "ReportViewer.aspx"
  Select Case True
    Case strEdition > "SQL2008"
      ' Nothing
    Case Else
      strOldText    = "</title>" & Chr(13) & Chr(10) & "</head>"
      strNewText    = "</title>" & Chr(13) & Chr(10)
      strNewText    = strNewText & "  <meta http-equiv=""X-UA-Compatible"" content=""IE=EmulateIE7"">"
      Call SetRSIECompatibility(strSSRSPath, strSSRSFile, strOldText, strNewText)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetRSIECompatibility(strSSRSPath, strSSRSFile, strOldText, strNewText)
  Call DebugLog("SetRSIECompatibility: " & strSSRSFile)
' Based on www.sqlservercentral.com/Forums/Topic1463247-150-1.aspx#bm1464554
' Postings from Neodynamic and Jhickman_21029
  Dim strSSRSConfig, strSSRSPathFile, strTestText

  Call DebugLog("Saving original SSRS file")
  strPath           = strHKLMSQL & strInstRegRS & strInstRSDir
  strPath           = objShell.RegRead(strPath) & strSSRSPath
  strSSRSPathFile   = strPath & strSSRSFile
  strDebugMsg1      = "Source: " & strSSRSPathFile
  objFSO.CopyFile strSSRSPathFile, strSSRSPathFile & ".original" 

  Call DebugLog("Setting SSRS configuration values")
  Set objFile       = objFSO.OpenTextFile(strSSRSPathFile, 1)
  strSSRSConfig  = objFile.ReadAll
  objFile.Close

  strTestText     = Left(strOldText, Instr(strOldText, ">"))
  If Instr(Replace(strSSRSConfig, " ", ""), strOldText) > 0 Then
    strSSRSConfig  = Replace(strSSRSConfig, strTestText, strNewText)
  End If

  Set objFile = objFSO.OpenTextFile(strSSRSPathFile, 2)
  objFile.WriteLine strSSRSConfig
  objFile.Close

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSKeepAlive()
  Call SetProcessId("4RAT", "Setup KeepAlive Job for SSRS")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call SetXMLParm(objInstParm, "InstFile",    "Install.vbs")
  Call RunInstall("RSKeepAlive", GetBuildfileValue("RSKeepAliveCab"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupRptTaskPad()
  Call SetProcessId("4RB", "Install Taskpad View Report")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "SetupOption", "Copy")
  Call SetXMLParm(objInstParm, "InstOption",  "None")
  Call SetXMLParm(objInstParm, "InstTarget",  strDirDBA & "\SQL Server Management Studio\Custom Reports\")
  Call RunInstall("RptTaskPad", GetBuildfileValue("RptTaskPadRdl"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupRSScripter()
  Call SetProcessId("4RC", "Install Report Services Scripter")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSysX86)
  Call SetXMLParm(objInstParm, "InstFile",   "RSScripter.exe")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "RSScripter")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLRS)
  Call RunInstall("RSScripter", GetBuildfileValue("RSScripterZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupRSLinkGen()
  Call SetProcessId("4RD", "Install Linked Report Generator")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSysX86)
  Call SetXMLParm(objInstParm, "InstFile",   "RSLinkgen.exe")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "Linked Report Generator")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLRS)
  Call RunInstall("RSLinkGen", GetBuildfileValue("RSLinkGenZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupPowerBIDesktop()
  Call SetProcessId("4RE", "Install PowerBI Desktop")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmXtra",    "/IAcceptLicenseTerms")
  Call SetXMLParm(objInstParm, "ParmLog",     "/log")
  Call RunInstall("PowerBIDesktop", "PBIDesktopRS_x64.exe", objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSQLXtras()
  Call SetProcessId("4S", "SQL Server Extras")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strprocessId > "4SAZ"
      ' Nothing
    Case GetBuildfileValue("SetupBPAnalyzer") <> "YES"
      ' Nothing
    Case Else
      Call SetupBPAnalyzer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SB"
      ' Nothing
    Case GetBuildfileValue("SetupJavaDBC") <> "YES"
      ' Nothing
    Case Else
      Call SetupJavaDBC()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SCZ"
      ' Nothing
    Case GetBuildfileValue("SetupDB2OLE") <> "YES"
      ' Nothing
    Case Else
      Call InstallDB2OLE()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SD"
      ' Nothing
    Case GetBuildfileValue("SetupCacheManager") <> "YES"
      ' Nothing
    Case Else
      Call SetupCacheManager()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SE"
      ' Nothing
    Case GetBuildfileValue("SetupIntViewer") <> "YES"
      ' Nothing
    Case Else
      Call SetupIntViewer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SFZ"
      ' Nothing
    Case GetBuildfileValue("SetupMDS") <> "YES"
      ' Nothing
    Case Else
      Call SetupMDS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SG"
      ' Nothing
    Case GetBuildfileValue("SetupPerfDash") <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005" And Left(strSQLVersionFull, 3) < "9.2"
      Call SetBuildfileValue("SetupPerfDashStatus", strStatusBypassed) ' Performance Dashboard is dependant on SQL2005 SP2
    Case Else
      Call SetupPerfDash()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SH"
      ' Nothing
    Case GetBuildfileValue("SetupSystemViews") <> "YES"
      ' Nothing
    Case Else
      Call SetupSystemViews()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SI"
      ' Nothing
    Case GetBuildfileValue("SetupSQLNS") <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"     ' NS included in SQL 2005 install
      ' Nothing
    Case Else
      Call SetupSQLNS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SJ"
      ' Nothing
    Case GetBuildfileValue("SetupStreamInsight") <> "YES"
      ' Nothing
    Case Else
      Call SetupStreamInsight()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SKZ"
      ' Nothing
    Case GetBuildfileValue("SetupSamples") <> "YES"
      ' Nothing
    Case Else
'      Call SetupSamples()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SL"
      ' Nothing
    Case GetBuildfileValue("SetupSemantics") <> "YES"
      ' Nothing
    Case Else
      Call SetupSemantics()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SMZ"
      ' Nothing
    Case GetBuildfileValue("SetupDQ") <> "YES"
      ' Nothing
    Case Else
      Call SetupDQ()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SN"
      ' Nothing
    Case GetBuildfileValue("SetupDistributor") <> "YES"
      ' Nothing
    Case Else
      Call SetupDistributor()
  End Select

  Call SetProcessId("4SZ", " SQL Server Extras" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub InstallBPAnalyzer()
  Call SetProcessId("4SA", "Install Best Practice Analyzer")

  strSetupBPAnalyzer = GetBuildfileValue("SetupBPAnalyzer")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strprocessId > "4SAA"
      ' Nothing
    Case GetBuildfileValue("SetupBPAnalyzer") <> "YES"
      ' Nothing
    Case Else
      Call SetupBPAnalyzer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SAB"
      ' Nothing
    Case GetBuildfileValue("SetupKB2781514") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2781514()
  End Select

  Call SetProcessId("4SAZ", " Install Best Practice Analyzer" & strStatusComplete)
  Call ProcessEnd("")

End Sub



Sub SetupBPAnalyzer()
  Call SetProcessId("4SAA", "Setup Best Practice Analyzer")
  Dim objInstParm
  Dim strMenuFolder

  Select Case True
    Case strSQLVersion = "SQL2005"
      strMenuFolder = "SQL Server 2005 BPA"
    Case strSQLVersion = "SQL2008"
      strMenuFolder = "SQL Server 2005 BPA"
    Case strSQLVersion = "SQL2008R2"
      strMenuFolder = "SQL Server 2008 R2 BPA"
    Case strSQLVersion >= "SQL2012"
      strMenuFolder = "SQL Server 2012 BPA"
      If strOSVersion <= "6.2" Then
        strCmd      = strCompatFlags & "{9660e391-5510-4ee2-8552-85bd9e5b5aa1}"
        Call Util_RegWrite(strCmd, 4, "REG_DWORD")
      End If
  End Select

  Call SetXMLParm(objInstParm, "MenuOption", "Move")
  Call SetXMLParm(objInstParm, "MenuSource", strAllUserProf & "\" & strMenuPrograms & "\" & strMenuFolder)
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuConfigTools)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("BPAnalyzer", GetBuildfileValue("BPAmsi"), objInstParm)

  If GetBuildfileValue("SetupBPAnalyzerStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call CopyBPA("SQL2012BPA")
  Call CopyBPA("Microsoft")

  Call SetBuildfileValue("SetupBPAnalyzerStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CopyBPA(strModel)
  Call DebugLog("CopyBPA: " & strModel)

  strPathOld        = strDirSys & "\system32\BestPractices\v1.0\Models\" & strModel
  Select Case True
    Case strOSVersion < "6.2"
      ' Nothing
    Case Not objFSO.FolderExists(strPathOld)
      ' Nothing
    Case Else
      strPathNew    = strDirSysData & "\Microsoft\Microsoft Baseline Configuration Analyzer 2\Models\" & strModel
      Call SetupFolder(strPathNew)
      strDebugMsg1  = "Source folder: " & strPathOld
      Set objFolder = objFSO.GetFolder(strPathOld)
      objFolder.Copy strPathNew, True
  End Select

End Sub


Sub SetupKB2781514()
  Call SetProcessId("4SAB", "Installing KB2781514")

  Call RunInstall("KB2781514", GetBuildfileValue("KB2781514exe"), "")

  Call ProcessEnd("")

End Sub


Sub SetupJavaDBC()
  Call SetProcessId("4SB", "Install Java DBC Driver")
  Dim objInstParm
  Dim strJavaDir, strJavaExe

  strJavaExe        = GetBuildfileValue("Javaexe")
  strJavaDir        = strDirProg & "\JavaDBC\" & Left(strJavaexe, 11) & "\" & Mid(strJavaexe, Len(strJavaexe) - 6, 3) & "\"
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call SetXMLParm(objInstParm, "SetupOption",  "Extract")
  Call SetXMLParm(objInstParm, "ParmExtract",  "/Auto")
  Call SetXMLParm(objInstParm, "InstOption",   "None")
  Call SetXMLParm(objInstParm, "InstTarget",   strDirProg)
  If UCase(strJavaExe) < UCase("sqljdbc_4") Then
    Call SetXMLParm(objInstParm, "MenuOption", "Build")
    Call SetXMLParm(objInstParm, "MenuError",  "Ignore")
    Call SetXMLParm(objInstParm, "MenuSource", strJavaDir & "help\default.htm")
    Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLDocs)
  End If
  Call RunInstall("JavaDBC", strJavaExe, objInstParm)

  If GetBuildfileValue("SetupJavaDBCStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case Else
      Call DebugLog(" Configuring Java DBC Driver")
      strPathOld    = strJavaDir & "xa\"
      strDebugMsg1  = "Source folder: " & strPathOld
      Set objFile   = objFSO.GetFile(strPathOld & strFileArc & "\sqljdbc_xa.dll")
      strPathNew    = strSQLBinRoot & "\" & objFile.name
      strDebugMsg2  = "Target folder: " & strPathNew
      objFile.Copy strPathNew, True
      If strActionSQLDB <> "ADDNODE" Then
        Call SetXMLParm(objInstParm, "PathMain",       strPathOld)
        Call SetXMLParm(objInstParm, "ParmXtra",       "-V 17")
        Call SetXMLParm(objInstParm, "LogXtra",        "xa_install")
        Call RunInstall("JavaDBC",   "xa_install.sql", objInstParm)
      End If
  End Select

  Call SetBuildfileValue("SetupJavaDBCStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub InstallDB2OLE()
  Call SetProcessId("4SC", "Install DB2 OLE Provider")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SCA"
      ' Nothing
    Case strSQLVersion >= "SQL2012"
      ' Nothing
    Case Else
      Call SetupDB2OLEV3
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SCB"
      ' Nothing
    Case strSQLVersion <= "SQL2008R2"
      ' Nothing
    Case strSQLVersion = "SQL2012"
      Call SetupDB2OLE("4.0")
    Case Else
      Call SetupDB2OLE("5.0")
  End Select

  Call SetProcessId("4SCZ", "Install DB2 OLE Provider" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupDB2OLEV3()
  Call SetProcessId("4SCA", "Install DB2 OLE Provider V3")
  Dim objInstParm
  Dim strInstTarget, strMenuFolder

  strInstTarget     = strPathTemp & "DB2OLE"
  strMenuFolder     = "Microsoft OLE DB Provider for DB2"
  Call SetXMLParm(objInstParm, "SetupOption", "Extract")
  Call SetXMLParm(objInstParm, "ParmExtract", "/Auto")
  Call SetXMLParm(objInstParm, "InstFile",    "Setup.exe")
  Call SetXMLParm(objInstParm, "ParmXtra",    "/S """ & strInstTarget & "\Support\HISDB2Config.xml"" /INSTALLDIR """ & strDirProg & "\DB2OLEDB"" /ADDLOCAL ALL")
  Call SetXMLParm(objInstParm, "ParmLog",     "/L")
  Call SetXMLParm(objInstParm, "ParmReboot",  "")
  Call SetXMLParm(objInstParm, "ParmSilent",  "/quiet")
  Call SetXMLParm(objInstParm, "MenuOption", "Move")
  Call SetXMLParm(objInstParm, "MenuSource", strAllUserProf & "\" & strMenuPrograms & "\" & strMenuFolder)
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuConfigTools & "\" & strMenuFolder)
  Call RunInstall("DB2OLE", GetBuildfileValue("DB2exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupDB2OLE(strVersion)
  Call SetProcessId("4SCB", "Install DB2 OLE Provider " & strVersion)
  Dim objInstParm
  Dim strMenuFolder

  strMenuFolder     = "Microsoft OLE DB Provider for DB2 Version " & strVersion
  Call SetXMLParm(objInstParm, "MenuOption", "Move")
  Call SetXMLParm(objInstParm, "MenuSource", strAllUserProf & "\" & strMenuPrograms & "\" & strMenuFolder)
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuConfigTools & "\" & strMenuFolder)
  Call RunInstall("DB2OLE", GetBuildfileValue("DB2OLEmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupCacheManager()
  Call SetProcessId("4SD", "Install SQL Cache Manager")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   "SQLServerCacheManager.application")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "SQLServerCacheManager")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuPerfTools)
  Call RunInstall("CacheManager", GetBuildfileValue("CacheManagerZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupIntViewer()
  Call SetProcessId("4SE", "Install SQL Internals Viewer")

  Call RunInstall("IntViewer", GetBuildfileValue("IntViewermsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupMDS()
  Call SetProcessId("4SF", "Setup Master Data Services")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SFA"
      ' Nothing
    Case Instr("SQL2008R2", strSQLVersion) = 0
      ' Nothing
    Case Else
      Call InstallMDS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SFB"
      ' Nothing
    Case Else
'      Call SetMDSPerms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SFC"
      ' Nothing
    Case Else
'      Call CreateMDSDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SFD"
      ' Nothing
    Case Else
'      Call SetupMDSWebApp()
  End Select

  Call SetBuildfileValue("SetupMDSStatus", strStatusComplete)
  Call SetProcessId("4SFZ", " Setup Master Data Services" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub InstallMDS()
  Call SetProcessId("4SFA", "Install Master Data Services")

  Dim objInstParm

  Call SetXMLParm(objInstParm,"PathAlt", strPathSQLMedia & "MasterDataServices\x64\1033_ENU\")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("MDS", "MasterDataServices.msi", objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetMDSPerms()
  Call SetProcessId("4SFB", "Set MDS Permissions")

  Dim strMDSGroup, strMDSWebPath

  strMDSGroup       = "MDS_ServiceAccounts"
  strMDSWebPath     = strPathMDS & "\WebApplication\"

  strCmd            = "NET LOCALGROUP """ & strGroupDistComUsers & """ """ & FormatAccount(strMDSAccount) & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  strCmd            = "NET LOCALGROUP """ & strGroupIISIUsers    & """ """ & FormatAccount(strMDSAccount) & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strCmd            = "NET LOCALGROUP """ & strMDSGroup & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  strCmd            = "NET LOCALGROUP """ & strMDSGroup & """ """ & FormatAccount(strMDSAccount) & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  strCmd            = "NET LOCALGROUP """ & strMDSGroup & """ """ & FormatAccount(GetBuildfileValue("UserAccount")) & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strPath           = strMDSWebPath & "web.config"
  strCmd            = """" & strPath & """ /T /C /E /G """ & strMDSGroup & """:R"
  Call RunCacls(strCmd)

  strPath           = strMDSWebPath & "Logs"
  Call SetupFolder(strPath)
  strCmd            = """" & strPath & """ /T /C /E /G """ & strMDSGroup & """:W"
  Call RunCacls(strCmd)

  Call SetRegPerm("HKEY_LOCAL_MACHINE\" & Mid(strHKLMSQL, 6),            strMDSGroup,      "F")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreateMDSDB ()
  Call SetProcessId("4SFC", "Create MDSDB")

  strCmd            = strCmdPS & " -ExecutionPolicy Bypass -File """ & strPathFBScripts & "Set-MDSDB.ps1"" -DLLPath """ & strPathMDS & "\WebApplication\bin\Microsoft.MasterDataServices.Configuration.dll"" -instance """ & strServInst & """ -dbName """ & GetBuildfileValue("MDSDB") & """ -Account """ & GetBuildfileValue("UserAccount") & """"
  Call Util_RunExec(strCmd, "", "", 0)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMDSWebApp()
  Call SetProcessId("4SFD", "Setup MDS Web Application")
  Dim strMDSSite

  strPath           = strDirSys & "\system32\inetsrv\"
  strPathNew        = GetBuildfileValue("IISRoot")
  strMDSSite        = GetBuildfileValue("MDSSite")

  Call DebugLog("Add MDS Site: " & strMDSSite)
  strCmd            = """" & strPath & "APPCMD.EXE"" ADD SITE /name:""" & strMDSSite & """ /bindings:""" & GetBuildfileValue("HTTP") & ":/*:" & GetBuildfileValue("MDSPort") & """ /physicalPath:""" & strPathNew & "\" & strMDSSite & """ "
  Call Util_RunExec(strCmd, "", "", 183)
  
  Call DebugLog("Add MDSAPP application")
  strCmd            = """" & strPath & "APPCMD.EXE"" ADD APP /site.name:""" & strMDSSite & """ /path:""" & "/MDSApp" & """ /physicalPath:""" & strPathNew & "\MDSApp" & """ "
  Call Util_RunExec(strCmd, "", "", 183)

  Call DebugLog("Add MDS Pool")
  strCmd            = """" & strPath & "APPCMD.EXE"" ADD APPPOOL /name:""MDSPool"" "
  Call Util_RunExec(strCmd, "", "", 183)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPerfDash()
  Call SetProcessId("4SG", "Install SQL Performance Dashboard")
  Dim objInstParm
  Dim strDirPerfDash, strPerfDashmsi

  strPerfDashmsi    = GetBuildfileValue("PerfDashmsi")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("PerfDash", strPerfDashmsi, objInstParm)

  If GetBuildfileValue("SetupPerfDashStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Installing Performance Dashboard reports")

  Select Case True
    Case UCase(Left(strPerfDashmsi, 13)) = UCase("SQLServer2012")
      strDirPerfDash = strVolSys & Mid(strDirProgX86, 2) & "\110\Tools\Performance Dashboard\" 
    Case UCase(Left(strPerfDashmsi, 13)) = UCase("SQLServer2005")
      strDirPerfDash = strVolSys & Mid(strDirProgX86, 2) & "\90\Tools\PerformanceDashboard\" 
    Case Else
      strDirPerfDash = strVolSys & Mid(strPathSSMSX86, 2) & "Performance Dashboard\"
  End Select

  If strSQLVersion < "SQL2012" Then
    strCmd          = "CSCRIPT """ & strPathFBScripts & "ReplaceText.vbs"" """ & strDirPerfDash & "setup.sql"" ""cpu_ticks / convert(float, cpu_ticks_in_ms)"" ""ms_ticks"""
    Call Util_RunExec(strCmd, "", "", 0)
  End If

  strPathOld        = strDirPerfDash & "*.RDL"
  strPathNew        = strDirDBA & "\SQL Server Management Studio\Custom Reports"
  strDebugMsg1      = "Source folder: " & strPathOld
  strDebugMsg2      = "Target folder: " & strPathNew
  objFSO.CopyFile strPathOld, strPathNew, True

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Not objFSO.FolderExists(strDirPerfDash)
      ' Nothing
    Case Else
      Call SetXMLParm(objInstParm, "PathMain",  strDirPerfDash)
      Call SetXMLParm(objInstParm, "LogXtra",   "setup")
      Call RunInstall("PerfDash",  "setup.sql", objInstParm)
  End Select

  If strSQLVersion >= "SQL2012" Then
    strCmd           = "WMIC process WHERE ""CommandLine LIKE '%Performance Dashboard%' AND CommandLine LIKE '%readme.txt%' AND Name LIKE '%notepad%'"" CALL terminate"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  Call SetBuildfileValue("SetupPerfDashStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSystemViews()
  Call SetProcessId("4SH", "Install SQL System Views Map")
  Dim objInstParm
  Dim strSystemViewsPDF

  strSystemViewsPDF = GetBuildfileValue("SystemViewsPDF")
  Call SetXMLParm(objInstParm, "SetupOption", "Copy")
  Call SetXMLParm(objInstParm, "InstOption",  "Menu")
  Call SetXMLParm(objInstParm, "InstTarget",  strPathBOL & "Books")
  Call SetXMLParm(objInstParm, "MenuOption",  "Build")
  Call SetXMLParm(objInstParm, "MenuName",    strSQLVersion & " System Views")
  Call SetXMLParm(objInstParm, "MenuPath",    strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLDocs)
  Call RunInstall("SystemViews", strSystemViewsPDF, objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSQLNS()
  Call SetProcessId("4SI", "Install Notification Services")
  Dim objInstParm
  Dim strMenuSQL2005Flag, strSQL2005Path

  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("SQLNS", GetBuildfileValue("SQLNSmsi"), objInstParm)

  If GetBuildfileValue("SetupSQLNSStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Move menus to " & strSQLVersion & " container")
  strSQL2005Path    = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL2005
  strMenuSQL2005Flag  = GetBuildfileValue("MenuSQL2005Flag")
  strPathOld        = strSQL2005Path & "\" & strMenuConfigTools
  strPath           = strPathOld & "\" & strMenuSQLNS & ".lnk"
  Select Case True
    Case Not objFSO.FolderExists(strPathOld)
      ' Nothing
    Case Not objFSO.FileExists(strPath)
      ' Nothing
    Case Else
      Set objFile   = objFSO.GetFile(strPath)
      strPathNew    = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuConfigTools
      Call SetupFolder(strPathNew)
      If strMenuSQL2005Flag <> "Y" Then
        objFile.Copy strPathNew & "\" & objFile.Name, True
        objFile.Delete(1)
      End If
  End Select

  Call SetBuildfileValue("SetupSQLNSStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupStreamInsight()
  Call SetProcessId("4SJ", "Install Stream Insight")
  Dim objInstParm
  Dim strInstFile, strInstParm, strInstStream, strStreamInsightPID

  strInstStream       = GetBuildfileValue("InstStream")
  strStreamInsightPID = GetBuildfileValue("StreamInsightPID")

  Select Case True
    Case strtype = "CLIENT" 
      strInstFile   = "StreamInsightClient.msi"
      strInstParm   = "IACCEPTLICENSETERMS=YES "
    Case Else
      strInstFile   = "StreamInsight.msi"
      strInstParm   = "IACCEPTLICENSETERMS=YES INSTANCENAME=""" & strInstStream & """ CREATESERVICE=1 "
  End Select
  If strStreamInsightPID <> "" Then
    strInstParm     = strInstParm & " PRODUCTKEY=""" & strStreamInsightPID & """ "
  End If

  Call SetXMLParm(objInstParm, "ParmXtra",   strInstParm)
  Call SetXMLParm(objInstParm, "PathAlt",    strPathSQLMedia & "StreamInsight\" & strFileArc & "\1033_ENU\")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("StreamInsight", strInstFile, objInstParm)

  If GetBuildfileValue("SetupStreamInsightStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Add DBA Group as StreamInsight Administrator")
  strCmd            = "NET LOCALGROUP ""StreamInsightUsers$" & strInstStream & """ /ADD"
  Call Util_RunExec(strCmd, "", "", 2)
  strCmd            = "NET LOCALGROUP ""StreamInsightUsers$" & strInstStream & """ """ & strGroupDBA & """ /ADD"
  Call Util_RunExec(strCmd, "", "", 2)

  Call SetBuildfileValue("SetupStreamInsightStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSamples()
  Call SetProcessId("4SK", "Install SQL Server Sample Databases")

  Select Case True
    Case strSQLVersion = "SQL2005"
      Call SetupSamplesSQL2005()
    Case strSQLVersion = "SQL2008"
      Call SetupSamplesSQL2008()
    Case strSQLVersion = "SQL2008R2"
      Call SetupSamplesSQL2008R2()
    Case strSQLVersion = "SQL2012"
      Call SetupSamplesSQL2012()
    Case strSQLVersion = "SQL2014"
      Call SetupSamplesSQL2014()
    Case strSQLVersion = "SQL2016"
      Call SetupSamplesSQL2016()
  End Select

  Call SetProcessId("4SKZ", " Install SQL Server Sample Databases" & strStatusComplete)
  Call ProcessEnd("")

End Sub



Sub SetupSamplesSQL2005()
  Call SetProcessId("4SKA", "Install SQL 2005 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSamplesSQL2008()
  Call SetProcessId("4SKA", "Install SQL 2008 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSamplesSQL2008R2()
  Call SetProcessId("4SKA", "Install SQL 2008R2 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSamplesSQL2012()
  Call SetProcessId("4SKA", "Install SQL 2012 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSamplesSQL2014()
  Call SetProcessId("4SKA", "Install SQL 2014 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSamplesSQL2016()
  Call SetProcessId("4SKA", "Install SQL 2016 Sample Databases")

  Call RunInstall("Samples", GetBuildfileValue("Samplesmsi"), "")

  Call ProcessEnd("")

End Sub


Sub SetupSemantics()
  Call SetProcessId("4SL", "Install Semantic Search")
  Dim objInstParm
  Dim strFileArc, strInstArc

  Select Case True
    Case strProcArc = "X86"
      strFileArc    = "X86"
      strInstArc    = "32"
    Case strProcArc = "AMD64" And strWOWX86 = "TRUE"
      strFileArc    = "X86"
      strInstArc    = "32"
    Case Else
      strFileArc    = "X64"
      strInstArc    = "64"
  End Select

  strPathOld        = strDirProgSys & "\Microsoft Semantic Language Database\" & strSQLVersion & "." & strFileArc & "\"
  Call SetXMLParm(objInstParm, "PathMain",     strPathSQLMedia & strFileArc & "\Setup")
  Call SetXMLParm(objInstParm, "ParmXtra",     "SQLSEMLANGDB_" & strInstArc & "=""" & strPathOld & """ ")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("Semantics", "SemanticLanguageDatabase.msi", objInstParm)

  If GetBuildfileValue("SetupSemanticsStatus") <> strStatusProgress Then
    Exit Sub
  End If

  If strActionSQLDB <> "ADDNODE" Then
    Call DebugLog(" Attaching SemanticsDB")
    strPathNew      = strDirData & "\SemanticsDB\"
    Call SetupFolder(strPathNew)
    Set objFile     = objFSO.GetFile(strPathOld & "semanticsDB.mdf")
    objFile.Copy strPathNew & "semanticsDB.mdf"
    strPathNew      = strDirLog & "\"
    strDebugMsg2    = "Target: " & strPathNew
    Set objFile     = objFSO.GetFile(strPathOld & "semanticsdb_log.ldf")
    objFile.Copy strPathNew & "semanticsdb_log.ldf"
    strCmd          = "CREATE DATABASE [SemanticsDB] ON"
    strCmd          = strCmd & " (FILENAME=N'" & strDirData & "\SemanticsDB\semanticsDB.mdf') "
    strCmd          = strCmd & ",(FILENAME=N'" & strDirLog & "\semanticsdb_log.ldf') "
    strCmd          = strCmd & "FOR ATTACH "
    Call Util_ExecSQL(strCmdSQL & "-Q ", """" & strCmd & ";""", 1)
  End If

  strCmd            = "EXEC sp_fulltext_semantic_register_language_statistics_db @dbname=N'SemanticsDB' "
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)

  If strSQLVersion >= "SQL2014" Then
    strCmd          = "ALTER DATABASE [SemanticsDB] SET COMPATIBILITY_LEVEL=" & strSQLVersionNum
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)
    strCmd          = "ALTER DATABASE [SemanticsDB] SET DELAYED_DURABILITY=FORCED"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & ";""", 1)
  End If

  Call SetBuildfileValue("SetupSemanticsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDQ()
  Call SetProcessId("4SM", "Setup Data Quality Services")

  strDQSInstall     = "YES"

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SMA"
      ' Nothing
    Case Else
      Call InstallDQComponents()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SMB"
      ' Nothing
    Case Else
      Call SetupDQDatabase()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SMC"
      ' Nothing
    Case strDQSInstall <> "YES"
      ' Nothing
    Case Else
      Call MoveDQLogfile()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4SMD"
      ' Nothing
    Case strDQSInstall <> "YES"
      ' Nothing
    Case Else
      Call SetupDQSecurity()
  End Select

  Call SetProcessId("4SMZ", " Install Data Quality Services" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub InstallDQComponents()
  Call SetProcessId("4SMA", "Install Data Quality Components")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call SetXMLParm(objInstParm, "PathAlt",      strPathSQLMedia & strFileArc & "\Setup")
  Call SetXMLParm(objInstParm, "ParmXtra",     "INSTANCEID=""" & strInstance & """ ")
  Call RunInstall("DQ", "sql_dq.msi", objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupDQDatabase()
  Call SetProcessId("4SMB", "Install Data Quality Databases")
  Dim objInstParm
  Dim strDQAction, strDQSOptions

  Select Case True
    Case strActionSQLDB = "ADDNODE"
      Call MoveToNode(strClusterGroupSQL, "")
      strDQAction   = "-upgradedlls"
    Case Else
      strDQAction   = "-install"
  End Select
  strDQSOptions = strDQAction & " -instance """ & strInstance & """ -catalog ""DQS"" -password """ & GetBuildfileValue("DQPassword") & """ "
  If Instr("SQL2008 SQL2008R2", strSQLVersion) > 0 Then
    strDQSOptions   = strDQSOptions & " -kj"
  End If

  WScript.Sleep strWaitShort
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call SetXMLParm(objInstParm, "PathMain",     strSQLBinRoot)
  Call SetXMLParm(objInstParm, "ParmXtra",     strDQSOptions)
  Call SetXMLParm(objInstParm, "ParmLog",      "")
  Call SetXMLParm(objInstParm, "ParmReboot",   "")
  Call SetXMLParm(objInstParm, "ParmRetry",   "1")
  Call SetXMLParm(objInstParm, "ParmSilent",   "-silent")
  Call RunInstall("DQ", "DQSInstaller.exe", objInstParm)

  Call ProcessEnd("")

End Sub


Sub MoveDQLogfile()
  Call SetProcessId("4SMC", "Data Quality Services Log File")

  strPathOld        = GetBuildfileValue("PathFB") & "DQS_install.log"
  strPathAlt        = GetBuildfileValue("PathFBStart") & "\DQS_install.log"
  Select Case True
    Case objFSO.FileExists(strPathOld)
      strPath       = strPathOld
    Case objFSO.FileExists(strPathAlt)
      strPath       = strPathAlt
    Case Else
      strPath       = ""
  End Select

  If strPath <> "" Then
    strPath         = FormatFolder(strPath)
    strDebugMsg1    = "Source: " & strPath
    Set objFile     = objFSO.GetFile(strPath)
    strPathLog      = Replace(GetPathLog(""), """", "")
    strDebugMsg2    = "Target: " & strPathLog
    objFile.Copy strPathLog, True
    objFile.Delete(1)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDQSecurity()
  Call SetProcessId("4SMD", "Setup Data Quality Security")

  strCmd            = "CREATE USER [" & strGroupDBA & "] FOR LOGIN [" & strGroupDBA & "]"
  Call Util_ExecSQL(strCmdSQL & "-d ""DQS_MAIN"" -Q", """" & strCmd & ";""", 1)
  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='dqs_administrator', @MEMBERNAME='" & strGroupDBA & "'"
  Call Util_ExecSQL(strCmdSQL & "-d ""DQS_MAIN"" -Q", """" & strCmd & ";""", 0)

  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      strCmd        = "CREATE USER [" & strGroupDBANonSA & "] FOR LOGIN [" & strGroupDBANonSA & "]"
      Call Util_ExecSQL(strCmdSQL & "-d ""DQS_MAIN"" -Q", """" & strCmd & ";""", 1)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='dqs_kb_operator', @MEMBERNAME='" & strGroupDBANonSA & "'"
      Call Util_ExecSQL(strCmdSQL & "-d ""DQS_MAIN"" -Q", """" & strCmd & ";""", 0)
  End Select

  Call SetBuildfileValue("SetupDQStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDistributor()
  Call SetProcessId("4SN", "Install Replication Distributor")
  Dim objSQL, objSQLData
  Dim strDistributionInstalled, strDistributionDBInstalled, strDistDatabase, strDistPassword, strDirectory

  Set objSQL        = CreateObject("ADODB.Connection")
  Set objSQLData    = CreateObject("ADODB.Recordset")
  strDistDatabase   = GetBuildfileValue("DistDatabase")
  strDistPassword   = GetBuildfileValue("DistPassword")
  strPath           = strHKLMSQL & strInstRegSQL & "\Replication\WorkingDirectory"
  strDirectory      = objShell.RegRead(strPath)

  Call DebugLog("Check if Distributor already installed")
  objSQL.Provider   = "SQLOLEDB"
  objSQL.ConnectionString = "Server=" & strServInst & ";Database=master;Trusted_Connection=Yes;"
  objSQL.Open 
  strCmd            = "EXEC sp_get_distributor"
  Set objSQLData    = objSQL.Execute(strCmd)
  Do Until objSQLData.EOF
    strDistributionInstalled   = objSQLData.Fields("installed")
    strDistributionDBInstalled = objSQLData.Fields("distribution db installed")
    objSQLData.MoveNext
  Loop

  If strDistributionInstalled = "False" Then
    Call DebugLog("Add Distributor")
    strCmd          = "EXEC sp_adddistributor @distributor='" & strServInst & "'"
    strCmd          = strCmd & ",@password='" & strDistPassword & "'"
    Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & ";""", 0)
  End If

  If strDistributionDBInstalled = "False" Then
    Call DebugLog("Create Distribution DB")
    strPathNew      = strDirData & "\" & strDistDatabase
    Call SetupFolder(strPathNew)
    strCmd          = "EXEC sp_adddistributiondb @database='" & strDistDatabase & "'"
    strCmd          = strCmd & ",@security_mode=1"
    strCmd          = strCmd & ",@data_folder='" & strPathNew & "'"
    strCmd          = strCmd & ",@log_folder='" & strDirLog & "'"
    Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & ";""", 0)
  End If

  Call DebugLog("Setup Distribution publishing")
  strCmd            = "EXEC sp_adddistpublisher @publisher='" & strServInst & "'"
  strCmd            = strCmd & ",distribution_db='" & strDistDatabase & "'"
  strCmd            = strCmd & ",@security_mode=1"
  strCmd            = strCmd & ",@password='" & strDistPassword & "'"
  strCmd            = strCmd & ",@working_directory='" & strDirectory & "'"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strDistDatabase & """ -Q", """" & strCmd & ";""", 1)

  objSQL.Close
  Set objSQL        = Nothing
  Set objSQLData    = Nothing

  Call SetBuildfileValue("SetupDistributorStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupToolsXtras()
  Call SetProcessId("4T", "Tools Extras")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TA"
      ' Nothing
    Case GetBuildfileValue("SetupABE") <> "YES"
      ' Nothing
    Case Else
      Call SetupABE()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TB"
      ' Nothing
    Case GetBuildfileValue("SetupXEvents") <> "YES"
      ' Nothing
    Case Else
      Call SetupXEvents()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TC"
      ' Nothing
    Case GetBuildfileValue("SetupPDFReader") <> "YES"
      ' Nothing
    Case Else
      Call SetupPDFReader()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TD"
      ' Nothing
    Case GetBuildfileValue("SetupProcExp") <> "YES"
      ' Nothing
    Case Else
      Call SetupProcExp()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TE"
      ' Nothing
    Case GetBuildfileValue("SetupProcMon") <> "YES"
      ' Nothing
    Case Else
      Call SetupProcMon()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TF"
      ' Nothing
    Case GetBuildfileValue("SetupRMLTools") <> "YES"
      ' Nothing
    Case Not Checkstatus("ReportViewer")
      Call SetBuildfileValue("SetupRMLToolsStatus", strStatusBypassed)
    Case Else
      Call SetupRMLTools()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TGZ"
      ' Nothing
    Case GetBuildfileValue("SetupSQLNexus") <> "YES"
      ' Nothing
    Case Not Checkstatus("ReportViewer")
      Call SetBuildfileValue("SetupSQLNexusStatus", strStatusBypassed)
    Case Else
      Call SetupSQLNexus()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TH"
      ' Nothing
    Case GetBuildfileValue("SetupTrouble") <> "YES"
      ' Nothing
    Case Else
      Call SetupTrouble()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TI"
      ' Nothing
    Case GetBuildfileValue("SetupXMLNotepad") <> "YES"
      ' Nothing
    Case Else
      Call SetupXMLNotepad()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TJ"
      ' Nothing
    Case GetBuildfileValue("SetupPlanExplorer") <> "YES"
      ' Nothing
    Case Else
      Call SetupPlanExplorer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TK"
      ' Nothing
    Case GetBuildfileValue("SetupPlanExpAddin") <> "YES"
      ' Nothing
    Case Else
      Call SetupPlanExpAddin()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "4TL"
      ' Nothing
    Case GetBuildfileValue("SetupZoomIt") <> "YES"
      ' Nothing
    Case Else
      Call SetupZoomIt()
  End Select

  Call SetProcessId("4TZ", "Tools Extras" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupABE()
  Call SetProcessId("4TA", "Install Windows Access Based Enumeration")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "MSIAutoOS",    "5.2")
  Call SetXMLParm(objInstParm, "SetupOption",  "Copy")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("ABE", GetBuildfileValue("ABEmsi"), objInstParm)

  If GetBuildfileValue("SetupABEStatus") <>  strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Enable ABE")
  strCmd            = "ABECMD /enable /all"
  Call Util_RunExec(strCmd, "", "", 0)

  Call SetBuildfileValue("SetupABEStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupXEvents()
  Call SetProcessId("4TB", "Install Extended Events Manager")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmXtra",  "ALLUSERS=1")
  Call RunInstall("XEvents", GetBuildfileValue("XEventsmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupPDFReader()
  Call SetProcessId("4TC", "Install PDF Reader")
  Dim objInstParm
  Dim strPDFPath, strPDFReg

  Call SetXMLParm(objInstParm, "ParmLog",    "")
  Call SetXMLParm(objInstParm, "ParmReboot", "")
  Call SetXMLParm(objInstParm, "ParmSilent", "/S")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("PDFReader", GetBuildfileValue("PDFexe"), objInstParm)

  If GetBuildfileValue("SetupPDFReaderStatus") <>  strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Register Sumatra as PDF Reader application")
  strPDFPath        = strDirProgSysX86 & "\SumatraPDF\" & strPDFreg & ".exe"
  If objFSO.FileExists(strPDFPath) Then
    Call DebugLog("Set Sumatra as default PDF Reader")
    strPDFreg       = GetBuildfileValue("PDFreg")
    strCmd          = "%COMSPEC% /D /C FTYPE Sumatra=" & strPDFPath & " %1 %*"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
    strCmd          = "%COMSPEC% /D /C ASSOC .PDF=Sumatra"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End If

  Call SetBuildfileValue("SetupPDFReaderStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupProcExp()
  Call SetProcessId("4TD", "Install Process Explorer")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   GetBuildfileValue("ProcExpexe"))
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "Process Explorer")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuAdminTools)
  Call RunInstall("ProcExp", GetBuildfileValue("ProcExpZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupProcMon()
  Call SetProcessId("4TE", "Install Process Monitor")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   GetBuildfileValue("ProcMonexe"))
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "Process Monitor")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuAdminTools)
  Call RunInstall("ProcMon", GetBuildfileValue("ProcMonZip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupRMLTools()
  Call SetProcessId("4TF", "Install RML Tools")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "MenuOption",  "Move")
  Call SetXMLParm(objInstParm, "MenuSource",  strAllUserProf & "\" & strMenuPrograms & "\RML Utilities for SQL Server")
  Call SetXMLParm(objInstParm, "MenuPath",    strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuPerfTools & "\")
  Call RunInstall("RMLTools", GetBuildfileValue("RMLToolsmsi"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSQLNexus()
  Call SetProcessId("4TG", "Install SQL Nexus")

  Select Case True
    Case strSQLVersion = "SQL2005"
      Call SetupSQLNexusV3()
    Case Else
      Call SetupSQLNexusV4Plus()
  End Select

  Call ProcessEnd("")

End Sub


Sub SetupSQLNexusV3()
  Call SetProcessId("4TGA", "Install SQL Nexus V3")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   "sqlnexus.exe")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "SQL Nexus")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuPerfTools)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("SQLNexus", GetBuildfileValue("SQLNexuszip"), objInstParm)

  If GetBuildfileValue("SetupSQLNexusStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Configuring SQL Nexus Reports")
  strCmd      = "CSCRIPT """ & strPathFBScripts & "ReplaceText.vbs"" """ & Left(strPathInst, InstrRev(strPathInst, "\")) & "Reports\Realtime - Server Status.rdl"" ""<DataField>cpu_ticks_in_ms</DataField>"" ""<DataField>ms_ticks</DataField>"""
  Call Util_RunExec(strCmd, "", "", 0)

  Call SetBuildfileValue("SetupSQLNexusStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLNexusV4Plus()
  Call SetProcessId("4TGB", "Install SQL Nexus V4 Plus")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "None")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   "sqlnexus.exe")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "SQL Nexus")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuPerfTools)
  Call RunInstall("SQLNexus", GetBuildfileValue("SQLNexuszip"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupTrouble()
  Call SetProcessId("4TH", "Install SQL Troubleshooting Guide")
  Dim objInstParm
  Dim strTroubleFile

  strTroubleFile    = GetBuildfileValue("AccidentalDBAzip")
  Select Case True
    Case strTroubleFile <> ""
      Call SetXMLParm(objInstParm, "InstFile", Left(strTroubleFile, InstrRev(strTroubleFile, ".")) & "PDF")
    Case Else
      strTroubleFile = GetBuildfileValue("TroublePDF")  
  End Select

  Call SetXMLParm(objInstParm, "InstOption", "None")
  Call SetXMLParm(objInstParm, "InstTarget", strPathBOL & "Books")
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "Troubleshoting Performance Problems")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLDocs)
  Call RunInstall("Trouble", strTroubleFile, objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupXMLNotepad()
  Call SetProcessId("4TI", "Install XML Notepad")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "MenuOption", "Remove")
  Call SetXMLParm(objInstParm, "MenuName",   "XML Notepad 2007")
  Call SetXMLParm(objInstParm, "MenuPath",   strUserDTop)
  Call RunInstall("XMLNotepad", GetBuildfileValue("XMLmsi"), objInstParm)

  Call ProcessEnd("")

End Sub

Sub SetupPlanExplorer()
  Call SetProcessId("4TJ", "Install Plan Explorer")
' Code contributed by Brian Davis https://twitter.com/brian78
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmLog",    "/l")
  Call SetXMLParm(objInstParm, "ParmReboot", "")
  Call SetXMLParm(objInstParm, "ParmSilent", "/s")
  Call SetXMLParm(objInstParm, "MenuOption", "Remove")
  Call SetXMLParm(objInstParm, "MenuName",   "SQL Sentry Plan Explorer")
  Call SetXMLParm(objInstParm, "MenuPath",   strUserDTop)
  Call RunInstall("PlanExplorer", GetBuildfileValue("PlanExpexe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupPlanExpAddin()
  Call SetProcessId("4TK", "Install Plan Explorer SSMS Addin")
' Code contributed by Brian Davis https://twitter.com/brian78

  Call RunInstall("PlanExpAddin", GetBuildfileValue("PlanExpAddinmsi"), "")

  Call ProcessEnd("")
End Sub


Sub SetupZoomIt()
  Call SetProcessId("4TL", "Install Zoom It")
' Code contributed by Brian Davis https://twitter.com/brian78
  Dim objInstParm

  Call SetXMLParm(objInstParm, "InstOption", "Menu")
  Call SetXMLParm(objInstParm, "InstTarget", strDirProgSys)
  Call SetXMLParm(objInstParm, "InstFile",   GetBuildfileValue("ZoomItExe"))
  Call SetXMLParm(objInstParm, "MenuOption", "Build")
  Call SetXMLParm(objInstParm, "MenuName",   "ZoomIt")
  Call SetXMLParm(objInstParm, "MenuPath",   strAllUserProf & "\" & strMenuPrograms & "\" & strMenuAdminTools)
  Call RunInstall("ZoomIt", GetBuildfileValue("ZoomItZip"), objInstParm)

  Call ProcessEnd("")

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


Sub ProcessUserRSRegistry(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessUserRSRegistry: ")

  Call SetRegPerm(strKey & strSid & "\Software\Microsoft\Avalon.Graphics", GetBuildfileValue("NTAuthEveryone"), "R")

End Sub