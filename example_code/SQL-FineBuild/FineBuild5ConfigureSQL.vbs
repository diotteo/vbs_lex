''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild5ConfigureSQL.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      Configures a SQL Server instance
'
'  Author:       Ed Vassie
'
'  Change History
'  Version  Author        Date         Description
'  2.1      Ed Vassie     14 sep 2009  Initial version for SQL Server 2008 R2
'  2.0      Ed Vassie     02 Jul 2008  Initial SQL Server 2008 version for FineBuild v2.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colPrcEnvVars
Dim objAccount, objApp, objAutoUpdate, objFile, strFileArc, objFolder, objSubFolder, objFSO, objShell, objShortcut, objTarget, objWMI, objWMIReg
Dim intIdx, intMaxDop, intProcNum, intServerLen
Dim strAction, strActionAO, strActionSQLAS, strActionSQLDB, strActionSQLRS, strAgentJobHistory, strAgentMaxHistory, strAllUserDTop, strAllUserProf, strAnyKey, strAuditLevel, strBPEFile, strAgtAccount, strASAccount, strBuiltinDom
Dim strClusIPAddress, strClusterAction, strClusterName, strClusterNameRS, strClusterNameSQL, strClusterIPV4DB, strClusterIPV6DB, strClusterGroupAS, strClusterGroupRS, strClusterGroupSQL, strCmd, strCmdPS, strCmdRS, strCmdSQL, strCmdshellAccount, strCmdshellPassword
Dim strDBA_DB, strDBAEmail, strDBMailOK, strDBMailProfile, strDBOwnerAccount, strDfltProf, strDefaultUser, strDomain, strCltAccount, strDTCClusterRes
Dim strDirBackup, strDirBPE, strDirData, strDirDataFT, strDirDBA, strDirLog, strDirLogTemp, strDirProg, strDirProgX86, strDirProgSys, strDirProgSysX86, strDirSys, strDirSysData, strDirSystemDataBackup, strDirSystemDataShared, strDirTempData, strGroupDBA, strGroupDBAAlt, strGroupDBANonSA, strGroupDBANonSAAlt, strGroupDistComUsers, strLocalAdmin, strManagementDW, strManagementInstance, strManagementServer, strManagementServerList, strManagementServerName, strOSName, strOSType, strOSVersion, strGroupAO
Dim strEdition, strEditionEnt, strErrMsg, strMailServer, strMailServerType, strMDSDB, strSetCLREnabled, strSetCostThreshold, strSetMemOptHybridBP, strSetMemOptTempdb, strSQLMaxMemory, strSQLMinMemory, strSetOptimizeForAdHocWorkloads, strSetRemoteAdminConnections, strSetRemoteProcTrans, strSetxpCmdshell, strNTAuthAccount, strNTAuthOSName, strNumErrorLogs, strNumLogins, strNumTF, strProfileName, strSetupDQ, strSetupLog, strSetupSSISCluster, strSetupStretch, strSPLevel, strSPCULevel
Dim strInstance, strInstADHelper, strInstAgent, strInstAnal, strInstAS, strInstDTCClusterRes, strInstFT, strInstIS, strSetupMDS, strInstNode, strSetupSQLAS, strSetupSQLASCluster, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLDBFT, strSetupSQLIS, strSetupSQLRS, strSetupSQLRSCluster, strSetupSQLTools, strInstRegSQL, strInstRS, strInstRSDir, strInstRSHost, strInstRSURL, strInstSQL, strMainInstance, strMenuAccessories, strMenuAdminTools, strMenuBOL, strMenuPerfTools, strMenuPrograms, strMenuSQL, strMenuSQL2005Flag, strMenuSQLDocs, strMenuSQLRS, strMenuSSMS, strMenuSystem
Dim strLocalDomain, strHKU, strHKLMSQL, strHTTP, strInstLog, strPath, strPathCScript, strPathFB, strPathFBScripts, strPathLog, strPathNew, strPathOld, strPathTemp, strServer, strServInst, strType
Dim strMDWAccount, strMDWPassword
Dim strRegSSIS, strRegSSISSetup, strResSuffix, strRoleDBANonSA, strRsExecAccount, strRsExecPassword
Dim strRSDBName, strRSHost, strRSInstallMode, strRSVersionNum, strsaName, strsaPwd, strSIDDistComUsers, strSqlAccount, strSqlPassword, strBrAccount, strBrPassword, strExtSvcAccount, strExtSvcPassword, strFtAccount, strFtPassword, strIsAccount, strIsPassword, strIsSvcStartuptype, strIISRoot, strSQLList, strSQLVersion, strSQLVersionNet, strSQLVersionNum
Dim strSetupAlwaysOn, strSetupAnalytics, strSetupDBMail, strSetupDRUClt, strSetupSnapshot, strSetupSQLMail, strSetupSQLNS, strSetupSSISDB, strSetupTempDb, strSSISDB, strSSISPassword

Dim strSqlBrowserStartup, strSQLEmail, strSQLMailOK, strSQLOperator, strSQLTempdbFileCount, strtempdbFile, strtempdbLogFile, strTCPPort, strTCPPortAS, strTCPPortDAC, strTCPPortRS, strUserDNSDomain, strUserName, strWaitLong, strWaitMed, strWaitShort, strWriterSvcStartupType

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "5TZ" ' 5UA to 5UZ reserved for User Configuration routine
      ' Nothing
    Case Else
      Call ProcessConfiguration()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "5TZ"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("Instance Configuration complete", 2, "Instance Configuration" ,64)
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
      Call FBLog(" Instance Configuration failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBUtils.vbs"
  Include "FBManageCluster.vbs"
  Include "FBManageInstall.vbs"
  Include "FBManageService.vbs"
  Call SetProcessIdCode("FB5C")

  Set objApp        = CreateObject ("Shell.Application")
  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")

  strHKU            = &H80000003
  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAgentJobHistory  = GetBuildfileValue("AgentJobHistory")
  strAgentMaxHistory  = GetBuildfileValue("AgentMaxHistory")
  strAgtAccount     = GetBuildfileValue("AgtAccount")
  strAction         = GetBuildfileValue("Action")
  strActionAO       = GetBuildfileValue("ActionAO")
  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strActionSQLRS    = GetBuildfileValue("ActionSQLRS")
  strAllUserDTop    = GetBuildfileValue("AllUserDTop")
  strAllUserProf    = GetBuildfileValue("AllUserProf")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strAsAccount      = GetBuildfileValue("AsAccount")
  strAuditLevel     = GetBuildfileValue("AuditLevel")
  strBPEFile        = GetBuildfileValue("BPEFile")
  strBrAccount      = GetBuildfileValue("SqlBrowserAccount")
  strBrPassword     = GetBuildfileValue("SqlBrowserPassword")
  strBuiltinDom     = GetBuildfileValue("BuiltinDom")
  strClusIPAddress  = GetBuildfileValue("ClusIPAddress")
  strClusterGroupAS = GetBuildfileValue("ClusterGroupAS")
  strClusterGroupRS = GetBuildfileValue("ClusterGroupRS")
  strClusterGroupSQL  = GetBuildfileValue("ClusterGroupSQL")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterNameRS  = GetBuildfileValue("ClusterNameRS")
  strClusterIPV4DB  = GetBuildfileValue("ClusterIPV4DB")
  strClusterIPV6DB  = GetBuildfileValue("ClusterIPV6DB")
  strClusterNameSQL = GetBuildfileValue("ClusterNameSQL")
  strCmdshellAccount  = GetBuildfileValue("CmdShellAccount")
  strCmdshellPassword = GetBuildfileValue("CmdShellPassword")
  strCmdPS          = GetBuildfileValue("CmdPS")
  strCmdRS          = GetBuildfileValue("CmdRS")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strDBA_DB         = GetBuildfileValue("DBA_DB")
  strDBAEmail       = GetBuildfileValue("DBAEmail")
  strDBMailOK       = ""
  strDBMailProfile  = GetBuildfileValue("DBMailProfile")
  strDBOwnerAccount = GetBuildfileValue("DBOwnerAccount")
  strDefaultUser    = GetBuildfileValue("DefaultUser")
  strDfltProf       = GetBuildfileValue("DfltProf")
  strDirBackup      = GetBuildfileValue("DirBackup")
  strDirBPE         = GetBuildfileValue("DirBPE")
  strDirData        = GetBuildfileValue("DirData")
  strDirDataFT      = GetBuildfileValue("DirDataFT")
  strDirDBA         = GetBuildfileValue("DirDBA")
  strDirLog         = GetBuildfileValue("DirLog")
  strDirLogTemp     = GetBuildfileValue("DirLogTemp")
  strDirProg        = GetBuildfileValue("DirProg")
  strDirProgX86     = GetBuildfileValue("DirProgX86")
  strDirProgSys     = GetBuildfileValue("DirProgSys")
  strDirProgSysX86  = GetBuildfileValue("DirProgSysX86")
  strDirSys         = GetBuildfileValue("DirSys")
  strDirSysData     = GetBuildfileValue("DirSysData")
  strDirSystemDataBackup = FormatFolder(GetBuildfileValue("DirSystemDataBackup"))
  strDirSystemDataShared = FormatFolder(GetBuildfileValue("DirSystemDataShared"))
  strDirTempData    = GetBuildfileValue("DirTemp")
  strDomain         = GetBuildfileValue("Domain")
  strCltAccount     = GetBuildfileValue("CltAccount")
  strDTCClusterRes  = GetBuildfileValue("DTCClusterRes")
  strEdition        = GetBuildfileValue("AuditEdition")
  strEditionEnt     = GetBuildfileValue("EditionEnt")
  strExtSvcAccount  = GetBuildfileValue("ExtSvcAccount")
  strExtSvcPassword = GetBuildfileValue("ExtSvcrPassword")
  strFileArc        = GetBuildfileValue("FileArc")
  strFTAccount      = GetBuildfileValue("FtAccount")
  strFtPassword     = GetBuildfileValue("FtPassword")
  strGroupAO        = GetBuildfileValue("GroupAO")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBAAlt    = GetBuildfileValue("GroupDBAAlt")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strGroupDBANonSAAlt  = GetBuildfileValue("GroupDBANonSAAlt")
  strGroupDistComUsers = GetBuildfileValue("GroupDistComUsers")
  strHTTP           = GetBuildfileValue("HTTP")
  strInstance       = GetBuildfileValue("Instance")
  strInstADHelper   = GetBuildfileValue("InstADHelper")
  strInstAgent      = GetBuildfileValue("InstAgent")
  strInstAnal       = GetBuildfileValue("InstAnal")
  strInstAS	    = GetBuildfileValue("InstAS")
  strInstDTCClusterRes = GetBuildfileValue("InstDTCClusterRes")
  strInstFT	    = GetBuildfileValue("InstFT")  
  strInstLog        = GetBuildfileValue("InstLog")
  strInstIS         = GetBuildfileValue("InstIS")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstRSDir      = GetBuildfileValue("InstRSDir")
  strInstRSHost     = GetBuildfileValue("InstRSHost") 
  strInstRSURL      = GetBuildfileValue("InstRSURL")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strIISRoot        = GetBuildfileValue("IISRoot")
  strIsAccount      = GetBuildfileValue("IsAccount")
  strIsPassword     = GetBuildfileValue("IsPassword")
  strIsSvcStartuptype = GetBuildfileValue("IsSvcStartupType")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strLocalDomain    = GetBuildfileValue("LocalDomain")
  strMailServer     = GetBuildfileValue("MailServer")
  strMailServerType = GetBuildfileValue("MailServerType")
  strMainInstance   = GetBuildfileValue("MainInstance")
  strManagementDW   = GetBuildfileValue("ManagementDW")
  strManagementInstance   = GetBuildfileValue("ManagementInstance")
  strManagementServer     = GetBuildfileValue("ManagementServer")
  strManagementServerName = GetBuildfileValue("ManagementServerName")
  intMaxDop         = GetBuildfileValue("SQLMaxDop")
  strMDSDB          = GetBuildfileValue("MDSDB")
  strMDWAccount     = GetBuildfileValue("MDWAccount")
  strMDWPassword    = GetBuildfileValue("MDWPassword")
  strMenuAccessories  = GetBuildfileValue("MenuAccessories")
  strMenuAdminTools = GetBuildfileValue("MenuAdminTools")
  strMenuPrograms   = GetBuildfileValue("MenuPrograms")
  strMenuBOL        = GetBuildfileValue("MenuBOL")
  strMenuPerfTools  = GetBuildfileValue("MenuPerfTools")
  strMenuSQL        = GetBuildfileValue("MenuSQL")
  strMenuSQL2005Flag  = GetBuildfileValue("MenuSQL2005Flag")
  strMenuSQLDocs    = GetBuildfileValue("MenuSQLDocs")
  strMenuSQLRS      = GetBuildfileValue("MenuSQLRS")
  strMenuSSMS       = GetBuildfileValue("MenuSSMS")
  strMenuSystem     = GetBuildfileValue("MenuSystem")
  strNTAuthAccount  = GetBuildfileValue("NTAuthAccount")
  strNTAuthOSName   = GetBuildfileValue("NTAuthOSName")
  strNumErrorLogs   = GetBuildfileValue("NumErrorLogs")
  strNumLogins      = GetBuildfileValue("NumLogins")
  strNumTF          = GetBuildfileValue("NumTF")
  strSetCLREnabled  = GetBuildfileValue("SetCLREnabled")
  strSetCostThreshold             = GetBuildfileValue("SetCostThreshold")
  strSetMemOptHybridBP            = GetBuildfileValue("SetMemOptHybridBP")
  strSetMemOptTempdb              = GetBuildfileValue("SetMemOptTempdb")
  strSQLMaxMemory   = GetBuildfileValue("SQLMaxMemory")
  strSQLMinMemory   = GetBuildfileValue("SQLMinMemory")
  strSetOptimizeForAdHocWorkloads = GetBuildfileValue("SetOptimizeForAdHocWorkloads")
  strSetRemoteAdminConnections    = GetBuildfileValue("SetRemoteAdminConnections")
  strSetRemoteProcTrans           = GetBuildfileValue("SetRemoteProcTrans")
  strSetxpCmdshell  = GetBuildfileValue("SetxpCmdshell")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathCScript    = GetBuildfileValue("PathCScript")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  intProcNum        = GetBuildfileValue("ProcNum")
  strProfileName    = GetBuildfileValue("ProfileName")
  strRegSSIS        = GetBuildfileValue("RegSSIS")
  strRegSSISSetup   = GetBuildfileValue("RegSSISSetup")
  strResSuffix      = GetBuildfileValue("ResSuffix")
  strRoleDBANonSA   = GetBuildfileValue("RoleDBANonSA")
  strRSDBName       = GetBuildfileValue("RSDBName")
  strRsExecAccount  = GetBuildfileValue("RsExecAccount")
  strRsExecPassword = GetBuildfileValue("RsExecPassword")
  strRSInstallMode  = GetBuildfileValue("RSInstallMode")
  strRSVersionNum   = GetBuildfileValue("RSVersionNum")
  strsaName         = GetBuildfileValue("saName")
  strsaPwd          = GetBuildfileValue("saPwd")
  strServer         = GetBuildfileValue("AuditServer")
  intServerLen      = Len(strServer)
  strServInst       = GetBuildfileValue("ServInst")
  strSetupAlwaysOn  = GetBuildfileValue("SetupAlwaysOn")
  strSetupAnalytics = GetBuildfileValue("SetupAnalytics")
  strSetupDBMail    = GetBuildfileValue("SetupDBMail")
  strSetupSQLMail   = GetBuildfileValue("SetupSQLMail")
  strSetupDQ        = GetBuildfileValue("SetupDQ")
  strSetupDRUClt    = GetBuildfileValue("SetupDRUClt")
  strSetupLog       = Ucase(objShell.ExpandEnvironmentStrings("%SQLLOGTXT%"))
  strSetupLog       = Left(strSetupLog, InStrRev(strSetupLog, "\"))
  strSetupMDS       = GetBuildfileValue("SetupMDS")
  strSetupSnapshot  = GetBuildfileValue("SetupSnapshot")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLASCluster = GetBuildfileValue("SetupSQLASCluster")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strSetupSQLDBFT   = GetBuildfileValue("SetupSQLDBFT")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLNS     = GetBuildfileValue("SetupSQLNS")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSetupSSISCluster = GetBuildfileValue("SetupSSISCluster")
  strSetupSSISDB    = GetBuildfileValue("SetupSSISDB")
  strSetupStretch   = GetBuildfileValue("SetupStretch")
  strSetupTempDb    = GetBuildfileValue("SetupTempDb")
  strSIDDistComUsers  = GetBuildfileValue("SIDDistComUsers")
  strSSISDB         = GetBuildfileValue("SSISDB")
  strSSISPassword   = GetBuildfileValue("SSISPassword")
  strSPLevel        = GetBuildfileValue("SPLevel")
  strSPCULevel      = GetBuildfileValue("SPCULevel")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSQLEmail       = GetBuildfileValue("SQLEmail")
  strSQLList        = GetBuildfileValue("SQLList")
  strSQLMailOK      = ""
  strSqlPassword    = GetBuildfileValue("SqlPassword")
  strSqlBrowserStartup  = GetBuildfileValue("SqlBrowserStartup")
  strSQLOperator    = GetBuildfileValue("SQLOperator")
  strSQLTempdbFileCount = GetBuildfileValue("SQLTempdbFileCount")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNet  = GetBuildfileValue("SQLVersionNet") 
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strType           = GetBuildfileValue("Type")
  strtempdbFile     = GetBuildfileValue("tempdbFile")
  strtempdbLogFile  = GetBuildfileValue("tempdbLogFile")
  strTCPPort        = GetBuildfileValue("TCPPort")
  strTCPPortAS      = GetBuildfileValue("TCPPortAS")
  strTCPPortDAC     = GetBuildfileValue("TCPPortDAC")
  strTCPPortRS      = GetBuildfileValue("TCPPortRS")
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  strUserName       = GetBuildfileValue("AuditUser")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitMed        = GetBuildfileValue("WaitMed")
  strWaitShort      = GetBuildfileValue("WaitShort")
  strWriterSvcStartupType = GetBuildfileValue("SqlWriterStartupType")

  strManagementServerList = " " & strServer & " " & strClusterNameSQL & " " & strGroupAO & " "

End Sub


Sub ProcessConfiguration()
  Call SetProcessId("5", strSQLVersion & " Configuration processing (FineBuild5ConfigureSQL.vbs)")

  Call SetUpdate("ON")
  If strActionSQLDB <> "ADDNODE" Then
    Call StartSQL()
  End If

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AZ"
      ' Nothing
    Case Else
      Call ConfigServices()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BZ"
      ' Nothing
    Case Else
      Call ConfigInstance()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CZ"
      ' Nothing
    Case Else
      Call ConfigAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5DZ"
      ' Nothing
    Case Else
      Call ConfigDBs()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EZ"
      ' Nothing
    Case Else
      Call ConfigManagement()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5FZ"
      ' Nothing
    Case Else
      Call ConfigWindows()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5GZ"
      ' Nothing
    Case Else
      Call ConfigTidy()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5H"
      ' Nothing
    Case GetBuildfileValue("SetupAutoConfig") <> "YES"
      ' Nothing
    Case Else
      Call SetupAutoConfig()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5ZZ"
      ' Nothing
    Case Else
      Call UserConfiguration()
  End Select

  Call SetUpdate("OFF")
  Call SetProcessId("5ZZ", " SQL Configuration processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigServices()
  Call SetProcessId("5A", "Services configuration")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AA"
      ' Nothing
    Case GetBuildfileValue("SetupDCom") <> "YES"
      ' Nothing
    Case strOSVersion < "6.0"
      Call SetBuildfileValue("SetupDComStatus", strStatusManual)
    Case Else
      Call ConfigureComSecurity()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AB"
      ' Nothing
    Case GetBuildfileValue("SetupNetwork") <> "YES"
      ' Nothing
    Case Else
      Call ConfigureSQLNetworkProtocols()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AC"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call ConfigServiceAC()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AD"
      ' Nothing
    Case GetBuildfileValue("SetupSSL") <> "YES"
      ' Nothing
    Case Else
      Call ConfigSSL()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AE"
      ' Nothing
    Case GetBuildfileValue("SetupParam") <> "YES"
      ' Nothing
    Case Else
      Call ConfigParam()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AF"
      ' Nothing
    Case GetBuildfileValue("SetupParam") <> "YES"
      ' Nothing
    Case Else
      Call ConfigErrorLog()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AG"
      ' Nothing
    Case GetBuildfileValue("SetupServices") <> "YES"
      ' Nothing
    Case Else
      Call ConfigServiceRecovery()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5AH"
      ' Nothing
    Case GetBuildfileValue("SetupServiceRights") <> "YES"
      ' Nothing
    Case Else
      Call ConfigServiceRights()
  End Select

  Call SetProcessId("5AZ", " Services configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigureComSecurity()
  Call SetProcessId("5AA", "Setup COM Security")
  Dim strCLSId

  strCLSId          = GetBuildfileValue("CLSIdDRUCtlr")
  If strCLSId > "" Then
    Call SetDCOMSecurity("AppID\{" & strCLSId & "}\")
  End If

  strCLSId          = GetBuildfileValue("CLSIdDTExec")
  If strCLSId > "" Then
    Call SetDCOMSecurity("AppID\{" & strCLSId & "}\")
  End If

  strCLSId          = GetBuildfileValue("CLSIdSSIS")
  If strCLSId > "" Then
    Call SetDCOMSecurity("AppID\{" & strCLSId & "}\")
  End If

  strCmd            = "NET LOCALGROUP """ & strGroupDistComUsers & """ """ & strSqlAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  Call SetBuildfileValue("SetupDComStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigureSQLNetworkProtocols()
  Call SetProcessId("5AB", "Configure SQL Server Ports")

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call ConfigSQLDBIP()
      strPath       = strHKLMSQL & strInstRegSQL & "\MSSQLServer\SuperSocketNetLib\AdminConnection\Tcp\"
      Call Util_RegWrite(strPath & "TcpDynamicPorts", strTCPPortDAC, "REG_SZ")
  End Select

  Call SetBuildfileValue("SetupNetworkStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSQLDBIP()
  Call DebugLog("ConfigSQLDBIP: ports for SQL DB Engine")
  Dim intUBound
  Dim arrTCPIP, strTCPIP, strIPAddr, strIPFound

  strPath           = strHKLMSQL & strInstRegSQL & "\MSSQLServer\SuperSocketNetLib\Tcp\"
  objWMIReg.EnumKey strHKLM, Mid(strPath, 6), arrTCPIP
  strIPFound        = "N"
  intUBound         = 0

  For Each strTCPIP In arrTCPIP
    intUBound       = intUBound + 1
    objWMIReg.GetStringValue strHKLM, Mid(strPath, 6) & strTCPIP & "\IpAddress", strIPAddr
    Select Case True
      Case strIPAddr = "127.0.0.1"
        ' Nothing
      Case strIPAddr = "::1"
        ' Nothing
      Case strIPAddr = strClusterIPV4DB
        strIPFound   = "Y"
        Call ConfigNetworkIP(strPath & strTCPIP & "\")
      Case Else
        Call ConfigNetworkIP(strPath & strTCPIP & "\")
    End Select
  Next

  If strIPFound <> "Y" Then
    strTCPIP        = "IP" & Cstr(intUBound)
    strPath         = strHKLMSQL & strInstRegSQL & "\MSSQLServer\SuperSocketNetLib\Tcp\" & strTCPIP & "\"
    Call Util_RegWrite(strPath & "Active",      "1",                   "REG_DWORD")
    Call Util_RegWrite(strPath & "DisplayName", "Specific IP Address", "REG_SZ")
    Call Util_RegWrite(strPath & "Enabled",     "1",                   "REG_DWORD")
    Call Util_RegWrite(strPath & "IPAddress",   strClusterIPV4DB,      "REG_SZ")
    Call ConfigNetworkIP(strPath)
  End If

End Sub


Sub ConfigNetworkIP(strPath)
  Call DebugLog("ConfigNetworkIP: " & strPath)

  Call Util_RegWrite(strPath & "TcpDynamicPorts", "",         "REG_SZ")
  Call Util_RegWrite(strPath & "TcpPort",         strTCPPort, "REG_SZ")

End Sub


Sub ConfigServiceAC()
  Call SetProcessId("5AC", "Setup Service Account Names")
  Dim objSQLConfig

  strPath           = "winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\ROOT\Microsoft\SqlServer\ComputerManagement"
  Select Case True
    Case strSQLVersion >= "SQL2008"
      strPath       = strPath & Left(strSQLVersionNum, 2)
  End Select
  Set objSQLConfig  = GetObject(strPath)

  If strSQLVersion <= "SQL2012" Then
    Call SetServiceAccount(objSQLConfig, "SQLBrowser", "7",  strBrAccount,     strBrPassword,     "")
  End If
'  Call SetServiceAccount(objSQLConfig, strInstFT,      "9",  strFtAccount,     strFtPassword,     "")
  Call SetServiceAccount(objSQLConfig, strInstIS,      "4",  strIsAccount,     strIsPassword,     "Services\SSIS Server\GroupPrefix")
  If strSetupAnalytics = "YES" Then
    Call SetServiceAccount(objSQLConfig, strInstAnal,  "12", strExtSvcAccount, strExtSvcPassword, "")
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetServiceAccount(ByRef objSQLConfig, strServiceName, strServiceType, strServiceAccount, strServicePassword, strGroupPath)
  Call DebugLog("SetServiceAccount: " & strServiceName)

  Dim colService
  Dim objService, objSQLService, objSQLAccount, objSQLResult
  Dim strAccountOld, strServiceGroup

  strCmd            = "SELECT * FROM Win32_Service WHERE Name = '" & strServiceName & "'"
  Set colService    = objWMI.ExecQuery (strCmd)
  For Each objService in colService
    Select Case True
      Case objService.StartName = strServiceAccount
        ' Nothing
      Case Left(objService.StartName, Len(strDomain) + 1) = strDomain & "\" ' Do not change away from Domain account
        ' Nothing
      Case Else
        strAccountOld       = objService.StartName
        Call Util_RunExec("NET STOP " & strServiceName, "", "", -1)
        If intErrSave = -2147023843 Then
          Wscript.Sleep strWaitLong
          Call Util_RunExec("NET STOP " & strServiceName, "", "", 2)
        End If
        Call DebugLog("Change service account for " & strServiceName)
        strCmd              = "SqlService.ServiceName=""" & strServiceName & """,SQLServiceType=" & strServiceType
        strDebugMsg1        = strCmd
        Set objSQLService   = objSQLConfig.Get(strCmd)
        Set objSQLAccount   = objSQLService.Methods_("SetServiceAccount").inParameters.SpawnInstance_()
        objSQLAccount.Properties_.item("ServiceStartName")     = strServiceAccount
        objSQLAccount.Properties_.item("ServiceStartPassword") = strServicePassword
        strDebugMsg2        = "Save new credential: " & strServiceAccount
        Set objSQLResult    = objSQLService.ExecMethod_("SetServiceAccount", objSQLAccount) 
        If objSQLResult.ReturnValue <> 0 Then
          Call SetBuildMessage(strMsgError, "Could not change the service account for " & strServiceName & " to " & strServiceAccount & ", error " & CStr(objSQLResult.ReturnValue))
        End If
      
        Select Case True
          Case strSQLVersion <> "SQL2005"
            ' Nothing
          Case strGroupPath = ""
            ' Nothing
          Case strAccountOld <> "" 
            Call DebugLog("Add new Account in Windows Group")
            strPath         = strHKLMSQL & strGroupPath
            strServiceGroup = objShell.RegRead(strPath)
            strCmd          = "NET LOCALGROUP """ & strServiceGroup & """ """ & strServiceAccount & """ /ADD"
            Call Util_RunExec(strCmd, "", "", 2)
            strCmd          = "NET LOCALGROUP """ & strServiceGroup & """ """ & strAccountOld & """ /DELETE"
            Call Util_RunExec(strCmd, "", "", 2)
        End Select
    End Select
  Next
  Set colService    = Nothing

End Sub


Sub ConfigSSL()
  Call SetProcessId("5AD", "Setup SSL Security")

' 1) Import SSL Certificate
' 2) Configure SSL for SQL AS
' 3) Configure SSL for SQL DB
' 4) Configure SSL for SQL RS

'  Call SetBuildfileValue("SetupSSLStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigParam()
  Call SetProcessId("5AE", "Setup SQL Startup Parameters")
  Dim intTF, strTF

  Call DebugLog("Setup Trace flags")
  intTF             = 2
  For intIdx = 1 To CInt(strNumTF)
    strTF           = GetBuildfileValue("TF" & Right("0" & Cstr(intIdx), 2))
    If strTF <> "" Then
      intTF         = intTF + 1
      If Left(strTF, 1) <> "-" Then
        strTF       = "-" & strTF
      End If
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg" & Cstr(intTF), strTF, "REG_SZ")
    End If
  Next

  Call DebugLog("Configuring SQL Server parameters")
  strCmd            = strHKLMSQL & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg1"
  strPathOld        = objShell.RegRead(strCmd)
  If Right(strPathOld, 4) <> ".OUT" Then
    Call Util_RegWrite(strCmd, strPathOld & ".OUT", "REG_SZ")                                         ' SQL Server Error Log location
  End If

  Call SetBuildfileValue("SetupParamStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigErrorLog()
  Call SetProcessId("5AF", "Configure ErrorLog Files")

  Call StopSQL()

  strPathOld        = Mid(Left(strPathOld, InstrRev(strPathOld, "\")), 3)
  If objFSO.FileExists(strPathOld & "ERRORLOG") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG"" ERRORLOG.OUT",  "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.1") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.1"" ERRORLOG1.OUT", "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.2") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.2"" ERRORLOG2.OUT", "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.3") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.3"" ERRORLOG3.OUT", "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.4") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.4"" ERRORLOG4.OUT", "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.5") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.5"" ERRORLOG5.OUT", "", strResponseYes, 2)
  End If
  If objFSO.FileExists(strPathOld & "ERRORLOG.6") Then
    Call Util_RunExec("%COMSPEC% /D /C RENAME """ & strPathOld & "ERRORLOG.6"" ERRORLOG6.OUT", "", strResponseYes, 2)
  End If


  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\NumErrorLogs", strNumErrorLogs, "REG_DWORD")

  Call StartSQL()
  If strSetupSQLDBAG = "YES" Then
    Call StartSQLAgent()
  End If

  Call SetBuildfileValue("SetupParamStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigServiceRecovery()
  Call SetProcessId("5AG", "Setup SQL Service Recovery")

  Select Case True
    Case strSetupSQLDB <> "YES" 
      ' Nothing
    Case strInstADHelper = "" 
      ' Nothing
    Case Else
      strCmd        = "SC CONFIG " & strInstADHelper & " START= DEMAND" 
      Call Util_RunExec(strCmd, "", "", 2) 
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES" 
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case Else
      strCmd        = "SC CONFIG SQLBrowser START= AUTO"
      Call Util_RunExec(strCmd, "", "", 2)  
      strCmd        = "NET START ""SQLBrowser"""
      Call Util_RunExec(strCmd, "", "", 2)  
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES" 
      ' Nothing
    Case strSetupSQLDBCluster = "YES" 
      ' Nothing
    Case Else
      strCmd        = "SC FAILURE " & strInstSQL   & " RESET= 88400 ACTIONS= RESTART/180000/RESTART/180000/""""/0" ' Restart after 3 minute delay up to 2 times in 1 day
      Call Util_RunExec(strCmd, "", "", 2)
  End Select

  Select Case True
    Case strSetupSQLDBAG <> "YES" 
      ' Nothing
    Case strSetupSQLDBCluster = "YES" 
      ' Nothing
    Case Else
      strCmd        = "SC FAILURE " & strInstAgent & " RESET= 88400 ACTIONS= RESTART/180000/RESTART/180000/""""/0"
      Call Util_RunExec(strCmd, "", "", 2) 
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES" 
      ' Nothing
    Case strSetupSQLASCluster = "YES" 
      ' Nothing
    Case Else
      strCmd        = "SC FAILURE " & strInstAS & " RESET= 88400 ACTIONS= RESTART/180000/RESTART/180000/""""/0"
      Call Util_RunExec(strCmd, "", "", 2) 
  End Select

  Select Case True
    Case strSetupSQLIS <> "YES" 
      ' Nothing
    Case strSetupSSISCluster = "YES" 
      ' Nothing
    Case Else
      strCmd        = "SC FAILURE " & strInstIS & " RESET= 88400 ACTIONS= RESTART/180000/RESTART/180000/""""/0"
      Call Util_RunExec(strCmd, "", "", 2)
      strCmd        = "SC CONFIG " & strInstIS & " START= " & GetStartupMode(strInstIS, strIsSvcStartuptype)
      Call Util_RunExec(strCmd, "", "", 2) 
  End Select

  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case (StrComp(Left(strRSInstallMode, 9), "FilesOnly", vbTextCompare) = 0) And (strSetupSQLRSCluster <> "YES")
      ' Nothing
    Case Else
      strCmd        = "SC FAILURE " & strInstRS & " RESET= 88400 ACTIONS= RESTART/180000/RESTART/180000/""""/0"
      Call Util_RunExec(strCmd, "", "", 2) 
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES" 
      ' Nothing
    Case Else
      strCmd        = "SC CONFIG SQLWriter START= " & GetStartupMode("SQLWriter", strWriterSvcStartupType)
      Call Util_RunExec(strCmd, "", "", 2)
  End Select

  Call DebugLog("Start required Services")
  Select Case True
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strSetupSQLDBAG <> "YES"
      ' Nothing
    Case Else
      Call StartSQLAgent()
      If strSetupSQLRS = "YES" Then
        Call StartSSRS("")
      End If
  End Select

  Call SetBuildfileValue("SetupServicesStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Function GetStartupMode(strService, strMode)
  Call DebugLog("GetStartupMode: " & strService)

  Select Case True
    Case strMode = "0"
      GetStartupMode = "DEMAND"
    Case strMode = "MANUAL"
      GetStartupMode = "DEMAND"
    Case strMode = "1"
      GetStartupMode = "AUTO"
    Case strMode = "AUTOMATIC"
      GetStartupMode = "AUTO"
    Case strMode = "2"
      Call SetBuildMessage(strMsgInfo, "Startup mode for " & strService & " changed to DEMAND")
      GetStartupMode = "DEMAND"
    Case strMode = "DISABLED"
      Call SetBuildMessage(strMsgInfo, "Startup mode for " & strService & " changed to DEMAND")
      GetStartupMode = "DEMAND"
    Case Else
      GetStartupMode = strMode
  End Select

End Function


Sub ConfigServiceRights()
  Call SetProcessId("5AH", "Setup Service Rights")

'  Call Util_RunExec("SC SDSHOW Eventlog         > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW RpcSs            > %TEMP%\SCData.txt", 0)
'  Call Util_RunExec("SC SDSHOW clr_optimization_v2.0.50727_32 > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW clr_optimization_v2.0.50727_64 > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW MSDTC           > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW MSSQLFDLauncher > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW MSSQLServer     > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW SQLSERVERAGENT  > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW MSSQLServerOLAPService > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW SQLBrowser      > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW MsDtsServer100  > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW ReportServer    > %TEMP%\SCData.txt", "", "", 0)
'  Call Util_RunExec("SC SDSHOW SQLWriter       > %TEMP%\SCData.txt", "", "", 0)

  Call SetBuildfileValue("SetupServiceRightsStatus", strStatusManual) ' Temporary until this process is automated
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigInstance()
  Call SetProcessId("5B", "Instance configuration")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BA"
      ' Nothing
    Case GetBuildfileValue("SetupSQLServer") <> "YES"
      ' Nothing
    Case Else
      Call ConfigSQLServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BB"
      ' Nothing
    Case strSetupDBMail <> "YES"
      ' Nothing
    Case Else
      Call ConfigDBMail()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BC"
      ' Nothing
    Case strSetupSQLMail <> "YES"
      ' Nothing
    Case Else
      Call ConfigSQLMail()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BD"
      ' Nothing
    Case GetBuildfileValue("SetupSQLInst") <> "YES"
      ' Nothing
    Case Else
      Call ConfigSQLInstance()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BE"
      ' Nothing
    Case GetBuildfileValue("SetupSQLAgent") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSQLAgentStatus", strStatusPreConfig)
    Case Else
      Call ConfigSQLAgent()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BFZ"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      Call ConfigOLAPInstance()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BGZ"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case Else
      Call ConfigSSIS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BHZ"
      ' Nothing
    Case strSetupSQLNS <> "YES"
      ' Nothing
    Case Else
      Call SetupSSNS()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BIZ"
      ' Nothing
    Case GetBuildfileValue("SetupPolyBase") <> "YES"
      ' Nothing
    Case Else
      Call SetupPolyBase()
  End Select

  Call SetProcessId("5BZ", " Instance configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigSQLServer()
  Call SetProcessId("5BA", "SQL Surface Area options")

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'show advanced options',           '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)
  Wscript.Sleep strWaitShort
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'clr enabled',                     '" & strSetCLREnabled & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'cost threshold for parallelism',  '" & strSetCostThreshold & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'disallow results from triggers',  '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'Ole Automation Procedures',       '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'remote access',                   '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'remote admin connections',        '" & strSetRemoteAdminConnections & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'remote proc trans',               '" & strSetRemoteProcTrans & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'SMO and DMO XPs',                 '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'xp_cmdshell',                     '" & strSetxpCmdshell & "';""", 0)
  If strSetupSQLDBAG = "YES" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'Agent XPs',                     '1';""", 0)
  End If

  Select Case True
    Case strEdition <> "STANDARD"
      ' Nothing
    Case strSQLVersion >= "SQL2008R2"
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'backup compression default',    '1';""", 0)
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2014"
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'backup checksum default',       '1';""", 0)
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strEdition = "BUSINESS INTELLIGENCE"
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'backup compression default',    '1';""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'optimize for ad hoc workloads', '" & strSetOptimizeForAdHocWorkloads & "';""", 0)
    Case strEditionEnt = "YES"
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'backup compression default',    '1';""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'optimize for ad hoc workloads', '" & strSetOptimizeForAdHocWorkloads & "';""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'EKM provider enabled',          '1';""", 0)
      If strSQLVersion >= "SQL2016" Then
        If strSetupStretch = "YES" Then
          Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'remote data archive',       '1';""", 0)
        End If
      End If
  End Select

  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  If strSQLVersion = "SQL2005" Then
    Call Debuglog("Set default backup location")
    Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\BackupDirectory", strDirBackup, "REG_SZ")
  End If

  Select Case True
    Case (strSQLVersion = "SQL2005") And (strEdition = "EXPRESS")
      ' Nothing
    Case Else
      Call FBLog(" Backing Up Master Keys")
      Call DebugLog("Database Master Key")
      Call Util_ExecSQL(strCmdSQL & "-Q", """IF NOT EXISTS (SELECT * FROM sys.symmetric_keys WHERE symmetric_key_id = 101) CREATE MASTER KEY ENCRYPTION BY PASSWORD='" & strsaPwd & "';""", 0)
      strPathNew        = strDirSystemDataBackup & "master" & "DBMasterKey.snk"
      If objFSO.FileExists(strPathNew) Then
        Call objFSO.DeleteFile(strPathNew, True)
        Wscript.Sleep strWaitShort
      End If
      Call Util_ExecSQL(strCmdSQL & "-Q", """BACKUP MASTER KEY TO FILE='" & strPathNew & "' ENCRYPTION BY PASSWORD='" & strsaPwd & "';""", 0)
      Call DebugLog("Service Master Key")
      strPathNew    = strDirSystemDataBackup & "ServiceMasterKey.snk"
      If objFSO.FileExists(strPathNew) Then
        Call objFSO.DeleteFile(strPathNew, True)
        Wscript.Sleep strWaitShort
      End If
      Call Util_ExecSQL(strCmdSQL & "-Q", """BACKUP SERVICE MASTER KEY TO FILE='" & strPathNew & "' ENCRYPTION BY PASSWORD='" & strsaPwd & "';""", 0)
  End Select


  Call SetBuildfileValue("SetupSQLServerStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigDBMail()
  Call SetProcessId("5BB", "Setup DB Mail")

  intIdx            = InStr(strSqlAccount, "\")
  Select Case True
    Case strUserDNSDomain = ""
      ' Nothing - User is not in a domain
    Case intIdx = 0
      ' Nothing - SQL is not running using a domain account
    Case Left(strSqlAccount, intIdx) = Left(strNTAuthAccount, InStr(strNTAuthAccount, "\"))
      ' Nothing - SQL is not running using a domain account
    Case strMailServer = ""
      Call SetBuildMessage(strMsgWarning, "Unable to find Mail Server, DB Mail profile not created")
    Case Else
      Call SetupDBMailProfile(strSQLEmail)
  End Select

  Select Case True
    Case strDBMailOK <> "Y" 
      Call SetBuildfileValue("SetupDBMailStatus", strStatusBypassed & ", no Mail Server")
    Case Else
      Call SetBuildfileValue("SetupDBMailStatus", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDBMailProfile(strMailAcnt)
  Call DebugLog("SetupDBMailProfile: " & strMailAcnt)
  Dim objInstParm

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'Database Mail XPs', '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  Call SetXMLParm(objInstParm, "PathMain",       strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",       "-v strMailAcnt='" & strMailAcnt & "' strMailServer='" & strMailServer & "'")
  Call SetXMLParm(objInstParm, "LogXtra",        "Set-DBMail")
  Call RunInstall("DBMail",    "Set-DBMail.sql", objInstParm)

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC msdb.dbo.sysmail_update_profile_sp @profile_name = '" & strDBMailProfile & "', @profile_id = 1;""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC msdb.dbo.sysmail_add_principalprofile_sp @profile_name = '" & strDBMailProfile & "', @principal_name = 'public',  @is_default = 1;""", 0)

  strDBMailOK       = "Y"

End Sub


Sub ConfigSQLMail()
  Call SetProcessId("5BC", "Setup SQL Mail")

  Select Case True
    Case strMailServer = "" 
      Call SetBuildMessage(strMsgWarning, " Unable to find Mail Server, SQL Mail profile not created")
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'SQL Mail XPs', '1';""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)
  '   Set registry entries recommended in blogs.msdn.com/stephen_griffin/archive/2007/12/28/sqlmail-hates-mapi.aspx
      strCmd        = strHKLMSQL & strInstRegSQL & "\MSSQLServer\MAPI_NO_MAIL" ' See KB 329375
      Call Util_RegWrite(strCmd, 1, "REG_DWORD")
      strCmd        = strHKLMSQL & strInstRegSQL & "\MSSQLServer\MAPIInitializeFlags" ' See KB 839405
      Call Util_RegWrite(strCmd, 65537, "REG_DWORD") ' Hex 10001
      If strInstance = "MSSQLSERVER" Then
        strCmd      = "HKLM\SOFTWARE\Microsoft\MSSQLSERVER\MSSQLSERVER\MAPI_NO_MAIL"
        Call Util_RegWrite(strCmd, 1, "REG_DWORD")
        strCmd      = "HKLM\SOFTWARE\Microsoft\MSSQLSERVER\MSSQLSERVER\MAPIInitializeFlags"
        Call Util_RegWrite(strCmd, 65537, "REG_DWORD") ' Hex 10001
      End If
      strSQLMailOK  = "Y"
  End Select

  Select Case True
    Case strSQLMailOK <> "Y" 
      Call SetBuildfileValue("SetupSQLMailStatus", strStatusBypassed & ", no Mail Server")
    Case Else
      Call SetBuildfileValue("SetupSQLMailStatus", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSQLInstance()
  Call SetProcessId("5BD", "SQL Instance properties")

  Select Case True
    Case Instr("EXPRESS WORKGROUP", strEdition) > 0
      ' Nothing
    Case strSQLVersion >= "SQL2012"
      ' Nothing
    Case Else
      Call DebugLog("Set AWE")
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'awe enabled', '1';""", 0)
  End Select

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'max server memory (MB)', '" & strSQLMaxMemory & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'min server memory (MB)', '" & strSQLMinMemory & "';""", 0)

  Call DebugLog("Set SQL Audit level")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\AuditLevel", strAuditLevel, "REG_DWORD")

  Call DebugLog("Max degree of parallelism")
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'Max Degree of Parallelism', '" & Cstr(intMaxDop) & "';""", 0)

  If strSQLVersion = "SQL2005" Then
    Call DebugLog("Default DB Data File location")
    Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\DefaultData", strDirData, "REG_SZ")
    Call DebugLog("Default DB Log File location")
    Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\MSSQLServer\DefaultLog", strDirLog, "REG_SZ")
  End If

  Call DebugLog("Default Alerts Operator")
  strCmd            = """EXEC msdb.dbo.sp_add_operator @name = N'" & strSQLOperator & "', @enabled = 1, @email_address = N'" & strSqlEmail & "', @category_name = N'[Uncategorized]', @weekday_pager_start_time = 80000, @weekday_pager_end_time = 180000, @saturday_pager_start_time = 80000, @saturday_pager_end_time = 180000, @sunday_pager_start_time = 80000, @sunday_pager_end_time = 180000, @pager_days = 62;"""
  Call Util_ExecSQL(strCmdSQL & "-Q", strCmd, 1)

  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  Call SetBuildfileValue("SetupSQLInstStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSQLAgent()
  Call SetProcessId("5BE", "SQL Agent options")

  Call DebugLog("Allow Token substitution in Batch Jobs")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\AlertReplaceRuntimeTokens", 1, "REG_DWORD")

  Call DebugLog("Mail Profile")
  Select Case True
    Case strDBMailOK = "Y"
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\DatabaseMailProfile",   strDBMailProfile, "REG_SZ")
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\UseDatabaseMail",       1, "REG_DWORD")
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\EmailSaveSent",         0, "REG_DWORD")
    Case strSQLMailOK = "Y"
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\EmailProfile",          "Microsoft Exchange Settings", "REG_SZ")
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\UseDatabaseMail",       0, "REG_DWORD")
      Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\EmailSaveSent",         0, "REG_DWORD")
    Case strSetupDBMail = "YES" Or strSetupSQLMail = "YES"
      Call DebugLog(" SQL Agent mail setup bypassed")
  End Select

  Call DebugLog("SQL Alert Emails")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\AlertFailSafeOperator",     strSQLOperator, "REG_SZ")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\AlertNotificationMethod",   1,              "REG_DWORD")

  Call DebugLog("Number of job history rows")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\JobHistoryMaxRows",         strAgentMaxHistory , "REG_DWORD")

  Call DebugLog("Maximum history rows per job")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\JobHistoryMaxRowsPerJob",   strAgentJobHistory, "REG_DWORD")

  Call DebugLog("Do not automatically restart Agent if it fails")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\MonitorAutoStart",          0, "REG_DWORD")

  Call DebugLog("Do not automatically restart SQL Server if it fails")
  Call Util_RegWrite(strHKLMSQL & strInstRegSQL & "\SQLServerAgent\RestartSQLServer",          0, "REG_DWORD")

  Call SetBuildfileValue("SetupSQLAgentStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigOLAPInstance()
  Call SetProcessId("5BF", "AS Options")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BFA"
      ' Nothing
    Case GetBuildfileValue("SetupOLAPAPI") <> "YES"
      ' Nothing
    Case Else
      Call ConfigOLAPAPI()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BFB"
      ' Nothing
    Case GetBuildfileValue("SetupOLAP") <> "YES"
      ' Nothing
    Case Else
      Call ConfigSSASInstance()
  End Select

  Call StartSSAS()
  Call SetProcessId("5BFZ", " AS Options" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigOLAPAPI()
  Call SetProcessId("5BFA", "Setup SSAS Management API")
  Dim strASDLL

  strPathOld        = GetBuildfileValue("DirASDLL") & "\"
  strASDLL          = "Microsoft.AnalysisServices.dll"
  Select Case True
    Case strSQLVersion >= "SQL2016"
      strPathNew    = strDirProg & "\" & strSQLVersionNum & "\SDK\Assemblies" & "\"
      Call CopyASDLL(strPathOld, strPathNew, strASDLL)
      strASDLL      = "Microsoft.AnalysisServices.Core.dll"
      Call CopyASDLL(strPathOld, strPathNew, strASDLL)
    Case strSQLVersion >= "SQL2012"
      strPathNew    = strDirProg & "\" & strSQLVersionNum & "\SDK\Assemblies" & "\"
      Call CopyASDLL(strPathOld, strPathNew, strASDLL)
    Case strSQLVersion >= "SQL2008"
      strPathNew    = strDirProgX86 & "\" & strSQLVersionNum & "\SDK\Assemblies" & "\"
      Call CopyASDLL(strPathOld, strPathNew, strASDLL)
  End Select

  If Not objFSO.FileExists(strPathNew & strASDLL) Then
    Call SetBuildfileValue("SetupOLAPAPIStatus", strStatusBypassed)
    Call DebugLog(" " & strProcessIdDesc & strStatusBypassed)
    Exit Sub
  End If

  Call SetBuildfileValue("SetupOLAPAPIStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CopyASDLL(strPathOld, strPathNew, strASDLL)
  Call DebugLog("CopyASDLL: " & strPathOld & ", " & strPathNew & ", " & strASDLL)

  strDebugMsg1      = "Source: " & strPathOld & strASDLL
  Select Case True
    Case strPathOld = ""
      ' Nothing
    Case Not objFSO.FileExists(strPathOld & strASDLL)
      ' Nothing
    Case objFSO.FileExists(strPathNew & strASDLL)
      ' Nothing
    Case Else
      strDebugMsg2  = "Target: " & strPathNew & strASDLL
      objFSO.CopyFile strPathOld & strASDLL, strPathNew & strASDLL
      Wscript.Sleep strWaitShort 
  End Select

  Select Case True
    Case Not objFSO.FileExists(strPathNew & strASDLL)
      ' Nothing
    Case Else
      strCmd        = GetBuildfileValue("RegasmExe") & """" & strPathNew & strASDLL & """"
      Call Util_RunExec(strCmd, "", "", 1)
  End Select

End Sub


Sub ConfigSSASInstance()
  Call SetProcessId("5BFB", "Setup SSAS Instance settings")
  Dim objASConfig, objASNode, objFolder, objInstParm
  Dim colASConfig
  Dim strASConfig, strDirOlap, strDirBackupAS, strDirDataAS, strDirDataASOld, strDirLogAS, strDirLogASOld, strDirTempAS, strFileData, strInstRegAS, strVolDataAS

  strInstRegAS      = GetBuildfileValue("InstRegAS")
  strDirBackupAS    = GetBuildfileValue("DirBackupAS")
  strDirDataAS      = GetBuildfileValue("DirDataAS")
  strDirLogAS       = GetBuildfileValue("DirLogAS")
  strDirTempAS      = GetBuildfileValue("DirTempAS")
  strVolDataAS      = GetBuildfileValue("VolDataAS")
  Select Case True
    Case strSQLVersion = "SQL2005"
      strDirOlap    = objShell.RegRead(strHKLMSQL & strInstRegAS & "\Setup\DataDir") & strInstRegAS & "\OLAP"
    Case Else
      strDirOlap    = strDirDataAS
  End Select

  Call DebugLog("Adding Analysis Services administration accounts")
  Call StartSSAS()  ' Starting for the first time causes the Configuration file to be populated
  Select Case True
    Case strSQLVersion = "SQL2005"
      Call DebugLog("Adding Analysis Services administration accounts")
      Call SetXMLParm(objInstParm, "PathMain",     strPathFBScripts)
      Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
      Call RunInstall("SetupOLAP", "Set-ASAccounts.vbs", objInstParm)
    Case Else
      ' Nothing (Accounts added as part of Install)
  End Select
  Call StopSSAS()

  Call DebugLog("Set msmdsrv.ini Options in " & strDirOLAP & "\Config")
  strASConfig       = strDirOLAP & "\Config\msmdsrv.ini" 
  objFSO.CopyFile strASConfig, strASConfig & ".original"
  Set objASConfig   = CreateObject ("Microsoft.XMLDOM") 
  objASConfig.async = "false"
  objASConfig.load(strASConfig & ".original")

  Select Case True
    Case strType = "UPGRADE" 
      ' Nothing
    Case strSQLVersion = "SQL2005" 
      Set objASNode   = objASConfig.documentElement.selectSingleNode("DataDir")
      strDirDataASOld = objASNode.Text
      If strDirDataASOld <> strDirDataAS Then
        Call DebugLog("Moving AS data files")
        Set objFolder = objFSO.GetFolder(strDirDataASOld)
        objFolder.Copy strDirDataAS
        objFolder.Delete(True)
      End If
      Call SetXMLConfigValue(objASConfig, "", "DataDir",                                      strDirDataAS, "")
      Set objASNode   = objASConfig.documentElement.selectSingleNode("LogDir")
      strDirLogASOld  = objASNode.Text
      If strDirLogASOld <> strDirLogAS Then
        Call DebugLog("Moving AS log files")
        Set objFolder = objFSO.GetFolder(strDirLogASOld)
        objFolder.Copy strDirLogAS
        objFolder.Delete(True)
      End If
      Call SetXMLConfigValue(objASConfig, "", "LogDir",                                       strDirLogAS, "")
      Call SetXMLConfigValue(objASConfig, "", "BackupDir",                                    strDirBackupAS, "")
      Call SetXMLConfigValue(objASConfig, "", "TempDir",                                      strDirTempAS, "")
      Call SetXMLConfigValue(objASConfig, "", "AllowedBrowsingFolders",                       strDirBackupAS & "|" & strDirDataAS, "")
  End Select

  If strType <> "WORKSTATION" Then
    Call SetXMLConfigValue(objASConfig, "Security", "BuiltinAdminsAreServerAdmins", "0", "")
  End If

  Call SetXMLConfigValue(objASConfig,   "Memory", "LowMemoryLimit",                           GetBuildfileValue("SetLowMemLimit"), "")
  Call SetXMLConfigValue(objASConfig,   "Memory", "HardMemoryLimit",                          GetBuildfileValue("SetHardMemLimit"), "")
  Call SetXMLConfigValue(objASConfig,   "Memory", "TotalMemoryLimit",                         GetBuildfileValue("SetTotalMemLimit"), "")

  If strSQLVersion >= "SQL2012" Then
    Call DebugLog("Setting ROLAP Optimisations") ' Google each item to find reasons why they are set
    Call SetXMLConfigValue(objASConfig, "Memory",           "VertiPaqMemoryLimit",            GetBuildfileValue("SetVertiMemLimit"), "")
    Call SetXMLConfigValue(objASConfig, "OLAP/ProcessPlan", "EnableRolapDistinctCountOnDataSource", "1", "")
    Call SetXMLConfigValue(objASConfig, "OLAP/Process",     "ROLAPDimensionProcessingEffort", GetBuildfileValue("SetROLAPDimensionProcessingEffort"), "")
    Call SetXMLConfigValue(objASConfig, "OLAP/Process",     "CheckDistinctRecordSortOrder",   "0", "")
    Call SetXMLConfigValue(objASConfig, "OLAP/Query",       "SkipROLAPDatasourceMatching",    "1", "")
  End If

  objASConfig.save strASConfig

  Call DebugLog("Set AS Port value")
  Select Case True
    Case strSQLVersion = "SQL2005" 
      strASConfig   = strDirProgSysX86 & "\Microsoft SQL Server\" & strSQLVersionNum & "\Shared\ASConfig\msmredir.ini"
      If objFSO.FileExists(strASConfig) Then
        objASConfig.load(strASConfig)
        Call SetXMLConfigValue(objASConfig, "Instances/" & strInstAS, "Port",                 strTCPPortAS, "") 
      End If
    Case Else  
      Call SetXMLConfigValue(objASConfig, "", "Port",                                         strTCPPortAS, "")
  End Select

  objASConfig.save strASConfig
  Set objASConfig   = Nothing  

  Call SetBuildfileValue("SetupOLAPStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSSIS()
  Call SetProcessId("5BG", "Setup SSIS Service")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BGA"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case strSetupSSISCluster = "YES"
      ' Nothing
    Case Else
      Call ConfigSSISOptions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BGB"
      ' Nothing
    Case strSetupSSISDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call ConfigSSISDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BGC"
      ' Nothing
    Case strSetupSSISDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call OptimiseSSISDB()
  End Select

  Call SetProcessId("5BGZ", " Setup SSIS Service" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigSSISOptions()
  Call SetProcessId("5BGA", "SSIS Options")
  Dim colSSISConfig
  Dim objSSISConfig, objSSISFolder, objSSISFolderNew, objSSISFolderCol
  Dim strSetupFound, strSSISConfig, strSSISFolder, strSSISNode

  Call DebugLog("Saving original SSIS configuration file")
  Call Util_RunExec("NET STOP " & strInstIS, "", "", 2)
  strSSISConfig     = objShell.RegRead(strRegSSIS)
  If strSSISConfig = "" Then
    strSSISConfig   = objShell.RegRead(strRegSSISSetup) & "Binn\MsDtsSrvr.ini.xml"
    Call Util_RegWrite(strRegSSIS, strSSISConfig, "REG_SZ")
  End If
  objFSO.CopyFile strSSISConfig, strSSISConfig & ".original" 

  Call DebugLog("Setting SSIS configuration values")
  strSetupFound    = "N"
  Set objSSISFolderNew = Nothing
  Set objSSISConfig = CreateObject ("Microsoft.XMLDOM") 
  objSSISConfig.async  = "false"
  objSSISConfig.load(strSSISConfig)
  Set objSSISFolderCol = objSSISConfig.documentElement.selectSingleNode("TopLevelFolders")
  Set colSSISConfig    = objSSISConfig.getElementsByTagName("Folder")
  For Each objSSISFolder In colSSISConfig
    strSSISFolder      = objSSISFolder.getAttribute("xsi:type")
    If strSSISFolder = "SqlServerFolder" Then
      Set strSSISNode  = objSSISFolder.selectSingleNode("./ServerName")
      Select Case True
        Case strSSISNode.Text = strServInst
          strSetupFound   = "Y"
        Case Left(strSSISNode.Text, 1) = "." 
          strSSISNode.Text = strServInst
          strSetupFound   = "Y"
        Case Else
          Set objSSISFolderNew = objSSISFolder.cloneNode(True)
      End Select
    End If
  Next

  Select Case True
    Case strSetupFound = "Y"
      ' Nothing
    Case objSSISFolderNew Is Nothing
      ' Nothing
    Case Else
      Set strSSISNode  = objSSISFolderNew.selectSingleNode("./ServerName")
      strSSISNode.Text = strServInst
      objSSISFolderCol.appendChild(objSSISFolderNew)
  End Select

  objSSISConfig.save(strSSISConfig)
  Set colSSISConfig = Nothing
  Set objSSISConfig = Nothing

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSSISDB()
  Call SetProcessId("5BGB", "Configure SSIS Catalog DB")
  Dim strPathPS

  strPathNew        = strPathTemp & "\Set-SSISDB"
  Call SetupFolder(strPathNew)

  Set objFile       = objFSO.GetFile(strPathFBScripts & "Set-SSISDB.ps1")
  strPathPS         = strPathNew & "\Set-SSISDB.ps1"
  objFile.Copy strPathPS, True
  Wscript.Sleep strWaitShort
  strCmd            = strCmdPS & " -ExecutionPolicy Bypass -File """ & strPathPS & """ -HostServer """ & strServInst & """ -dbname """ & strSSISDB & """ -password """ & strSSISPassword & """"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  Set objFolder = objFSO.GetFolder(strPathNew)
  objFolder.Delete(True)

  Call BackupDBMasterKey(strSSISDB, strSSISPassword)

  Call SetBuildfileValue("SetupSSISDBStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub OptimiseSSISDB()
  Call SetProcessId("5BGC", "Optimise SSIS Catalog DB")

  strCmd            = """EXEC sp_procoption @ProcName = 'sp_ssis_startup', @OptionName = 'startup', @OptionValue = 1;"""
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", strCmd, 0)

' Needed if SSISDB is moved - See https://docs.microsoft.com/en-us/sql/integration-services/backup-restore-and-move-the-ssis-catalog
  strCmd            = """CREATE ASYMMETRIC KEY MS_SQLEnableSystemAssemblyLoadingKey FROM Executable File = '" & GetBuildfileValue("PathSSIS") & "';"""
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", strCmd, -1)
  strCmd            = """CREATE LOGIN ##MS_SQLEnableSystemAssemblyLoadingUser## FROM ASYMMETRIC KEY MS_SQLEnableSystemAssemblyLoadingKey;"""
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", strCmd, -1)
  strCmd            = """GRANT UNSAFE ASSEMBLY TO ##MS_SQLEnableSystemAssemblyLoadingUser##;"""
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", strCmd, 0)

  Select Case True
    Case strActionAO = "ADDNODE"
      Call SetupSSISSecondary()
    Case Else
      Call SetupSSISPrimary()
  End Select

  If strSetupAlwaysOn = "YES" Then
    Call SetupSSISMaintenance()
  End If

  Call SetBuildfileValue("SetupSSISDBStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSSISPrimary
  Call DebugLog("SetupSSISPrimary:")
  Dim objInstParm

  Call SetDBOptions(strSSISDB, "", "")

  Call SetXMLParm(objInstParm, "PathMain",              strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",              "-d """ & strSSISDB & """")
  Call RunInstall("SSISDB",    "Set-SSISDBOptions.sql", objInstParm)
  
  Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", """EXEC catalog.configure_catalog 'OPERATION_CLEANUP_ENABLED', 'TRUE';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", """EXEC catalog.configure_catalog 'RETENTION_WINDOW',          '" & GetBuildfileValue("SSISRetention") & "';""", 0)

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case Else ' See https://www.sqlservermigrations.com/2018/08/adding-an-integration-services-catalog-to-always-on-availability-groups-subquery-error/
      strCmd        = """ALTER ROLE ssis_failover_monitoring_agent ADD MEMBER ##MS_SSISServerCleanupJobUser##;"""
      Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", strCmd, 0)
      strCmd        = """GRANT EXECUTE ON internal.refresh_replica_status TO ssis_failover_monitoring_agent;"""
      Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", strCmd, 0)
      strCmd        = """GRANT EXECUTE ON internal.update_replica_info TO ssis_failover_monitoring_agent;"""
      Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", strCmd, 0)
  End Select
  
  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else  
      Call Util_ExecSQL(strCmdSQL & "-d """ & strSSISDB & """ -Q", "EXEC [internal].[add_replica_info] @server_name = '" & strServInst & "', @state = 1", -1)
  End Select

End Sub


Sub SetupSSISSecondary
  Call DebugLog("SetupSSISSecondary:")

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else  
      Call Util_ExecSQL("""" & GetBuildfileValue("PathCmdSQL") & """ -S """ & strGroupAO & """ -E -b -e " & "-Q", "EXEC [internal].[add_replica_info] @server_name = '" & strServInst & "', @state = 2", -1)
  End Select

End Sub


Sub SetupSSISMaintenance
  Call DebugLog("SetupSSISMaintenance:")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",              strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",              "-d """ & strSSISDB & """")
  Call RunInstall("SSISDB",    "Set-SSISDBMaintenance.sql", objInstParm)

End Sub


Sub SetupSSNS()
  Call SetProcessId("5BH", "Configure SSNS Instance")

'  No actions needed

  Call SetBuildfileValue("SetupSSNSStatus", strStatusComplete)
  Call SetProcessId("5BHZ", " Setup SSNS Instance" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupPolyBase()
  Call SetProcessId("5BI", "Configure PolyBase Instance")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BIA"
      ' Nothing
    Case Else
      Call ConfigPolyBaseWindows()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5BIB"
      ' Nothing
    Case strSQLVersion < "SQL2019"
      ' Nothing
    Case Else
      Call ConfigPolyBaseSQL()
  End Select

  Call SetProcessId("5BIZ", " Setup PolyBase Instance" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigPolyBaseWindows()
  Call SetProcessId("5BIA", "Configure PolyBase for Windows")
  Dim strRuleName

  strRuleName       = "SQL Server PolyBase - Database Engine - " & strInstance & " (TCP-In)"
  strCmd            = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""" & strRuleName & """ "
  strCmd            = strCmd & "NEW PROFILE=ALL ENABLE=YES"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  strRuleName       = "SQL Server PolyBase - PolyBase Services - " & strInstance & " (TCP-In)"
  strCmd            = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""" & strRuleName & """ "
  strCmd            = strCmd & "NEW PROFILE=ALL ENABLE=YES"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  strRuleName       = "SQL Server PolyBase - SQL Browser - (UDP-In)"
  strCmd            = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""" & strRuleName & """ "
  strCmd            = strCmd & "NEW PROFILE=ALL ENABLE=YES"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigPolyBaseSQL()
  Call SetProcessId("5BIB", "Configure PolyBase for SQL")

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'polybase enabled',                '1';""", 0)

  Call ProcessEnd(strStatusComplete)

End Sub

 
Sub ConfigAccounts()
  Call SetProcessId("5C", "Account configuration")
  Dim strSetupStdAccounts

  strSetupStdAccounts = GetBuildfileValue("SetupStdAccounts")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CA"
      ' Nothing
    Case strSetupStdAccounts <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupStdAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigStdAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CB"
      ' Nothing
    Case GetBuildfileValue("SetupSAAccounts") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSAAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigSAAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CC"
      ' Nothing
    Case GetBuildfileValue("SetupNonSAAccounts") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupNonSAAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigNonSAAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CD"
      ' Nothing
    Case GetBuildfileValue("SetupDisableSA") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupDisableSAStatus", strStatusPreConfig)
    Case Else
      Call ConfigDisableSA()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CE"
      ' Nothing
    Case GetBuildfileValue("SetupCmdshell") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupCmdshellStatus", strStatusPreConfig)
    Case Else
      Call ConfigCmdshell()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CF"
      ' Nothing
    Case strSetupStdAccounts <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupStdAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigDBOwnerAccount()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5CG"
      ' Nothing
    Case strSetupStdAccounts <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupStdAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigUserAccounts()
  End Select

  Call SetProcessId("5CZ", " Account configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigStdAccounts()
  Call SetProcessId("5CA", "Setup Standard Accounts")
  Dim strLogin

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case Else
      If strUserDNSDomain <> "" Then
        strLogin    = strDomain & "\" & strUserName
      Else
        strLogin    = strServer & "\" & strUserName
      End If
      strCmd        = "CREATE LOGIN [" & strLogin & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-r -Q", """" & strCmd & """", 1)
  End Select     

  strLogin          = ""
  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strAsAccount = strSqlAccount
      ' Nothing
    Case Ucase(strAsAccount) = Ucase(strNTAuthOSName )
      strLogin      = strNTAuthAccount
    Case Else
      strLogin      = strAsAccount
  End Select
  If strLogin > "" Then
    strCmd          = "CREATE LOGIN [" & strLogin & "] FROM WINDOWS;"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
    strCmd          = "GRANT ALTER TRACE TO [" & strLogin & "];"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
  End If

  Call DebugLog("Cluster Service Account")
  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strOSVersion >= "6.0"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case Else
      strLogin      = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\ClusSvc\ObjectName")
      strDebugMsg1  = strLogin
      Select Case True
        Case Left(strLogin, 2) = ".\"
          strLogin  = strServer & Mid(strLogin, 2)
        Case Else
          strLogin  = strDomain & Mid(strLogin, Instr(strLogin, "\"))
      End Select
      strCmd        = "CREATE LOGIN [" & strLogin & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-r -Q", """" & strCmd & """", 1)
  End Select  

  Call SetBuildfileValue("SetupStdAccountsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigSAAccounts()
  Call SetProcessId("5CB", "Setup Sysadmin Accounts")

  strCmd        = "CREATE LOGIN [" & strGroupDBAAlt & "] FROM WINDOWS;"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case Instr(strErrMsg, "Msg 15401") > 0 And Left(strGroupDBAAlt, intServerLen + 1) = strServer & "\"
      strGroupDBAAlt = strBuiltinDom & Mid(strGroupDBA, Instr(strGroupDBA, "\"))
      Call SetBuildfileValue("GroupDBAAlt", strGroupDBAAlt)
      strCmd    = "CREATE LOGIN [" & strGroupDBAAlt & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
  End Select 

  strCmd            = "EXEC sp_addsrvrolemember '" & strGroupDBAAlt & "', 'sysadmin';"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strUserDNSDomain <> "" 
      strCmd        = strDomain & "\" & strUserName
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strCmd & "', 'sysadmin';""", 1)
    Case Else
      strCmd        = strServer & "\" & strUserName
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strCmd & "', 'sysadmin';""", 1)
  End Select

  Select Case True
    Case strSetupDRUClt <> "YES"
      ' Nothing
    Case strCltAccount = strNTAuthOSName 
      ' Nothing
    Case strCltAccount = strSqlAccount 
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """CREATE LOGIN [" & strCltAccount & "] FROM WINDOWS;""", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strCltAccount & "', 'sysadmin';""", 1)
      Call SetBuildfileValue("SetupDRUCltStatus", strStatusComplete)
  End Select

  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER LOGIN [sa] WITH PASSWORD='" & strsaPwd & "';""", 1)

  Call SetBuildfileValue("SetupSAAccountsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigNonSAAccounts()
  Call SetProcessId("5CC", "Setup DBA Non-Sysadmin Group")

  Call DebugLog("Create DBA Non-SA Account")
  strCmd            = "CREATE LOGIN [" & strGroupDBANonSAAlt & "] FROM WINDOWS;"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case Instr(strErrMsg, "Msg 15401") > 0 And Left(strGroupDBANonSAAlt, intServerLen + 1) = strServer & "\"
      strGroupDBANonSAAlt = strBuiltinDom & Mid(strGroupDBANonSA, Instr(strGroupDBANonSA, "\"))
      Call SetBuildfileValue("GroupDBANonSAAlt", strGroupDBANonSAAlt)
      strCmd        = "CREATE LOGIN [" & strGroupDBANonSAAlt & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
  End Select

  Call DebugLog("Add Server permissions")
  Select Case True
    Case strSQLVersion >= "SQL2012"
      strCmd        = "CREATE SERVER ROLE [" & strRoleDBANonSA & "] AUTHORIZATION [sysadmin];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", -1)
      strCmd        = "GRANT ALTER TRACE TO [" & strRoleDBANonSA & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT CONNECT SQL TO [" & strRoleDBANonSA & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW ANY DATABASE TO [" & strRoleDBANonSA & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW ANY DEFINITION TO [" & strRoleDBANonSA & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW SERVER STATE TO [" & strRoleDBANonSA & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      If strSQLVersion >= "SQL2014" Then
        strCmd      = "GRANT CONNECT ANY DATABASE TO [" & strRoleDBANonSA & "];"
        Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      End If
      strCmd        = "ALTER SERVER ROLE [" & strRoleDBANonSA & "] ADD MEMBER [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
    Case Else
      strCmd        = "GRANT ALTER TRACE TO         [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW ANY DATABASE TO   [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW ANY DEFINITION TO [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT VIEW SERVER STATE TO   [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
  End Select

  Call DebugLog("Add permissions for all databases")
  strCmd            = "EXECUTE master..sp_msForEachDB 'USE [?];CREATE USER [" & strGroupDBANonSAAlt & "] FOR LOGIN [" & strGroupDBANonSAAlt & "]';"
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 1)
  Select Case True
    Case strSQLVersion >= "SQL2012"
      strCmd        = "EXECUTE master..sp_msForEachDB 'USE [?];ALTER ROLE [db_datareader] ADD MEMBER [" & strGroupDBANonSAAlt & "]';"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
    Case Else
      strCmd        = "EXECUTE master..sp_msForEachDB 'USE [?];EXECUTE sp_addrolemember N''db_datareader'', N''" & strGroupDBANonSAAlt & "''';"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
  End Select

  Call DebugLog("Add specific permissions for msdb")
  strCmd            = "EXECUTE sp_addrole N'" & strRoleDBANonSA & "', N'dbo';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  strCmd            = "EXECUTE sp_addrolemember N'" & strRoleDBANonSA & "', N'" & strGroupDBANonSAAlt & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  strCmd            = "EXECUTE sp_addrolemember N'SQLAgentOperatorRole', N'" & strRoleDBANonSA & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  Select Case True
    Case strSQLVersion = "SQL2005"
      strCmd        = "EXECUTE sp_addrolemember N'db_dtsoperator', N'" & strRoleDBANonSA & "';"
      Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
    Case Else
      strCmd        = "EXECUTE sp_addrolemember N'db_ssisoperator', N'" & strRoleDBANonSA & "';"
      Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
      strCmd        = "EXECUTE sp_addrolemember N'ServerGroupReaderRole', N'" & strRoleDBANonSA & "';"
      Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  End Select
 
  Call SetBuildfileValue("SetupNonSAAccountsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigDisableSA()
  Call SetProcessId("5CD", "Setup sa Account")

  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER LOGIN [sa] DISABLE;""", 0)

  Select Case True
    Case Ucase(strsaName) = Ucase("sa") 
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER LOGIN [sa] WITH NAME=[" & strsaName & "];""", 1)
  End Select

  Call SetBuildfileValue("SetupDisableSAStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigCmdshell()
  Call SetProcessId("5CE", "Setup xp_cmdshell Proxy Account")

  If strCmdshellAccount = "" Then
    Call SetBuildfileValue("SetupCmdshellStatus", strStatusBypassed)
    Call DebugLog(" " & strProcessIdDesc & strStatusBypassed)
    Exit Sub
  End If

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'xp_cmdshell', '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_xp_cmdshell_proxy_account '" & strCmdshellAccount & "', '" & strCmdshellPassword & "';""", 0)

  Call SetBuildfileValue("SetupCmdshellStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigDBOwnerAccount()
  Call SetProcessId("5CF", "Setup Database Owner Account")

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strDBOwnerAccount = ""
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """CREATE CREDENTIAL StandardDBOwner WITH IDENTITY='" & strDBOwnerAccount & "';""", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """CREATE LOGIN [" & strDBOwnerAccount & "] WITH PASSWORD='" & strsaPwd & "', CHECK_POLICY=ON, CHECK_EXPIRATION=OFF, CREDENTIAL=StandardDBOwner;""", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """REVOKE CONNECT SQL TO [" & strDBOwnerAccount & "];""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER LOGIN [" & strDBOwnerAccount & "] DISABLE;""", 0)
  End Select

  Call SetBuildfileValue("SetupStdAccountsStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigUserAccounts()
  Call SetProcessId("5CG", "Setup User Accounts")
  Dim strLogin, strPassword

  For intIdx = 1 To strNumLogins
    strLogin        = GetBuildfileValue("WinLogin" & Right("0" & CStr(intIdx), 2))
    If strLogin <> "" Then
      Call Util_ExecSQL(strCmdSQL & "-Q", """CREATE LOGIN [" & strLogin & "] FROM WINDOWS;""", 1)
    End If
    strLogin        = GetBuildfileValue("UserLogin"  & Right("0" & CStr(intIdx), 2))
    strPassword     = GetBuildfileValue("UserPassword" & Right("0" & CStr(intIdx), 2))
    If strLogin <> "" And strPassword <> "" Then
      Call Util_ExecSQL(strCmdSQL & "-Q", """CREATE LOGIN [" & strLogin & "] WITH PASSWORD='" & strPassword & "';""", 1)
    End If
  Next

  Call SetBuildfileValue("SetupStdAccountsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigDBs()
  Call SetProcessId("5D", "Database configuration")

  Dim strSetupSysDB
  strSetupSysDB     = GetBuildfileValue("SetupSysDB")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5DAZ"
      ' Nothing
    Case strSetupSysDB <> "YES"
      ' Nothing
    Case Else
      Call ConfigSysDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5DB"
      ' Nothing
    Case strSetupSysDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call CopySysDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5DC"
      ' Nothing
    Case strSetupSysDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call RestartSQLServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5DD"
      ' Nothing
    Case strSetupSysDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSysDBStatus", strStatusPreConfig)
    Case Else
      Call TidySysDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DE"
      ' Nothing
    Case GetBuildfileValue("SetupSysIndex") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSysIndexStatus", strStatusPreConfig)
    Case Else
      Call ConfigSysIndex()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DF"
      ' Nothing
    Case strSetupSQLDBFT <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSQLDBFTStatus", strStatusPreConfig)
    Case Else
      Call ConfigSQLDBFT()
  End Select

  Call SetProcessId("5DZ", " Database configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigSysDB()
  Call SetProcessId("5DA", "Set System Database Options")
' Most of the work can be done while SQL Server is running normally.
' See 'Moving system Databases' in BOL for more details.

  Call SetBuildfileValue("SetupSysDBStatus", strStatusProgress)

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAA"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupModel()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAB"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupMSDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAC"
      ' Nothing
    Case strSetupTempDb <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupTempDbStatus", strStatusPreConfig)
    Case Else
      Call SetupTempDB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAD"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case UCase(Left(strRSInstallMode, 9)) = UCase("FilesOnly")
      ' Nothing
    Case Not objFSO.FileExists(strDirData & "\" & strRSDBName & ".MDF")
      ' Nothing
    Case Else
      Call SetupRS_DBs()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAE"
      ' Nothing
    Case strSetupDQ <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupDQ_DBs()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing 
    Case strProcessId > "5DAF"
      ' Nothing
    Case GetBuildfileValue("SetupPolyBase") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call SetupPolyBase_DBs()
  End Select

  Call SetProcessId("5DAZ", " System database setup" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupModel()
  Call SetProcessId("5DAA", "Set model Database Options")

  Call SetDBOptions("model", "", "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMSDB()
  Call SetProcessId("5DAB", "Set msdb Database Options")

  Call SetDBOptions("msdb", "200 MB", "50 MB")
  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE msdb MODIFY FILE (NAME=msdbdata, FILENAME = '" & strDirData & "\MSDBDATA.MDF');""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE msdb MODIFY FILE (NAME=msdblog,  FILENAME = '" & strDirLog  & "\MSDBLOG.LDF');""", 0)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupTempDB()
  Call SetProcessId("5DAC", "Set tempdb Database Options")

  strPathNew        = strDirTempData
  Call SetupFolder(strPathNew)

  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb MODIFY FILE (NAME=tempdev, FILENAME = '" & strDirTempData & "\TEMPDB.MDF', FILEGROWTH = " & strtempdbFile & ");""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb MODIFY FILE (NAME=tempdev, SIZE = " & strtempdbFile & ");""", 8)
  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb MODIFY FILE (NAME=templog, FILENAME = '" & strDirLog  & "\TEMPLOG.LDF', FILEGROWTH = " & strtempdbLogFile & ", MAXSIZE = UNLIMITED);""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb MODIFY FILE (NAME=templog, SIZE = " & strtempdbLogFile & ");""", 8)
 
  Select Case True
    Case strSQLTempdbFileCount < 2
      ' Nothing
    Case Else
      For intIdx = 2 To strSQLTempdbFileCount
        Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb ADD FILE (NAME=tempdev_" & CStr(intIdx) & ",  FILENAME = '" & strDirTempData & "\TEMPDB_" & CStr(intIdx) & ".NDF', FILEGROWTH = " & strtempdbFile & ", SIZE = " & strtempdbFile & ");""", 1)
      Next
  End Select

  If strSQLVersion >= "SQL2014" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE tempdb SET DELAYED_DURABILITY=FORCED;""", 1)
  End If

  If strSetMemOptHybridBP = "ON" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER SERVER CONFIGURATION SET MEMORY_OPTIMIZED HYBRID_BUFFER_POOL = ON;""", 0)
  End If
  If strSetMemOptTempdb = "ON" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER SERVER CONFIGURATION SET MEMORY_OPTIMIZED TEMPDB_METADATA = ON;""", 0)
  End If

  Call SetBuildfileValue("SetupTempDbStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRS_DBs()
  Call SetProcessId("5DAD", "Set SSRS Database Options")

  Call SetDBOptions(strRSDBName, "200 MB", "50 MB")

  Call SetDBOptions(strRSDBName & "TempDB", "200 MB", "50 MB")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDQ_DBs()
  Call SetProcessId("5DAE", "Set DQ Database Options")

  If objFSO.FileExists(strDirData & "\DQS_Main.MDF") Then
    Call SetDBOptions("DQS_Main", "200 MB", "50 MB")
  End If

  If objFSO.FileExists(strDirData & "\DQS_Projects.MDF") Then
    Call SetDBOptions("DQS_Projects", "200 MB", "50 MB")
  End If

  If objFSO.FileExists(strDirData & "\DQS_Staging_Data.MDF") Then
    Call SetDBOptions("DQS_Staging_Data", "200 MB", "50 MB")
  End If

  Call SetBuildfileValue("SetupSysDBStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPolyBase_DBs()
  Call SetProcessId("5DAF", "Set PolyBase Database Options")

  If objFSO.FileExists(strDirData & "\DWConfiguration.MDF") Then
    Call SetDBOptions("DWConfiguration", "", "")
  End If

  If objFSO.FileExists(strDirData & "\DWDiagnostics.MDF") Then
    Call SetDBOptions("DWDiagnostics", "", "")
  End If

  If objFSO.FileExists(strDirData & "\DWQueue.MDF") Then
    Call SetDBOptions("DWQueue", "", "")
  End If

  Call SetBuildfileValue("SetupSysDBStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetDBOptions(strDatabase, strMDFGrowth, strLogGrowth)
  Call DebugLog("SetDBOptions: " & strDatabase)

  Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET AUTO_CLOSE OFF WITH ROLLBACK IMMEDIATE;""", 0)

  Select Case True
    Case strDatabase = "msdb"
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET RECOVERY SIMPLE WITH ROLLBACK IMMEDIATE;""", 0)
    Case strDatabase = "SemanticsDB"
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET RECOVERY SIMPLE WITH ROLLBACK IMMEDIATE;""", 0)
    Case strSetupAlwaysOn <> "YES"
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET RECOVERY SIMPLE WITH ROLLBACK IMMEDIATE;""", 0)
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET RECOVERY FULL   WITH ROLLBACK IMMEDIATE;""", 0)
  End Select

  Select Case True
    Case strSetupSnapshot <> "YES"
      ' Nothing
    Case strDatabase = "model"
      ' Nothing
    Case strDatabase = "msdb"
      ' Nothing
    Case Left(strDatabase, Len(strRSDBName)) = strRSDBName
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET ALLOW_SNAPSHOT_ISOLATION ON;""", 0)
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET READ_COMMITTED_SNAPSHOT  ON WITH ROLLBACK IMMEDIATE;""", 0)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2014"
      ' Nothing
    Case strDatabase = "model"
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] SET DELAYED_DURABILITY=FORCED WITH ROLLBACK IMMEDIATE;""", -1)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strDatabase = "model"
      ' Nothing
    Case strDatabase = "msdb"
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """ALTER DATABASE [" & strDatabase & "] SET MIXED_PAGE_ALLOCATION OFF;""", -1)
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """ALTER DATABASE [" & strDatabase & "] MODIFY FILEGROUP [primary] AUTOGROW_ALL_FILES;""", -1)
  End Select

  Select Case True
    Case strMDFGrowth = ""
      ' Nothing
    Case strDatabase = "msdb"
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] MODIFY FILE (NAME=msdbdata,                 FILEGROWTH = " & strMDFGrowth & ", MAXSIZE = UNLIMITED);""", 0)
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] MODIFY FILE (NAME=" & strDatabase & ",      FILEGROWTH = " & strMDFGrowth & ", MAXSIZE = UNLIMITED);""", 0)
  End Select

  Select Case True
    Case strLogGrowth = ""
      ' Nothing
    Case strDatabase = "msdb"
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] MODIFY FILE (NAME=msdblog,                  FILEGROWTH = " & strLogGrowth & ", MAXSIZE = UNLIMITED);""", 0)
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strDatabase & "] MODIFY FILE (NAME=" & strDatabase & "_log,  FILEGROWTH = " & strLogGrowth & ", MAXSIZE = UNLIMITED);""", 0)
  End Select

End Sub


Sub CopySysDB()
  Call SetProcessId("5DB", "Copy System Database Files")

  strPathOld        = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg0")
  strPathOld        = Mid(strPathOld, 3, Len(strPathOld) - 12)

  Call StopSQLServer()
  Wscript.Sleep strWaitMed

  Call CopyDBMdf("msdb",         strPathOld, strDirData, "msdbdata.mdf")
  Call CopyDBLog("msdb",         strPathOld, strDirLog,  "msdblog.ldf")

  Call DebugLog("Backing up System Database files")
  strPathLog        = GetPathLog("")
  strCmd            = "%COMSPEC% /D /C CSCRIPT.EXE """ & strPathFBScripts & "SqlSysDBCopy.vbs"" >> " & strPathLog
  Call Util_RunExec(strCmd, "", "", 0)

  Call SetBuildfileValue("SetupSysDBStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CopyDBMdf(strCopyDB, strPathOld, strPathNew, strCopyMDF)
  Call DebugLog("CopyDBMdf: " & strCopyDB)
  Dim strPathDB

  strPath           = strPathOld
  If Right(strPath, 1) <> "\" Then
    strPath         = strPath & "\"
  End If
  strPath           = strPath & strCopyMDF
  strDebugMsg1      = "Source: " & strPath

  strPathDB         = strPathNew
  If Right(strPathDB, 1) <> "\" Then
    strPathDB       = strPathDB & "\"
  End If
  strDebugMsg2      = "Target: " & strPathDB

  Select Case True
    Case Not objFSO.FileExists(strPath)
      ' Nothing
    Case Else
      Set objFile   = objFSO.GetFile(strPath)
      strPath       = strPathDB & objFile.Name
      objFile.Copy strPath
  End Select

End Sub


Sub CopyDBLog(strCopyDB, strPathOld, strPathNew, strCopyLog)
  Call DebugLog("CopyDBLog: " & strCopyDB)
  Dim strPathDB

  strPath           = strPathOld
  If Right(strPath, 1) <> "\" Then
    strPath         = strPath & "\"
  End If
  strPath           = strPath & strCopyLog
  strDebugMsg1      = "Source: " & strPath

  strPathDB         = strPathNew
  If Right(strPathDB, 1) <> "\" Then
    strPathDB       = strPathDB & "\"
  End If
  strDebugMsg2      = "Target: " & strPathDB

  Select Case True
    Case Not objFSO.FileExists(strPath)
      ' Nothing
    Case Else
      Set objFile   = objFSO.GetFile(strPath)
      strPath       = strPathDB & objFile.Name
      objFile.Copy strPath
  End Select

End Sub


Sub RestartSQLServer()
  Call SetProcessId("5DC", "Restart SQL Server Services")

  If strSetupSQLDB = "YES" Then
    Call StartSQL()
  End If

  If strSetupSQLDBAG = "YES" Then
    Call StartSQLAgent()
  End If

  If strSetupSQLRS = "YES" Then
    Call StartSSRS("")
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub TidySysDB()
  Call SetProcessId("5DD", "Tidy System Database Files")

  Call DebugLog("Removing msdb DB files")
  strPathOld        = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\MSSQLServer\Parameters\SQLArg0")
  strPathOld        = Mid(strPathOld, 3, Len(strPathOld) - 12)
  Call DeleteDBFile(strPathOld & "msdbdata.mdf")
  Call DeleteDBFile(strPathOld & "msdblog.ldf")

  Select Case True
    Case strSetupTempDb <> "YES"
      ' Nothing
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case Else
      Call DebugLog("Removing tempdb DB files")
      Call DeleteDBFile(strPathOld & "tempdb.mdf")
      Call DeleteDBFile(strPathOld & "templog.ldf")
  End Select

  Call SetBuildfileValue("SetupSysDBStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub DeleteDBFile(strFile)
  Call DebugLog("DeleteDBFile: " & strFile)

  If objFSO.FileExists(strFile) Then
    Set objFile = objFSO.GetFile(strFile)
    objFile.Delete(True)
  End If

End Sub


Sub ConfigSysIndex()
  Call SetProcessId("5DE", "Setup Extra System indexes")
  ' Code based on posting by Alex Wilson on HTTP://IPTDBA.BLOGSPOT.COM/
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",                strPathFBScripts)
  Call RunInstall("SysIndex",  "Set-SystemDBOptions.sql", objInstParm)

  Call ProcessEnd("")

End Sub


Sub ConfigSQLDBFT()
  Call SetProcessId("5DF", "Setup Full Text Search")

  If strSQLVersion = "SQL2005" Then
    Call DebugLog("Copy Full Text data to new location")
    strPath         = strHKLMSQL & strInstRegSQL & "\Setup\SQLDataRoot"  
    strPathOld      = objShell.RegRead(strPath) & "\FTData"
    strCmd          = "XCOPY """ & strPathOld & "\*.*"" """ & strDirDataFT & "\*.*"" /C /E /H /K /O /R /V /X /Y"
    Call Util_RunExec(strCmd, "", "", 0)
    Call DebugLog("Delete original Full Text Data")
    objFSO.DeleteFolder(strPathOld)
    Call DebugLog("Update SQL registry with new FT location")
    strPath         = strHKLMSQL & strInstRegSQL & "\Setup\FullTextDefaultPath"
    Call Util_RegWrite(strPath, strDirDataFT, "REG_SZ") 
  End If

  Call DebugLog("Setup Full Text SQL Configuration")
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'max full-text crawl range', '" & Cstr(intMaxDop) & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  Call SetBuildfileValue("SetupSQLDBFTStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigManagement()
  Call SetProcessId("5E", "SQL Management configuration")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EAZ"
      ' Nothing
    Case GetBuildfileValue("SetupSysManagement") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupSysManagementStatus", strStatusPreConfig)
    Case Else
      Call ConfigSysManagement()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EBZ"
      ' Nothing
    Case Else
      Call SetupInstanceManagement()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5ECZ"
      ' Nothing
    Case GetBuildfileValue("SetupDBAManagement") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupDBAManagementStatus", strStatusPreConfig)
    Case Else
      Call SetupDBAManagement()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDZ"
      ' Nothing
    Case GetBuildfileValue("SetupManagementDW") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupManagementDWStatus", strStatusPreConfig)
    Case Else
      Call ConfigManagementDW()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EE"
      ' Nothing
    Case GetBuildfileValue("SetupPBM") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupPBMStatus", strStatusPreConfig)
    Case Else
      Call ConfigPBM()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EF"
      ' Nothing
    Case GetBuildfileValue("SetupGenMaint") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupGenMaintStatus", strStatusPreConfig)
    Case Else
      Call ConfigGenMaint()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EG"
      ' Nothing
    Case GetBuildfileValue("SetupGovernor") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupGovernorStatus", strStatusPreConfig)
    Case Else
      Call ConfigGovernor()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EH"
      ' Nothing
    Case GetBuildfileValue("SetupDBOpts") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call ConfigDBOpts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EI"
      ' Nothing
    Case GetBuildfileValue("SetupDBOpts") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupDBOptsStatus", strStatusPreConfig)
    Case Else
      Call ConfigJobs()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EJZ"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case Else
      Call ConfigureAO()
  End Select

  Call SetProcessId("5EZ", " SQL Management configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigSysManagement()
  Call SetProcessId("5EA", "System DB Management")
  Dim objInstParm

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EAA"
      ' Nothing
    Case Else
      Call SetupFBSysManagement()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EAB"
      ' Nothing
    Case GetBuildfileValue("SetupStartJob") <> "YES"
      ' Nothing
    Case Else
      Call SetupStartJobProxy()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EAC"
      ' Nothing
    Case GetBuildfileValue("ClusterHost") <> "YES"
      ' Nothing
    Case Else
      Call SetupClusterGroup()
  End Select

  Call SetBuildfileValue("SetupSysManagementStatus", strStatusComplete)
  Call SetProcessId("5EAZ", " System DB Management" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupFBSysManagement()
  Call SetProcessId("5EAA", "Setup FB System Management Routines")
  Dim objInstParm

  Call SetXMLParm(objInstParm,     "PathMain",     strPathFBScripts)
  Call SetXMLParm(objInstParm,     "InstFile",     GetBuildfileValue("SysManagementbat"))
  Call SetXMLParm(objInstParm,     "ParmXtra",     strServInst)
  Call SetXMLParm(objInstParm,     "LogXtra",      "SysManagement")
  Call RunInstall("SysManagement", GetBuildfileValue("SysManagementCab"), objInstParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupStartJobProxy()
  Call SetProcessId("5EAB", "Setup Start Job Proxy")
  Dim objInstParm

  Call SetXMLParm(objInstParm,     "PathMain",              strPathFBScripts)
  Call SetXMLParm(objInstParm,     "ParmXtra",              "-v strStartJobPassword=""" & GetBuildfileValue("StartJobPassword") & """ strDirSystemDataShared=""" & strDirSystemDataShared & """")
  Call SetXMLParm(objInstParm,     "LogClean",              "Y")
  Call RunInstall("StartJob",      "Set-StartJobProxy.sql", objInstParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupClusterGroup()
  Call SetProcessId("5EAC", "Setup Cluster Group")
  Dim objInstParm

  Call SetXMLParm(objInstParm,     "PathMain",             strPathFBScripts)
  Call SetXMLParm(objInstParm,     "LogXtra",              "Set-ClusterGroup")
  Call RunInstall("SysManagement", "Set-ClusterGroup.sql", objInstParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupInstanceManagement()
  Call SetProcessId("5EB", "Setup Instance Management")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EBA"
      ' Nothing
    Case GetBuildfileValue("SetupBPE") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupBPEStatus", strStatusPreConfig)
    Case Else
      Call SetupBPE()
  End Select

  Call SetProcessId("5EBZ", " Setup Instance Management" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupBPE()
  Call SetProcessId("5EBA", "Setup BPE File")

  strCmd            = "ALTER SERVER CONFIGURATION SET BUFFER POOL EXTENSION ON (FILENAME = '" & strDirBPE & "\" & Replace(strBPEFile, " ", "") & ".BFE',SIZE = " & strBPEFile & ");"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupBPEStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDBAManagement()
  Call SetProcessId("5EC", "Setup DBA Management")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5ECA"
      ' Nothing
    Case Else
      Call CreateDBA_DB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5ECB"
      ' Nothing
    Case Else
      Call SetupDBA_DB()
  End Select

  Call SetBuildfileValue("SetupDBAManagementStatus", strStatusComplete)
  Call SetProcessId("5ECZ", " Setup DBA Management" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CreateDBA_DB()
  Call SetProcessId("5ECA", "Create DBA DB " & strDBA_DB)
  Dim objInstParm

  Call SetXMLParm(objInstParm,     "PathMain",      strPathFBScripts)
  Call SetXMLParm(objInstParm,     "ParmXtra",      "-v strDBA_DB=""" & strDBA_DB & """")
  Call RunInstall("DBAManagement", "Set-DBADB.sql", objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupDBA_DB()
  Call SetProcessId("5ECB", "Setup DBA Management Routines")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",     strPathFBScripts)
  Call SetXMLParm(objInstParm, "InstFile",     GetBuildfileValue("DBAManagementbat"))
  Call SetXMLParm(objInstParm, "ParmXtra",     strServInst)
  Call RunInstall("DBAManagement", GetBuildfileValue("DBAManagementCab"), objInstParm)

  Call ProcessEnd("")
End Sub


Sub ConfigManagementDW()
  Call SetProcessId("5ED", "Management Data Warehouse setup")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDA"
      ' Nothing
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      Call CreateManagementDW()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDB"
      ' Nothing
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      Call SetupManagementDWSchema()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDC"
      ' Nothing
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      Call SetupManagementDWAuth()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDD"
      ' Nothing
    Case Else
      Call SetupMsdbAuth()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDE"
      ' Nothing
    Case Else
      Call SetupManagementDWCollection()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDF"
      ' Nothing
    Case Else
      Call SetupManagementDWJobs()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDG"
      ' Nothing
    Case strMDWAccount = "" Or strMDWPassword = ""
      ' Nothing
    Case Else
      Call SetupMDWProxy()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDH"
      ' Nothing
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      Call SetupMDWIndexes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EDI"
      ' Nothing
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case strEditionEnt <> "YES" 
      ' Nothing
    Case Else
      Call SetupMDWCompression()
  End Select

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusComplete)
  Call SetProcessId("5EDZ", " Management Data Warehouse setup" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CreateManagementDW()
  Call SetProcessId("5EDA", "Create " & strManagementDW & " DB")
  Dim objInstParm

  Call SetXMLParm(objInstParm,    "PathMain",               strPathFBScripts)
  Call SetXMLParm(objInstParm,    "ParmXtra",               "-v strManagementDW=""" & strManagementDW & """")
  Call SetXMLParm(objInstParm,    "LogXtra",                "Set-ManagementDWDB")
  Call RunInstall("ManagementDW", "Set-ManagementDWDB.sql", objInstParm)

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupManagementDWSchema()
  Call SetProcessId("5EDB", "Setup " & strManagementDW & " Schema")
  Dim objInstParm

  Call SetXMLParm(objInstParm,    "PathMain",     objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLPath") & "\Install")
  Call SetXMLParm(objInstParm,    "ParmXtra",     "-d """ & strManagementDW & """ -x ")
  Call SetXMLParm(objInstParm,    "LogXtra",      "instmdw")
  Call RunInstall("ManagementDW", "instmdw.sql",  objInstParm)

  If GetBuildfileValue("SetupManagementDWStatusinstmdw") <> strStatusComplete Then
    Exit Sub
  End If

  If strSetupAlwaysOn <> "YES" Then
    Call Util_ExecSQL(strCmdSQL & " -Q", """ALTER DATABASE [" & strManagementDW & "] SET RECOVERY SIMPLE;""", 0)
  End If
  If strSetupSnapshot = "YES" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strManagementDW & "] SET ALLOW_SNAPSHOT_ISOLATION ON;""", 0)
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strManagementDW & "] SET READ_COMMITTED_SNAPSHOT ON WITH ROLLBACK IMMEDIATE;""", 0)
  End If
  If strSQLVersion >= "SQL2014" Then
    Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER DATABASE [" & strManagementDW & "] SET DELAYED_DURABILITY=FORCED;""", 1)
  End If

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupManagementDWAuth()
  Call SetProcessId("5EDC", "Setup " & strManagementDW & " DB Authorities")

  Call DebugLog("Authorities for DBA sysadmin account")
  strCmd            = "CREATE USER [" & strGroupDBAAlt & "] FOR LOGIN [" & strGroupDBAAlt & "];"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 1)
  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_admin', @MEMBERNAME='" & strGroupDBAAlt & "';"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_Writer', @MEMBERNAME='" & strGroupDBAAlt & "';"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_reader', @MEMBERNAME='" & strGroupDBAAlt & "';"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  Call DebugLog("Authorities for DBA non-sysadmin account")
  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      strCmd        = "CREATE USER [" & strGroupDBANonSAAlt & "] FOR LOGIN [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 1)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_reader', @MEMBERNAME='" & strGroupDBANonSAAlt & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  End Select

  Call DebugLog("Authorities for SQL Server service account")
  Select Case True
    Case Ucase(strSqlAccount) = UCase(strNTAuthOSName)
      ' Nothing
    Case Else
      strCmd        = "CREATE USER [" & strSqlAccount & "] FOR LOGIN [" & strSqlAccount & "] WITH DEFAULT_SCHEMA=dbo;"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_writer', @MEMBERNAME='" & strSqlAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_reader', @MEMBERNAME='" & strSqlAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT ADMINISTER BULK OPERATIONS TO [" & strSqlAccount & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
  End Select

  Call DebugLog("MDW reports authority")
  strCmd            = "GRANT VIEW DEFINITION TO [mdw_admin];"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "GRANT VIEW DEFINITION TO [mdw_reader];"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMsdbAuth()
  Call SetProcessId("5EDD", "Setup " & strManagementDW & " msdb DB Authorities")

  Call DebugLog("Create Report Reader Role") ' See Connect 558417
  strCmd            = "CREATE ROLE [dc_report_reader] AUTHORIZATION [dbo];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  strCmd            = "GRANT SELECT ON [dbo].[syscollector_collection_sets] TO [dc_report_reader];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  strCmd            = "GRANT SELECT ON [dbo].[syscollector_execution_log] TO [dc_report_reader];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  strCmd            = "GRANT SELECT ON [dbo].[syscollector_config_store] TO [dc_report_reader];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)

  Call DebugLog("Authorities for DBA sysadmin account")
  strCmd            = "CREATE USER [" & strGroupDBAAlt & "] FOR LOGIN [" & strGroupDBAAlt & "];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='dc_report_reader', @MEMBERNAME='" & strGroupDBAAlt & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)

  Call DebugLog("Authorities for DBA non-sysadmin account")
  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      strCmd        = "CREATE USER [" & strGroupDBANonSAAlt & "] FOR LOGIN [" & strGroupDBANonSAAlt & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='dc_report_reader', @MEMBERNAME='" & strGroupDBANonSAAlt & "';"
      Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  End Select

  Call DebugLog("Authorities for MDW dc_proxy")
  strCmd            = "GRANT EXECUTE ON [dbo].[sp_syscollector_sql_text_lookup] TO [dc_proxy];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  strCmd            = "GRANT EXECUTE ON [dbo].[sp_syscollector_text_query_plan_lookpup] TO [dc_proxy];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupManagementDWCollection()
  Call SetProcessId("5EDE", "Setup " & strManagementDW & " Data Collection")
  Dim objSQL, objSQLData
  Dim strColParameters, strColUID

  strColParameters  = "<ns:QueryActivityCollector xmlns:ns=""""DataCollectorType""""><Databases IncludeSystemDatabases=""""false"""" /></ns:QueryActivityCollector>"
  strColUID         = "2DC02BD6-E230-4C05-8516-4E8C0EF21F95"
  Set objSQL        = CreateObject("ADODB.Connection")
  Set objSQLData    = CreateObject("ADODB.Recordset")
  objSQL.Provider   = "SQLOLEDB"
  objSQL.ConnectionString = "Driver={SQL Server};Server=" & strServInst & ";Database=master;Trusted_Connection=Yes;"
  strDebugMsg1      = objSQL.ConnectionString
  objSQL.Open 

  Call DebugLog("Set Management Data Warehouse name for Data Collectors")

  strCmd            = "EXEC dbo.sp_syscollector_disable_collector;"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)

  strCmd            = "EXEC dbo.sp_syscollector_set_warehouse_instance_name @instance_name='" & strManagementServerName
  Select Case True
    Case strManagementInstance = ""
      strCmd        = strCmd & "'"
    Case strManagementInstance = "MSSQLSERVER"
      strCmd        = strCmd & "'"
    Case Else
      strCmd        = strCmd & "\" & strManagementInstance & "'"
  End Select
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & ";""", 0)

  strCmd            = "EXEC dbo.sp_syscollector_set_warehouse_database_name @database_name='" & strManagementDW & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)

  If GetBuildfileValue("DirMDW") <> "" Then
    strCmd          = "EXEC dbo.sp_syscollector_set_cache_directory @cache_directory='" & GetBuildfileValue("DirMDW") & "';"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  End If

  Call DebugLog("Start Data Collectors")

  If strSQLVersion >= "SQL2008R2" Then
    strcmd          = "EXEC dbo.sp_syscollector_enable_collector;"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  End If

  Wscript.Sleep strWaitLong 
  strCmd            = "SELECT collection_set_id FROM msdb.dbo.syscollector_collection_sets_internal WHERE schedule_uid IS NOT NULL ORDER BY collection_set_id;"
  Set objSQLData    = objSQL.Execute(strCmd)
  Do Until objSQLData.EOF
    intIdx          = objSQLData.Fields("collection_set_id")
    strcmd          = "EXEC dbo.sp_syscollector_start_collection_set @collection_set_id=" & Cstr(intIdx) & ";"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
    objSQLData.MoveNext
  Loop

  If strSQLVersion = "SQL2008" Then
    strcmd          = "EXEC dbo.sp_syscollector_enable_collector;"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  End If

  Call DebugLog("Disable System DB data collection: KB2793580")

  strCmd            = "SELECT collection_item_id "
  strCmd            = strCmd & "FROM msdb.dbo.syscollector_collection_items_internal AS CI "
  strCmd            = strCmd & "JOIN msdb.dbo.syscollector_collection_sets AS CS "
  strCmd            = strCmd & "ON CS.collection_set_id = CI.collection_set_id "
  strCmd            = strCmd & "WHERE CS.collection_set_uid = N'" & strColUid & "';"
  Set objSQLData    = objSQL.Execute(strCmd)
  objSQLData.MoveFirst
  strColUid         = objSQLData.Fields("collection_item_id")

  strcmd            = "EXEC dbo.sp_syscollector_update_collection_item "
  strCmd            = strCmd & " @collection_item_id = " & strColUID
  strCmd            = strCmd & ",@parameters = '" & strColParameters & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", -1)

  objSQLData.Close
  objSQL.Close

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupManagementDWJobs()
  Call SetProcessId("5EDF", "Setup " & strManagementDW & " Job Names")

  strcmd            = "EXEC " & strDBA_DB & ".dbo.spResetJobnamesMDW;"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMDWProxy()
  Call SetProcessId("5EDG", "Setup MDW Proxy Account")
  Dim strCredential

  intIdx            = Instr(strMDWAccount, "\")
  strCredential     = Mid(strMDWAccount, intIdx + 1)

  Call DebugLog("Setup Proxy Windows Security")  
  strCmd            = "NET LOCALGROUP """ & GetBuildfileValue("GroupDistComUsers") & """ """ & strMDWAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", "", -1)

  Call DebugLog("Setup Proxy SQL Security")  
  strCmd            = "CREATE LOGIN [" & strMDWAccount & "] FROM WINDOWS WITH DEFAULT_DATABASE = [master], DEFAULT_LANGUAGE = [us_english];"
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 1)
  strCmd        = "CREATE USER [" & strMDWAccount & "] FOR LOGIN [" & strMDWAccount & "];"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
  strCmd            = "EXEC SP_ADDROLEMEMBER @ROLENAME='dc_proxy', @MEMBERNAME='" & strMDWAccount & "';"
  Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  strCmd            = "EXEC sp_addsrvrolemember '" & strMDWAccount & "', 'sysadmin';"
  Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 1)

  Select Case True
    Case Instr(strManagementServerList, " " & strManagementServerName & " ") = 0
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      strCmd        = "CREATE USER [" & strMDWAccount & "] FOR LOGIN [" & strMDWAccount & "] WITH DEFAULT_SCHEMA=dbo;"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_writer', @MEMBERNAME='" & strMDWAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "EXEC SP_ADDROLEMEMBER @ROLENAME='mdw_reader', @MEMBERNAME='" & strMDWAccount & "';"
      Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
      strCmd        = "GRANT ADMINISTER BULK OPERATIONS TO [" & strMDWAccount & "];"
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 0)
  End Select

  If strMDWAccount <> strAgtAccount Then
    Call DebugLog("Create SQL Agent Proxy")
    strCmd          = "CREATE CREDENTIAL [" & strCredential & "] WITH IDENTITY = N'" & strMDWAccount & "', SECRET = N'" & strMDWPassword & "';"
    Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strCmd & """", 1)
    strCmd          = "EXEC sp_adAPP_d_proxy @proxy_name=N'" & strCredential & "', @credential_name=N'" & strCredential & "', @enabled=1, @description=N'MDW Proxy';"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
    strCmd          = "EXEC sp_grant_proxy_to_subsystem @proxy_name=N'" & strCredential & "', @subsystem_id=3;"  ' CmdExec
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)
    strCmd          = "EXEC sp_grant_proxy_to_subsystem @proxy_name=N'" & strCredential & "', @subsystem_id=11;" ' SSIS
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)

    strCmd          = "EXEC sp_grant_login_to_proxy @proxy_name=N'" & strCredential & "', @msdb_role=N'dc_admin'"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 1)

    Call DebugLog("Apply Proxy to SQL Agent MDW Jobs")
    strCmd          = "UPDATE sysjobsteps SET "
    strCmd          = strCmd & " proxy_id = p.proxy_id "
    strCmd          = strCmd & "FROM sysjobsteps s "
    strCmd          = strCmd & "INNER JOIN sysjobs j "
    strCmd          = strCmd & "   ON j.job_id      = s.job_id "
    strCmd          = strCmd & "INNER JOIN syscategories c "
    strCmd          = strCmd & "   ON j.category_id = c.category_id "
    strCmd          = strCmd & "  AND c.name        = 'Data Collector' "
    strCmd          = strCmd & " LEFT JOIN sysproxies p "
    strCmd          = strCmd & "   ON p.name        = '" & strCredential & "' "
    strCmd          = strCmd & "WHERE s.subsystem   = 'CMDEXEC' "
    strCmd          = strCmd & "  AND s.proxy_id    IS NULL;"
    Call Util_ExecSQL(strCmdSQL & "-d ""msdb"" -Q", """" & strCmd & """", 0)
  End If

  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMDWIndexes()
  Call SetProcessId("5EDH", "Setup " & strManagementDW & " Indexes")

  strCmd            = "CREATE INDEX [IDX_query_stats_sql_handle#FB] ON [snapshots].[query_stats] "
  strCmd            = strCmd & "([sql_handle] ASC,[snapshot_id] ASC) ON [PRIMARY];"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "CREATE UNIQUE INDEX [IDX_notable_query_text_sql_handle#FB] ON [snapshots].[notable_query_text] "
  strCmd            = strCmd & "([sql_handle] ASC,[source_id] ASC) ON [PRIMARY];"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
 
  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMDWCompression()
  Call SetProcessId("5EDI", "Setup MDW Table Compression")

  strCmd            = "ALTER TABLE [snapshots].[query_stats] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[query_stats] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[active_sessions_and_requests] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[active_sessions_and_requests] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[notable_query_plan] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[notable_query_plan] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[os_wait_stats] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[os_wait_stats] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[os_memory_clerks] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[os_memory_clerks] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[notable_query_text] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[notable_query_text] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[io_virtual_file_stats] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[io_virtual_file_stats] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [core].[snapshots_internal] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [core].[snapshots_internal] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)

  strCmd            = "ALTER TABLE [snapshots].[disk_usage] REBUILD PARTITION=ALL WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
  strCmd            = "ALTER INDEX ALL ON [snapshots].[disk_usage] REBUILD WITH (DATA_COMPRESSION=PAGE);"
  Call Util_ExecSQL(strCmdSQL & "-d """ & strManagementDW & """ -Q", """" & strCmd & """", 0)
 
  Call SetBuildfileValue("SetupManagementDWStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigPBM()
  Call SetProcessId("5EE", "Setup Policy Based Maintenance")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",     strPathFBScripts)
  Call SetXMLParm(objInstParm, "InstFile",     GetBuildfileValue("PBMBat"))
  Call SetXMLParm(objInstParm, "ParmXtra",     strServInst)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("PBM", GetBuildfileValue("PBMCab"), objInstParm)

  If GetBuildfileValue("SetupPBMStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog(" Updating PBM Job details")
  strcmd            = "EXEC " & strDBA_DB & ".dbo.spResetJobnamesPBM;"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupPBMStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigGenMaint()
  Call SetProcessId("5EF", "Setup Generic Maintenance")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call SetXMLParm(objInstParm, "InstFile",    GetBuildfileValue("GenMaintVbs"))
  Call RunInstall("GenMaint", GetBuildfileValue("GenMaintCab"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub ConfigGovernor()
  Call SetProcessId("5EG", "Resource Governor setup")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call RunInstall("Governor", GetBuildfileValue("GovernorSql"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub ConfigDBOpts()
  Call SetProcessId("5EH", "Apply standard DB options")

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC " & strDBA_DB & "..spSetDBOptions;""", 0)

  Call SetBuildfileValue("SetupDBOptsStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigJobs()
  Call SetProcessId("5EI", "Apply standard Job corrections")

  strCmd            = "UPDATE sysjobsteps SET "
  strCmd            = strCmd & " command=REPLACE(command, '" & strServer & "', '$(ESCAPE_DQUOTE(SRVR))');"
  Call Util_ExecSQL(strCmdSQL & "-x -d ""msdb"" -Q", """" & strCmd & """", 0)

  Call SetBuildfileValue("SetupDBOptsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigureAO()
  Call SetProcessId("5EJ", "Configure AlwaysOn Processes")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EJA"
      ' Nothing
    Case GetBuildfileValue("SetupAODB") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupAODBStatus", strStatusPreConfig)
    Case Else
      Call SetupAODB()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5EJB"
      ' Nothing
    Case GetBuildfileValue("SetupAOProcs") <> "YES"
      ' Nothing
    Case Else
      Call SetupAOProcs()
  End Select

  Call SetProcessId("5EJZ", " Configure AlwaysOn Processes" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupAODB()
  Call SetProcessId("5EJA", "Setup DBs for AlwaysOn")
  Dim objSQL, objSQLData
  Dim strDBName, strDBList

  Set objSQL        = CreateObject("ADODB.Connection")
  Set objSQLData    = CreateObject("ADODB.Recordset")
  objSQL.Provider   = "SQLOLEDB"
  objSQL.ConnectionString = "Driver={SQL Server};Server=" & strServInst & ";Database=master;Trusted_Connection=Yes;"
  strDebugMsg1      = objSQL.ConnectionString
  objSQL.Open 

  strDBList         = "name NOT IN ('master', 'model', 'msdb', 'tempdb', 'SemanticsDB', 'SSISDB', '" & strDBA_DB & "')"
  Select Case True
    Case strActionAO = "ADDNODE"
      strCmd        = "SELECT AG.name AS AGname,DB.database_name AS name FROM master.sys.availability_groups AS AG "
      strCmd        = strCmd & "JOIN master.sys.availability_databases_cluster DB ON DB.group_id = AG.group_id "
      strCmd        = strCmd & "WHERE " & strDBList & " ORDER BY AG.name,DB.database_name"
    Case Else
      strCmd        = "SELECT name FROM master.sys.sysdatabases AS d "
      strCmd        = strCmd & "JOIN master.sys.database_recovery_status AS s ON s.database_id = d.dbid "
      strCmd        = strCmd & "LEFT JOIN master.sys.availability_databases_cluster AS a ON a.database_name = d.name "
      strCmd        = strCmd & "WHERE Has_DBAccess(name) > 0 AND s.last_log_backup_lsn IS NOT NULL AND a.database_name IS NULL AND " & strDBList &" ORDER BY name"
  End Select
  strDebugMsg1      = strCmd
  Set objSQLData    = objSQL.Execute(strCmd)
  Do While Not objSQLData.EOF
    strDBName       = objSQLData.Fields("name")
    Select Case True
      Case(strDBName = strSSISDB) And (strSQLVersion <= "SQL2014")
        ' Nothing
      Case strActionAO = "ADDNODE"
        Call SetupAODBSecondary(strDBName, objSQLData.Fields("AGname"))
      Case Else
        Call SetupAODBPrimary(strDBName, strGroupAO)
    End Select
    objSQLData.MoveNext
  Loop

  objSQLData.Close
  objSQL.Close

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAODBPrimary(strDBName, strGroupAO)

  strCmd            = "ALTER AVAILABILITY GROUP [" & strGroupAO & "] ADD DATABASE [" & strDBName & "];"
  Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)

  If strDBName = strSSISDB Then
    strCmd          = "EXEC sp_control_dbmasterkey_password @db_name = N'" & strDBName & "', @password = N'" & strSSISPassword & "', @action = N'add';"
    Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 0)
  End If

End Sub


Sub SetupAODBSecondary(strDBName, strGroupAO)
  Call DebugLog("ConnectAOSecondary: " & strDBName & " to " & strGroupAO)

  If sreSQLVersion <= "SQL2016" Then
    strCmd          = """" & strPathFBScripts & "Set-AODBSecondary.bat"" """ & GetBuildfileValue("PathCmdSQL") & """ " & strGroupAO & " " & strDBName & " " & strSSISDB & " " & strSSISPassword & " > " & GetPathLog(strDBName)
    Call Util_RunCmdAsync(strCmd, 0)
  End If

End Sub


Sub SetupAOProcs()
  Call SetProcessId("5EJB", "Setup AlwaysOn Procedures")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",             strPathFBScripts)
  Call SetXMLParm(objInstParm, "LogXtra",              "Set-AOFailover")
  Call RunInstall("AODB",      "Set-AOFailover.sql",   objInstParm)

  Call SetXMLParm(objInstParm, "PathMain",             strPathFBScripts)
  Call SetXMLParm(objInstParm, "LogXtra",              "Set-AOSystemData")
  Call SetXMLParm(objInstParm, "ParmXtra",             "-v strAccount=""" & strAgtAccount & """")
  Call RunInstall("AODB",      "Set-AOSystemData.sql", objInstParm)

  Call SetBuildfileValue("SetupAOProcsStatus", strStatusComplete)

End Sub


Sub ConfigWindows()
  Call SetProcessId("5F", "Windows configuration")

  Dim strSetupMenus
  strSetupMenus     = GetBuildfileValue("SetupMenus")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5FA"
      ' Nothing
    Case strSetupMenus <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLMenus()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5FB"
      ' Nothing
    Case strSetupMenus <> "YES"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case UCase(Left(strRSInstallMode, 9)) = UCase("FilesOnly")
      ' Nothing
    Case Else
      Call SetupSQLRSMenus()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5FC"
      ' Nothing
    Case strSetupMenus <> "YES"
      ' Nothing
    Case Else
      Call SetupWindowsMenus()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5FD"
      ' Nothing
    Case strSetupMenus <> "YES"
      ' Nothing
    Case Else
      Call MergeWindowsMenus()
  End Select

  Call SetProcessId("5FZ", " Windows configuration" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSQLMenus()
  Call SetProcessId("5FA", "SQL menu items")

  strPath           = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSSMS & ".lnk"
  strDebugMsg1      = "Source path: " & strPath
  If objFSO.FileExists(strPath) Then
    Call DebugLog("SSMS Shortcut")
    Set objFile     = objFSO.GetFile(strPath)
    strPathNew      = strAllUserProf & "\" & objFile.Name
    strDebugMsg2    = "Target path: " & strPathNew
    objFile.Copy strPathNew
    strPathNew      = strAllUserDTop & "\" & objFile.Name
    strDebugMsg2    = "Target path: " & strPathNew
    objFile.Copy strPathNew
  End If

  Call DebugLog("BOL Shortcut")
  strPath           = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLDocs & "\" & strMenuBOL & ".lnk"
  strDebugMsg1      = "Source path: " & strPath
  If objFSO.FileExists(strPath) Then
    Call DebugLog("BOL Shortcut")
    Set objFile     = objFSO.GetFile(strPath)
    strPathNew      = strAllUserProf & "\" & objFile.Name
    strDebugMsg2    = "Target path: " & strPathNew
    objFile.Copy strPathNew
    strPathNew      = strAllUserDTop & "\" & objFile.Name
    strDebugMsg2    = "Target path: " & strPathNew
    objFile.Copy strPathNew
  End If

  Call DebugLog("SQLDIAG Shortcut")
  strPath           = strDirProg & "\" & strSQLVersionNum & "\Tools\binn\sqldiag.exe"
  strPathNew        = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuPerfTools 
  Call SetupFolder(strPathNew)
  strDebugMsg1      = "Source: " & strPath
  strPathNew        = strPathNew & "\SQLDiag.lnk"
  Set objShortcut   = objShell.CreateShortcut(strPathNew)
  objShortcut.TargetPath       = "CMD"
  objShortcut.WorkingDirectory = strDirProg & "\" & strSQLVersionNum & "\Tools\binn"
  objShortcut.Arguments        = "/K """ & strPath & """ /?"
  objShortcut.IconLocation     = """" & strPath & ", 0"""
  objShortcut.WindowStyle      = 1
  objShortcut.Save()

  Call SetBuildfileValue("SetupMenusStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLRSMenus()
  Call SetProcessId("5FB", "Reporting Services Menus")
  Dim strIEPath, strIEProg, strRSInstance, strRSMenuPath, strShortcutName, strTargetPath

  strIEPath         = strDirProgSys & "\Internet Explorer"
  strIEProg         = strIEPath & "\iexplore.exe"
  Select Case True
    Case strInstance <> "MSSQLServer"
      strRSInstance = " (" & strInstance & ")"
    Case Else
      strRSInstance = ""
  End Select

  Select Case True
    Case strTCPPortRS = "80"
      strRSHost     = strServer
    Case Else
      strRSHost     = strServer & ":" & strTCPPortRS
  End Select

  strRSMenuPath     = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSQLRS
  Call SetupFolder( strRSMenuPath)

  Call DebugLog("Report Builder Shortcut")
  Select Case True
    Case strSQLVersion = "SQL2005"
      strShortcutName   = "Report Builder" & strRSInstance 
      strTargetPath     = strHTTP & "://" & strRSHost & "/" & strInstRSURL & "/ReportBuilder/ReportBuilder.application"
    Case strSQLVersion = "SQL2008"
      If strSPLevel >= "SP1" Or strSPCULevel >= "CU1" Then   
        strShortcutName = "Report Builder V2" & strRSInstance
        strTargetPath   = strHTTP & "://" & strRSHost & "/" & strInstRSURL & "/ReportBuilder/ReportBuilder_2_0_0_0.application"
      Else
        strShortcutName = "Report Builder" & strRSInstance 
        strTargetPath   = strHTTP & "://" & strRSHost & "/" & strInstRSURL & "/ReportBuilder/ReportBuilder.application"
      End If
    Case Else
      strShortcutName   = "Report Builder V3" & strRSInstance 
      strTargetPath     = strHTTP & "://" & strRSHost & "/" & strInstRSURL & "/ReportBuilder/ReportBuilder_3_0_0_0.application"
  End Select  

  Select Case True
    Case objFSO.FileExists(strIEProg)  
      Set objShortcut              = objShell.CreateShortcut(strRSMenuPath & "\" & strShortcutName & ".lnk")
      objShortcut.TargetPath       = strIEProg
      objShortcut.WorkingDirectory = strIEPath
      objShortcut.Arguments        = strTargetPath
      objShortcut.IconLocation     = """" & strIEPath & ", 1"""
      objShortcut.WindowStyle      = 1
      objShortcut.Save()
    Case Else
      Set objShortcut        = objShell.CreateShortcut(strRSMenuPath & "\" & strShortcutName & ".url")
      objShortcut.TargetPath = strTargetPath
      objShortcut.Save()
  End Select

  Call DebugLog("Report Manager Shortcut")
  strShortcutName   = "Report Manager" & strRSInstance 
  strTargetPath     = strHTTP & "://" & strRSHost & "/Reports"
  Select Case True
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case Else
      strTargetPath = strTargetPath & "_" & strInstance
  End Select
  Select Case True
    Case objFSO.FileExists(strIEProg)  
      Set objShortcut              = objShell.CreateShortcut(strRSMenuPath & "\" & strShortcutName & ".lnk")
      objShortcut.TargetPath       = strIEProg
      objShortcut.WorkingDirectory = strIEPath
      objShortcut.Arguments        = strTargetPath
      objShortcut.IconLocation     = """" & strIEPath & ", 1"""
      objShortcut.WindowStyle      = 1
      objShortcut.Save()
    Case Else
      Set objShortcut        = objShell.CreateShortcut(strRSMenuPath & "\" & strShortcutName & ".url")
      objShortcut.TargetPath = strTargetPath
      objShortcut.Save()
  End Select

  Call SetBuildfileValue("SetupMenusStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWindowsMenus()
  Call SetProcessId("5FC", "Windows menu items")
  Dim strProgram

  Select Case true
    Case strOSVersion < "6.2"
      ' Nothing
    Case strOSVersion > "6.3"
      ' Nothing
    Case Else
      Call DebugLog("Start Menu Shortcut")
      strPath       = strDirSysData & "\Microsoft\Windows\Start Menu"
      strPathNew    = strAllUserDTop & "\Start Menu.lnk"
      strDebugMsg2  = "Target path: " & strPathNew
      Set objShortcut = objShell.CreateShortcut(strPathNew)
      objShortcut.TargetPath  = strPath
      objShortcut.Save()
  End Select

  If strOSVersion <= "6.2" Then
    Call SetMenuItem("Command Prompt", strDirSys & "\system32\cmd.exe", strAllUserProf)
    Call SetMenuItem("Command Prompt", strDirSys & "\system32\cmd.exe", strAllUserDTop)
  End If

  Call SetMenuItem("Notepad", strDirSys & "\system32\notepad.exe", strAllUserProf)
  Call SetMenuItem("Notepad", strDirSys & "\system32\notepad.exe", strAllUserDTop)

  Select Case True
    Case strOSVersion < "6.2"
      Call SetMenuItem("Windows Explorer", strDirSys & "\explorer.exe", strAllUserDTop)
    Case strOSVersion > "6.3"
      ' Nothing
    Case Else
      Call SetMenuItem("File Explorer", strDirSys & "\explorer.exe", strAllUserDTop)
  End Select

  Call SetMenuItem("Registry Editor", strDirSys & "\system32\regedt32.exe", strAllUserProf)

  Call SetMenuItem("Windows Management Instrumentation Tester", strDirSys & "\system32\wbem\wbemtest.exe", strAllUserProf)

  Call DebugLog("Shutdown Shortcut")
  strPath           = strDirSys & "\system32\shutdown.exe"
  strPathNew        = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuAdminTools & "\Remote Shutdown.lnk"
  strDebugMsg1      = "Source path: " & strPath
  strDebugMsg2      = "Target path: " & strPathNew
  Set objShortcut   = objShell.CreateShortcut(strPathNew)
  objShortcut.Arguments   = "-i"
  objShortcut.TargetPath  = strPath
  objShortcut.WindowStyle = 1 ' Normal window
  objShortcut.Save()

  Call SetBuildfileValue("SetupMenusStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub MergeWindowsMenus()
  Call SetProcessId("5FD", "Merge Windows menus")
  Dim arrSQL
  Dim intUBound, intIdx
  Dim strMergeVersion, strMergeMenu, strMergeMenuFlag

  arrSQL            = Split(strSQLList, " ", -1)
  intUBound         = UBound(arrSQL)
  For intIdx = 0 To intUBound
    strMergeVersion = arrSQL(intIdx)
    strMergeMenu    = GetBuildfileValue("Menu" & strMergeVersion)
    strMergeMenuFlag  = GetBuildfileValue("Menu" & strMergeVersion & "Flag")
    Select Case True
      Case strSQLVersion = strMergeVersion
        ' Nothing
      Case strMergeMenu = ""
        ' Nothing
      Case strMergeMenuFlag <> "N"
        ' Nothing
      Case Else
        Call ProcessMenuMerge(strMergeVersion, strMergeMenu)
    End Select
  Next

  Call SetBuildfileValue("SetupMenusStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ProcessMenuMerge(strMergeVersion, strMergeMenu)
  Call DebugLog("ProcessMenuMerge: " & strMergeVersion)
  Dim colFiles, colFolders

  strPath           = strAllUserProf & "\" & strMenuPrograms & "\" & strMergeMenu
  If objFSO.FolderExists(strPath) Then
    Call DebugLog("Merging " & strMergeVersion & " Menu to " & strMenuSQL)
    strDebugMsg1  = "Source: " & strPath
    Set objFolder = objFSO.GetFolder(strPath)
    Set colFiles  = objFolder.Files
    strPathNew    = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL
    strDebugMsg2  = "Target: " & strPathNew
    Select Case True
      Case colFiles Is Nothing
        ' Nothing
      Case IsNull(colFiles)
        ' Nothing
      Case Else
        For Each objFile In colFiles
          objFile.Copy strPathNew & "\" & objFile.Name, True
        Next 
    End Select
    Set colFolders = objFolder.SubFolders
    Select Case True
      Case colFolders Is Nothing
        ' Nothing
      Case IsNull(colFolders)
        ' Nothing
      Case Else
        For Each objSubFolder In colFolders
          objSubFolder.Copy strPathNew & "\" & objSubFolder.Name, True
        Next
    End Select
    objFolder.Delete(True)
  End If

End Sub


Sub SetMenuItem(strName, strSource, strTarget)
  Call DebugLog("SetMenuItem: " & strName & ": " & strSource & " , " & strTarget)

  Call SetupFolder(strTarget)
  strDebugMsg1      = "Source: " & strSource

  Select Case True
    Case Not objFSO.FileExists(strSource)
      Exit Sub
    Case Not objFSO.FileExists(strTarget & "\" & strName & ".lnk") 
      Set objShortcut = objShell.CreateShortcut(strTarget & "\" & strName & ".lnk")
      objShortcut.TargetPath       = strSource
      objShortcut.WorkingDirectory = Left(strSource, InstrRev(strSource, "\") - 1)
      objShortcut.Save()
  End Select

End Sub


Sub ConfigTidy()
  Call SetProcessId("5G", "Configuration Tidy-Up")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5GA"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case Else
      Call SetClusterOnline()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5GB"
      ' Nothing
    Case GetBuildfileValue("SetupAPCluster") <> "YES"
      ' Nothing
    Case Else
      Call ConfigAPCluster()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "5GC"
      ' Nothing
    Case GetBuildfileValue("SetupOldAccounts") <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildfileValue("SetupOldAccountsStatus", strStatusPreConfig)
    Case Else
      Call ConfigOldAccounts()
  End Select

  Call SetProcessId("5GZ", " Configuration Tidy-Up" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetClusterOnline()
  Call SetProcessId("5GA", "Set Cluster resources online")

  If strSetupSQLASCluster = "YES" Then
    Call SetResourceOn(strClusterGroupAS, "GROUP")
  End If

  If strSetupSQLDBCluster = "YES" Then
    Call SetResourceOn(strClusterGroupSQL, "GROUP")
  End If

  If CheckStatus("SQLRSCluster") Then
    Call SetResourceOn(strClusterGroupRS, "GROUP")
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigAPCluster()
  Call SetProcessId("5GB", "Setup AP Cluster Dependencies")

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroupAS & """ /PROP FailoverThreshold=2 FailoverPeriod=1"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      If strInstDTCClusterRes <> "" Then
        strCmd      = "SC CONFIG """ & strInstAS & """ DEPEND= ""MSDTC$" & strInstDTCClusterRes & """"
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
      End If
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ GROUP """ & strClusterGroupSQL & """ /PROP FailoverThreshold=2 FailoverPeriod=1"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      If strInstDTCClusterRes <> "" Then
        strCmd      = "SC CONFIG """ & strInstSQL & """ DEPEND= ""MSDTC$" & strInstDTCClusterRes & """"
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
      End If
  End Select

  Select Case True
    Case strDTCClusterRes = ""
      ' Nothing
    Case strInstDTCClusterRes <> strDTCClusterRes
      strCmd        = "SC CONFIG ""MSDTC$" & strInstDTCClusterRes & """ DEPEND= ""RPCSS""/""SamSS""/MSDTC$" & strDTCClusterRes & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  Call SetBuildfileValue("SetupAPClusterStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ConfigOldAccounts()
  Call SetProcessId("5GC", "Remove Redundant SQL Server Accounts")

  Select Case True
    Case strType = "WORKSTATION"
      ' Nothing
    Case strGroupDBA = strLocalAdmin
      ' Nothing
    Case strGroupDBA <> strGroupDBAAlt
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """DROP LOGIN [" & strLocalAdmin & "];""", 1)
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strMenuSQL2005Flag = "Y"
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """DROP LOGIN [" & strLocalDomain & "\SQLServer2005MSFTEUser$" & strServer & "$" & strInstance & "];""", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """DROP LOGIN [" & strLocalDomain & "\SQLServer2005MSSQLUser$" & strServer & "$" & strInstance & "];""", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """DROP LOGIN [" & strLocalDomain & "\SQLServer2005SQLAgentUser$" & strServer & "$" & strInstance & "];""", 1)
  End Select

  Call SetBuildfileValue("SetupOldAccountsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAutoConfig()
  Call SetProcessId("5H", "AutoConfig")
  Dim strPathAutoConfig

  Call SetBuildfileValue("SetupAutoConfigStatus", strStatusProgress)

  strPathAutoConfig = FormatFolder(GetBuildfileValue("PathAutoConfig"))
  Select Case True
    Case strInstance = "MSSQLSERVER"
      Call ProcessAutoConfig(strPathAutoConfig & strServer, "Y")
    Case Else
      Call ProcessAutoConfig(strPathAutoConfig & strServer & "$" & StrInstance, "Y")
  End Select

  Select Case True
    Case GetBuildfileValue("ActionAO") = ""
      ' Nothing
    Case Else
      Call ProcessAutoConfig(strPathAutoConfig & GetBuildfileValue("AGName"), "")
  End Select

  Select Case True
    Case GetBuildfileValue("ActionDAG") = ""
      ' Nothing
    Case Else
      Call ProcessAutoConfig(strPathAutoConfig & GetBuildfileValue("AGDagName"), "")
  End Select

  If strClusterName <> "" Then
    Call ProcessAutoConfig(strPathAutoConfig & strClusterName, "")
  End If

  Call ProcessAutoConfig(strPathAutoConfig, "")

  Call SetBuildfileValue("SetupAutoConfigStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub ProcessAutoConfig(strFolder, strAlways)
  Call DebugLog("ProcessAutoConfig: " & strFolder)
  Const adVarChar   = 200
  Dim colFiles
  Dim intIdx
  Dim objFile, objFolder, objInstParm
  Dim rsFiles
  Dim strAutoConfigList, strFile, strFileType, strLabel, strSetup

  If Not objFSO.FolderExists(strFolder) Then
    Exit Sub
  End If

  Call DebugLog("Define File List")
  Set rsFiles       = CreateObject("ADODB.Recordset")
  rsFiles.Fields.Append "Name",  adVarChar, 255
  rsFiles.Fields.Append "Label", adVarChar, 255
  rsFiles.Fields.Append "Setup", adVarChar, 20
  rsFiles.Open

  Call DebugLog("Build File List")
  Set objFolder     = objFSO.GetFolder(strFolder)
  Set colFiles      = objFolder.Files
  For Each objFile In colFiles
    strFile         = objFile.Name
    strDebugMsg1    = "File: " & strFile
    rsFiles.AddNew
    rsFiles("Name") = strFile
    Select Case True
      Case Instr(strFile, ".") = 0
        strLabel         = Replace(strFile, " ", "-")
        rsFiles("Label") = strLabel
        rsFiles("Setup") = strStatusBypassed
      Case UCase(Right(strFile, 4)) = UCase(".txt")
        strLabel         = Replace(Left(strFile, Instr(strFile, ".") - 1), " ", "-")
        rsFiles("Label") = strLabel
        rsFiles("Setup") = strStatusBypassed
      Case Else
        strLabel         = Replace(Left(strFile, Instr(strFile, ".") - 1), " ", "-")
        rsFiles("Label") = strLabel
        rsFiles("Setup") = GetBuildfileValue("SetupAutoConfig" & UCase(strLabel) & "Status")
    End Select
  Next

  Call DebugLog("Process File List")
  rsFiles.Sort = "Name ASC"
  Do Until rsFiles.EOF 
    strFile         = rsFiles("Name")
    strFile         = Mid(strFile, 2, Len(strFile) - 2)
    strFileType     = ""
    strLabel        = rsFiles("Label")
    strSetup        = rsFiles("Setup")
    intIdx          = InStrRev(strFile, ".")
    If intIdx > 0 Then
      strFileType   = Right(strFile, intIdx + 1)
    End If
    Select Case True
      Case(strActionSQLDB = "ADDNODE") & (UCase(strFileType) = "SQL") & (strAlways <> "Y")
        ' Nothing
      Case (strSetup = "") Or (strSetup = strStatusProgress)
        strAutoConfigList = GetBuildfileValue("AutoConfigList")
        If Instr(" " & strAutoConfigList, " " & strLabel & " ") = 0 Then
          Call SetBuildfileValue("AutoConfigList", strAutoConfigList & strLabel & " ")
        End If
        Call SetXMLParm(objInstParm, "PathMain",     strFolder)
        Call SetXMLParm(objInstParm, "LogClean",     GetLogClean(strFile))
        Call SetXMLParm(objInstParm, "LogXtra",      strLabel)
        Call SetXMLParm(objInstParm, "StatusOption", strStatusComplete)
        Call RunInstall("AutoConfig", strFile, objInstParm)
      Case Else
        ' Nothing
    End Select
    rsFiles.MoveNext
  Loop

  rsFiles.Close
  Set rsFiles       = Nothing

End Sub

Function GetLogClean(strFile)

  intIdx            = InStrRev(strFile, ".")
  Select Case True
    Case intIdx = 0
      GetLogClean   = ""
    Case Instr(" BAT CMD SQL PS1 ", " " & UCase(Mid(strFile, intIdx + 1)) & " ") > 0
      GetLogClean   = "Y"
    Case Else
      GetLogClean   = ""
  End Select

End Function


Sub UserConfiguration()
  Call SetProcessId("5U", "User Configuration Tasks")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",    GetBuildfileValue("FBParm"))
  Call RunInstall("UserPreparation", GetBuildfileValue("UserConfigurationvbs"), objInstParm)

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
