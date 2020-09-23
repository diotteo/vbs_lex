'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild2InstallSQL.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      SQL Server Install 
'
'  Author:       Ed Vassie
'
'  Date:         30 Jun 2008
'
'  Change History
'  Version  Author        Date         Description
'  2.2      Ed Vassie     18 Jun 2010  Initial version for SQL Server 2008 R2
'  2.1      Ed Vassie     17 Nov 2009  Upgraded to support Cluster install
'  2.0      Ed Vassie     30 Jun 2008  Initial SQL Server 2008 version for FineBuild v2.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colFolders, colPrcEnvVars, colUsrEnvVars, colVol
Dim intIndex
Dim objCmd, objConfig, objDrive, objFile, objFolder, objFSO, objInParam, objOCFile, objOutParam, objShell, objVol, objWMI, objWMIReg
Dim strAction, strActionSQLDB, strActionSQLAS, strActionSQLIS, strActionSQLRS, strActionSQLTools, strActionClusInst, strAllowUpgradeForRSSharePointMode, strAnyKey, strASProviderMSOlap, strAsServerMode, strAsSvcStartuptype, strAGTSvcStartuptype
Dim strSqlBrowserStartup, strCltStartupType, strCtlrStartupType, strIsSvcStartuptype, strSQLSvcStartuptype, strRole
Dim strCatalogInstance, strCatalogServer, strCatalogServerName, strCLSIdSQL, strCLSIdSQLSetup, strCLSIdVS, strCmd, strCmdPS, strCmdRS, strCmdSQL, strCompatFlags
Dim strClusStorage, strClusSubnet, strClusterAction, strClusterHost, strClusterName, strClusterNameAS, strClusterNetworkAS, strClusterNameIS, strClusterNamePE, strClusterNamePM, strClusterNameRS, strClusterGroupAS, strClusterGroupFS, strClusterGroupRS, strClusterGroupSQL, strClusterNameSQL, strClusterNetworkSQL, strClusterIPV4RS, strClusterIPV4SQL, strClusterIPV6RS, strClusterIPV6SQL, strClusterOptions, strClusterReport, strAgtDomaingroup, strASDomainGroup, strFTSDomainGroup, strSQLAdminAccounts, strSQLDomainGroup, strSSASAdminAccounts, strClusterNameDTC, strClusterNetworkDTC, strClusterGroupDTC, strClusIPAddress, strClusIPV4Network, strClusIPV6Network, strClusIPVersion, strFailoverClusterDisks, strFailoverClusterGroup, strFailoverClusterIPAddresses, strFailoverClusterNetworkName, strFailoverClusterRollOwnership, strLabDTC
Dim strDomain, strDTCClusterRes, strDTCMultiInstance
Dim strEdition, strEdType, strEnu, strErrorReporting, strFeatureList, strFirewallStatus, strFeatures, strFileArc, strFilePerm, strFSInstLevel, strFSLevel, strFSShareName, strFolderName, strHKLMFB, strHKLMSQL, strHTTP
Dim strIAcceptLicenseTerms, strInstance, strInstDTCClusterRes, strInstAgent, strInstAnal, strInstLog, strInstNode, strInstPE, strInstPM, strInstSQL, strInstAS, strInstASSQL, strInstIS, strInstRegAS, strInstRegSQL, strInstRS, strInstRSDir, strInstRSSQL, strIsInstallDBA, strISMasterPort, strISMasterStartupType, strISMasterThumbprint, strISWorkerCert, strISWorkerMaster, strISWorkerStartupType
Dim strManagementServer, strManagementServerRes, strManagementServerName, strManagementInstance, strMenuSQL, strMenuSQL2005Flag, strMode
Dim strNativeOS, strNTAuthAccount, strNTAuthOSName, strNTService
Dim strOCFeature, strOCFile, strOSLevel, strOSName, strOSType, strOSVersion
Dim strPath, strPathAddComp, strPathAlt, strPathFB, strPathFBScripts, strPathLog, strPathNew, strPathNLS, strPathOld, strPathSSMS, strPathSys, strPathTemp, strPathVS
Dim strPBEngSvcAccount, strPBEngSvcPassword, strPBEngSvcStartup, strPBDMSSvcAccount, strPBDMSSvcPassword, strPBDMSSvcStartup, strPBPortRange, strPBScaleout, strPID, strProcArc, strProgCacls
Dim strResSuffixAS, strResSuffixDB, strRsFxVersion, strRSDBName, strRSActualMode, strRSInstallMode, strRSShpInstallMode, strRSSQLLocal, strRSSvcStartupType, strRSVersionNum
Dim strReboot, strRegSSIS, strRegSSISSetup
Dim strServInst, strServName, strSetupAnalytics, strSetupJRE, strSetupPolyBase, strSetupPolyBaseCluster, strSetupPowerBI, strSetupPS2, strSetupPython, strSetupRServer, strSetupSSMS, strSetupDTCCluster, strSetupDTCClusterStatus, strSetupDTCNetAccess, strSetupDTCNetAccessStatus, strSetupSSISCluster
Dim strSetupSlipstream, strSetupStreamInsight, strSecurityMode, strServer, strSetupFlag
Dim strSQLBinRoot, strSQLExe, strSQLJavaDir, strSQLLanguage, strSQLLog, strSQLMedia, strSQLMediaBase, strSQLMediaOrig, strSQLVersion, strSQLVersionFull, strSQLVersionNum
Dim strPCUSource, strCUSource, strPathSQLSP, strPathSQLSPOrig, strSQLSharedMR, strSQLMediaArc, strSQLMediaPCU, strSQLSupportMsi, strSQMReporting, strSPLevel, strSPCULevel, strStopAt, strTCPPortDTC, strtempdbFile, strtempdbLogFile, strType
Dim strSetupBOL, strSetupBIDS, strSetupBPE, strSetupDTCCID, strSetupIIS, strSetupISMaster, strSetupISWorker, strSetupMDS, strSetupNet3, strSetupNet4, strSetupShares, strSetupSQLAS, strSetupSQLASCluster, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLRS, strSetupSQLRSCluster, strSetupSQLDBRepl, strSetupSQLDBFS, strSetupSQLDBFT, strSetupSQLIS, strSetupSQLNS, strSetupSQLTools, strSetupSP, strSetupSPCU, strSetupSQLBC
Dim strSKUUpgrade, strStatusAssumed, strStatusKB933789
Dim strVersionInst, strVersionNet2, strVersionNet3, strVersionNet4, strVersionPS
Dim strDirBackup, strDirBackupAS, strDirBPE, strDirData, strDirDataAS, strDirDataFS, strDirDataFT, strDirDataIS, strDirDRU, strDirSysDB, strDirLog, strDirLogAS, strDirLogTemp, strDirProg, strDirProgX86, strDirTemp, strDirTempAS, strVolProg, strDirSys, strWaitLong, strWaitShort, strWinDir, strWOWX86
Dim strCollationAS, strCollationSQL, strSetupStdAccounts, strGroupDBA, strGroupDBANonSA, strGroupDistComUsers, strGroupMSA, strGroupUsers, strListDir, strListFirstOpts, strListOpts, strLocalAdmin, strUserFeatures, strUserOptions
Dim strSetupDQ, strSetupDQC, strSetupDRUCtlr, strSetupDRUClt, strCltSvcAccount, strCltSvcPassword, strCtlrSvcAccount, strCtlrSvcPassword
Dim strExpVersion, strExpressOptions, strExtSvcAccount, strExtSvcPassword, strEnableRANU, strUpgradeOptions, strFTUpgradeOption, strRSDBAccount, strRSDBPassword, strUseSysDB
Dim strSqlAccount,  strAgtAccount,  strAsAccount,  strFarmAccount,  strFTAccount,  strIsAccount,  strIsMasterAccount,  strIsWorkerAccount,  strRsAccount,  strSqlBrowserAccount,  strAsSysadmin, strSqlSysadmin, strFarmAdminIPort, strPassPhrase
Dim strSqlPassword, strAgtPassword, strAsPassword, strFarmPassword, strFTPassword, strIsPassword, strIsMasterPassword, strIsWorkerPassword, strRsPassword, strSqlBrowserPassword, strsaPwd, strAdminPassword
Dim strUpdateSource, strUseFreeSSMS, strUserAccount, strUserDNSDomain, strUserName
Dim strVolBackup, strVolBackupAS, strVolBPE, strVolDBA, strVolData, strVolDataAS, strVolDataFS, strVolDataFT, strVolDTC, strVolLabel, strVolList, strVolLog, strVolLogAS, strVolLogTemp, strVolTemp, strVolTempAS, strVolSys, strVolSysDB

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "2Z"
      ' Nothing
    Case Else
      Call Process_Install()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "2Z"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("SQL Server Install complete", 2, "SQL Server Base Install" ,64)
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
      Call FBLog(" SQL Server Base Install failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation ()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBUtils.vbs"
  Include "FBManageAccount.vbs"
  Include "FBManageBoot.vbs"
  Include "FBManageCluster.vbs"
  Include "FBManageInstall.vbs"
  Include "FBManageService.vbs"
  Call SetProcessIdCode("FB2I")

  Set objConfig     = CreateObject("Microsoft.XMLDOM")  
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate,(Security)}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colPrcEnvVars = objShell.Environment("Process")
  Set colUsrEnvVars = objShell.Environment("User")

  strInstance       = GetBuildfileValue("Instance")
  strHKLMFB         = GetBuildfileValue("HKLMFB")
  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAction         = GetBuildfileValue("Action")
  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strActionSQLIS    = GetBuildfileValue("ActionSQLIS")
  strActionSQLRS    = GetBuildfileValue("ActionSQLRS")
  strActionSQLTools = GetBuildfileValue("ActionSQLTools")
  strActionClusInst = GetBuildfileValue("ActionClusInst")
  strAdminPassword  = GetBuildfileValue("AdminPassword")
  strAgtDomainGroup = GetBuildfileValue("AgtDomainGroup")
  strAgtAccount     = GetBuildfileValue("AgtAccount")
  strAgtPassword    = GetBuildfileValue("AgtPassword")
  strAgtSvcStartupType = GetBuildfileValue("AgtSvcStartupType")
  strAllowUpgradeForRSSharePointMode = GetBuildfileValue("AllowUpgradeForRSSharePointMode")
  strAnykey         = GetBuildfileValue("Anykey")
  strASDomainGroup  = GetBuildfileValue("ASDomainGroup")
  strAsAccount      = GetBuildfileValue("AsAccount")
  strAsPassword     = GetBuildfileValue("AsPassword")
  strAsServerMode   = GetBuildfileValue("ASServerMode")
  strAsSvcStartupType  = GetBuildfileValue("AsSvcStartupType")
  strCatalogInstance   = GetBuildfileValue("CatalogInstance")
  strCatalogServer     = GetBuildfileValue("CatalogServer")
  strCatalogServerName = GetBuildfileValue("CatalogServerName")
  strCLSIdSQL       = GetBuildfileValue("CLSIdSQL")
  strCLSIdSQLSetup  = GetBuildfileValue("CLSIdSQLSetup")
  strCLSIdVS        = GetBuildfileValue("CLSIdVS")
  strClusStorage    = GetBuildfileValue("ClusStorage")
  strClusSubnet     = GetBuildfileValue("ClusSubnet")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterNameAS  = GetBuildfileValue("ClusterNameAS")
  strClusterNameDTC = GetBuildfileValue("ClusterNameDTC")
  strClusterNetworkAS  = GetBuildfileValue("ClusterNetworkAS")
  strClusterNetworkDTC = GetBuildfileValue("ClusterNetworkDTC")
  strClusterNetworkSQL = GetBuildfileValue("ClusterNetworkSQL")
  strClusterNameIS  = GetBuildfileValue("ClusterNameIS")
  strClusterNamePE  = GetBuildfileValue("ClusterNamePE")
  strClusterNamePM  = GetBuildfileValue("ClusterNamePM")
  strClusterNameRS  = GetBuildfileValue("ClusterNameRS")
  strClusterGroupAS = GetBuildfileValue("ClusterGroupAS")
  strClusterGroupFS = GetBuildfileValue("ClusterGroupFS")
  strClusterGroupRS = GetBuildfileValue("ClusterGroupRS")
  strClusterGroupDTC  = GetBuildfileValue("ClusterGroupDTC")
  strClusterGroupSQL  = GetBuildfileValue("ClusterGroupSQL")
  strClusterNameSQL = GetBuildfileValue("ClusterNameSQL")
  strClusterReport  = GetBuildfileValue("ClusterReport")
  strClusIPAddress  = GetBuildfileValue("ClusIPAddress")
  strClusIPV4Network  = GetBuildfileValue("ClusIPV4Network")
  strClusIPV6Network  = GetBuildfileValue("ClusIPV6Network")
  strClusIPVersion  = GetBuildfileValue("ClusIPVersion")
  strCmdPS          = GetBuildfileValue("CmdPS")
  strCmdRS          = GetBuildfileValue("CmdRS")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strCompatFlags    = GetBuildfileValue("CompatFlags")
  strCtlrSvcAccount   = GetBuildfileValue("CtlrSvcAccount")
  strCtlrSvcPassword  = GetBuildfileValue("CtlrSvcPassword")
  strCtlrStartuptype  = GetBuildfileValue("CtlrStartupType")
  strCltSvcAccount  = GetBuildfileValue("CltSvcAccount")
  strCltSvcPassword = GetBuildfileValue("CltSvcPassword")
  strCltStartupType = GetBuildfileValue("CltStartupType")
  strDirBPE         = GetBuildfileValue("DirBPE")
  strDirProg        = GetBuildfileValue("DirProg")
  strDirProgX86     = GetBuildfileValue("DirProgX86")
  strDirBackupAS    = GetBuildfileValue("DirBackupAS")
  strDirDataAS      = GetBuildfileValue("DirDataAS")
  strDirLogAS       = GetBuildfileValue("DirLogAS")
  strDirTempAS      = GetBuildfileValue("DirTempAS")
  strDirBackup      = GetBuildfileValue("DirBackup")
  strDirData        = GetBuildfileValue("DirData")
  strDirDataFS      = GetBuildfileValue("DirDataFS")
  strDirDataFT      = GetBuildfileValue("DirDataFT")
  strDirDataIS      = GetBuildfileValue("DirDataIS")
  strDirDRU         = GetBuildfileValue("DirDRU")
  strDirLog         = GetBuildfileValue("DirLog")
  strDirLogTemp     = GetBuildfileValue("DirLogTemp")
  strDirTemp        = GetBuildfileValue("DirTemp")
  strDirSys         = GetBuildfileValue("DirSys")
  strDirSysDB       = GetBuildfileValue("DirSysDB")
  strDomain         = GetBuildfileValue("Domain")
  strDTCClusterRes  = GetBuildfileValue("DTCClusterRes")
  strDTCMultiInstance = GetBuildfileValue("DTCMultiInstance")
  strEdition        = GetBuildfileValue("AuditEdition")
  strEdType         = GetBuildfileValue("EdType")
  strEnableRANU     = GetBuildfileValue("EnableRANU")
  strEnu            = GetBuildfileValue("Enu")
  strErrorReporting = GetBuildfileValue("ErrorReporting")
  strExpVersion     = GetBuildfileValue("ExpVersion")
  strExtSvcAccount  = GetBuildfileValue("ExtSvcAccount")
  strExtSvcPassword = GetBuildfileValue("ExtSvcPassword")
  strFailoverClusterRollOwnership = GetBuildfileValue("FailoverClusterRollOwnership")
  strFarmAccount    = GetBuildfileValue("FarmAccount")
  strFarmPassword   = GetBuildfileValue("FarmPassword")
  strFarmAdminIPort = GetBuildfileValue("FarmAdminIPort")
  strFeatureList    = ""
  strFileArc        = GetBuildfileValue("FileArc")
  strFilePerm       = GetBuildfileValue("FilePerm")
  strFirewallStatus = GetBuildfileValue("FirewallStatus")
  strFSInstLevel    = GetBuildfileValue("FSInstLevel")
  strFSLevel        = GetBuildfileValue("FSLevel")
  strFSShareName    = GetBuildfileValue("FSShareName")
  strFTSDomainGroup = GetBuildfileValue("FTSDomainGroup")
  strFTAccount      = GetBuildfileValue("FtAccount")
  strFTPassword     = GetBuildfileValue("FtPassword")
  strFTUpgradeOption  = GetBuildfileValue("FTUpgradeOption")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strGroupDistComUsers   = GetBuildfileValue("GroupDistComUsers")
  strGroupMSA       = GetBuildfileValue("GroupMSA")
  strGroupUsers     = GetBuildfileValue("GroupUsers")
  strHTTP           = GetBuildfileValue("HTTP")
  strIAcceptLicenseTerms = GetBuildfileValue("IAcceptLicenseTerms")
  strInstDTCClusterRes   = GetBuildfileValue("InstDTCClusterRes")
  strInstAnal       = GetBuildfileValue("InstAnal")
  strInstLog        = GetBuildfileValue("InstLog")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstAgent      = GetBuildfileValue("InstAgent")
  strInstAS         = GetBuildfileValue("InstAS")
  strInstASSQL      = GetBuildfileValue("InstASSQL")
  strInstIS         = GetBuildfileValue("InstIS")
  strInstPE         = GetBuildfileValue("InstPE")
  strInstPM         = GetBuildfileValue("InstPM")
  strInstRegAS      = GetBuildfileValue("InstRegAS")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstRSDir      = GetBuildfileValue("InstRSDir")
  strInstRSSQL      = GetBuildfileValue("InstRSSQL")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strIsInstallDBA   = GetBuildfileValue("IsInstallDBA")
  strIsAccount      = GetBuildfileValue("IsAccount")
  strIsPassword     = GetBuildfileValue("IsPassword")
  strIsSvcStartupType = GetBuildfileValue("IsSvcStartupType")
  strIsMasterAccount  = GetBuildfileValue("IsMasterAccount")
  strIsMasterPassword = GetBuildfileValue("IsMasterPassword")
  strIsMasterPort     = GetBuildfileValue("IsMasterPort")
  strIsMasterStartupType = GetBuildfileValue("IsMasterStartupType")
  strIsMasterThumbprint  = GetBuildfileValue("IsMasterThumbprint")
  strIsWorkerAccount  = GetBuildfileValue("IsWorkerAccount")
  strIsWorkerPassword = GetBuildfileValue("IsWorkerPassword")
  strIsWorkerCert   = GetBuildfileValue("IsWorkerCert")
  strIsWorkerMaster = GetBuildfileValue("IsWorkerMaster")
  strIsWorkerStartupType = GetBuildfileValue("IsWorkerStartupType")
  strLabDTC         = GetBuildfileValue("LabDTC")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strManagementServer     = GetBuildfileValue("ManagementServer")
  strManagementServerName = GetBuildfileValue("ManagementServerName")
  strManagementInstance   = GetBuildfileValue("ManagementInstance")
  strManagementServerRes  = GetBuildfileValue("ManagementServerRes")
  strMenuSQL          = GetBuildfileValue("MenuSQL")
  strMenuSQL2005Flag  = GetBuildfileValue("MenuSQL2005Flag")
  strMode           = GetBuildfileValue("Mode")
  strNativeOS       = GetBuildfileValue("NativeOS")
  strNTAuthAccount  = GetBuildfileValue("NTAuthAccount")
  strNTAuthOSName   = GetBuildfileValue("NTAuthOSName")
  strNTService      = GetBuildfileValue("NTService")
  strOSLevel        = GetBuildfileValue("OSLevel")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPassphrase     = GetBuildfileValue("Passphrase")
  strPathSSMS       = GetBuildfileValue("PathSSMS")
  strPathSys        = GetBuildfileValue("PathSys")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPCUSource      = GetBuildFileValue("PCUSource")
  strCUSource       = GetBuildFileValue("CUSource")
  strPathAddComp    = FormatFolder("PathAddComp")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  strPathSQLSP      = FormatFolder("PathSQLSP")
  strPathSQLSPOrig  = GetBuildfileValue("PathSQLSPOrig")
  strPathVS         = GetBuildfileValue("PathVS")
  strPBEngSvcAccount  = GetBuildfileValue("PBEngSvcAccount")
  strPBEngSvcPassword = GetBuildfileValue("PBEngSvcPassword")
  strPBEngSvcStartup  = GetBuildfileValue("PBEngSvcStartup")
  strPBDMSSvcAccount  = GetBuildfileValue("PBDMSSvcAccount")
  strPBDMSSvcPassword = GetBuildfileValue("PBDMSSvcPassword")
  strPBDMSSvcStartup  = GetBuildfileValue("PBDMSSvcStartup")
  strPBPortrange    = GetBuildfileValue("PBPortRange")
  strPBScaleout     = GetBuildfileValue("PBScaleout")
  strPID            = GetBuildfileValue("PID")
  strProcArc        = GetBuildfileValue("ProcArc")
  strSetupJRE       = GetBuildfileValue("SetupJRE")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strReboot         = GetBuildfileValue("RebootStatus")
  strRegSSIS        = GetBuildfileValue("RegSSIS")
  strRegSSISSetup   = GetBuildfileValue("RegSSISSetup")
  strResSuffixAS    = GetBuildfileValue("ResSuffixAS")
  strResSuffixDB    = GetBuildfileValue("ResSuffixDB")
  strRole           = GetBuildfileValue("Role")
  strRsFxVersion    = GetBuildfileValue("RsFxVersion")
  strRsAccount      = GetBuildfileValue("RsAccount")
  strRsPassword     = GetBuildfileValue("RsPassword")
  strRsSvcStartupType = GetBuildfileValue("RsSvcStartupType")
  strRSDBName       = GetBuildfileValue("RSDBName")
  strRSDBAccount    = GetBuildfileValue("RSUpgradeDatabaseAccount")
  strRSDBPassword   = GetBuildfileValue("RSUpgradePassword")
  strRSInstallMode  = GetBuildfileValue("RSInstallMode")
  strRSShpInstallMode = GetBuildfileValue("RSShpInstallMode")
  strRSVersionNum   = GetBuildfileValue("RSVersionNum")
  strsaPwd          = GetBuildfileValue("saPwd")
  strServName       = GetBuildfileValue("ServName")
  strSetupAnalytics = GetBuildfileValue("SetupAnalytics")
  strSetupBol       = GetBuildfileValue("SetupBOL")
  strSetupBIDS      = GetBuildfileValue("SetupBIDS")
  strSetupBPE       = GetBuildfileValue("SetupBPE")
  strSetupDQ        = GetBuildfileValue("SetupDQ")
  strSetupDQC       = GetBuildfileValue("SetupDQC")
  strSetupDTCCID    = GetBuildfileValue("SetupDTCCID")
  strSetupDRUCtlr   = GetBuildfileValue("SetupDRUCtlr")
  strSetupDRUClt    = GetBuildfileValue("SetupDRUClt")
  strSetupDTCCluster         = GetBuildfileValue("SetupDTCCluster")
  strSetupDTCClusterStatus   = GetBuildfileValue("SetupDTCClusterStatus")
  strSetupDTCNetAccess       = GetBuildfileValue("SetupDTCNetAccess")
  strSetupDTCNetAccessStatus = GetBuildfileValue("SetupDTCNetAccessStatus")
  strSetupIIS       = GetBuildfileValue("SetupIIS")
  strSetupISMaster  = GetBuildfileValue("SetupISMaster")
  strSetupISWorker  = GetBuildfileValue("SetupISWorker")
  strSetupMDS       = GetBuildfileValue("SetupMDS")
  strSetupNet3      = GetBuildfileValue("SetupNet3")
  strSetupNet4      = GetBuildfileValue("SetupNet4")
  strSetupPolyBase  = GetBuildfileValue("SetupPolyBase")
  strSetupPolyBaseCluster = GetBuildfileValue("SetupPolyBaseCluster")
  strSetupPowerBI   = GetBuildfileValue("SetupPowerBI")
  strSetupPS2       = GetBuildfileValue("SetupPS2")
  strSetupPython    = GetBuildfileValue("SetupPython")
  strSetupRServer   = GetBuildfileValue("SetupRServer")
  strSetupShares    = GetBuildfileValue("SetupShares") 
  strSetupSlipstream  = GetBuildfileValue("SetupSlipstream")
  strSetupSSISCluster = GetBuildfileValue("SetupSSISCluster")
  strSetupSSMS      = GetBuildfileValue("SetupSSMS")
  strSetupSP        = GetBuildfileValue("SetupSP")    
  strSetupSPCU      = GetBuildfileValue("SetupSPCU")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLASCluster = GetBuildfileValue("SetupSQLASCluster")
  strSetupSQLBC     = GetBuildfileValue("SetupSQLBC")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster  = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLDBRepl = GetBuildfileValue("SetupSQLDBRepl")
  strSetupSQLDBFS   = GetBuildfileValue("SetupSQLDBFS")
  strSetupSQLDBFT   = GetBuildfileValue("SetupSQLDBFT")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLNS     = GetBuildfileValue("SetupSQLNS")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster  = GetBuildfileValue("SetupSQLRSCluster")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSetupStdAccounts   = GetBuildfileValue("SetupStdAccounts")
  strSetupStreamInsight = GetBuildfileValue("SetupStreamInsight")
  strServer         = GetBuildfileValue("AuditServer")
  strServInst       = GetBuildfileValue("ServInst")
  strSKUUpgrade     = GetBuildfileValue("SKUUpgrade")
  strSPLevel        = GetBuildfileValue("SPLevel")
  strSPCULevel      = GetBuildfileValue("SPCULevel")
  strSQLDomainGroup = GetBuildfileValue("SQLDomainGroup")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSqlPassword    = GetBuildfileValue("SqlPassword")
  strSqlAdminAccounts   = GetBuildfileValue("SQLAdminAccounts")
  strSqlSvcStartupType  = GetBuildfileValue("SqlSvcStartupType")
  strSqlBrowserAccount  = GetBuildfileValue("SqlBrowserAccount")
  strSqlBrowserPassword = GetBuildfileValue("SqlBrowserPassword")
  strSqlBrowserStartup  = GetBuildfileValue("SqlBrowserStartup")
  strSQLExe         = GetBuildfileValue("SQLExe")
  strSQLJavaDir     = GetBuildfileValue("SQLJavaDir")
  strSQLLanguage    = GetBuildfileValue("SQLLanguage")
  strSQLMedia       = FormatFolder("PathSQLMedia")
  strSQLMediaBase   = FormatFolder("PathSQLMediaBase")
  strSQLMediaOrig   = GetBuildfileValue("PathSQLMediaOrig")
  strSQLMediaArc    = GetBuildfileValue("SQLMediaArc")
  strSQLSharedMR    = GetBuildfileValue("SQLSharedMR")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strTCPPortDTC     = GetBuildfileValue("TCPPortDTC")
  strSQLSupportMsi  = GetBuildfileValue("SQLSupportMsi")
  strSQMReporting   = GetBuildfileValue("SQMReporting")
  strSSASAdminAccounts  = GetBuildfileValue("SSASAdminAccounts")
  strStatusAssumed  = GetBuildFileValue("StatusAssumed")
  strStopAt         = GetBuildFileValue("StopAt")
  strtempdbFile     = GetBuildfileValue("tempdbFile")
  strtempdbLogFile  = GetBuildfileValue("tempdbLogFile")
  strType           = GetBuildfileValue("Type")
  strUseFreeSSMS    = GetBuildfileValue("UseFreeSSMS")
  strUserAccount    = GetBuildfileValue("UserAccount")
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  strUserFeatures   = GetBuildfileValue("Features")
  strUserName       = GetBuildfileValue("AuditUser")
  strUserOptions    = GetBuildfileValue("Options")
  strUseSysDB       = GetBuildfileValue("UseSysDB")
  strUpdateSource   = FormatFolder("UpdateSource")
  strVersionNet3    = GetBuildfileValue("VersionNet3")
  strVersionNet4    = GetBuildfileValue("VersionNet4")
  strVolBackup      = GetBuildfileValue("VolBackup")
  strVolBackupAS    = GetBuildfileValue("VolBackupAS")
  strVolBPE         = GetBuildfileValue("VolBPE")
  strVolData        = GetBuildfileValue("VolData")
  strVolDataAS      = GetBuildfileValue("VolDataAS")
  strVolDataFS      = GetBuildfileValue("VolDataFS")
  strVolDataFT      = GetBuildfileValue("VolDataFT")
  strVolDBA         = GetBuildfileValue("VolDBA")
  strVolDTC         = GetBuildfileValue("VolDTC")
  strVolLog         = GetBuildfileValue("VolLog")
  strVolLogAS       = GetBuildfileValue("VolLogAS")
  strVolLogTemp     = GetBuildfileValue("VolLogTemp")
  strVolProg        = GetBuildfileValue("VolProg")
  strVolTemp        = GetBuildfileValue("VolTemp")
  strVolTempAS      = GetBuildfileValue("VolTempAS")
  strVolSys         = GetBuildfileValue("VolSys")
  strVolSysDB       = GetBuildfileValue("VolSysDB")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")
  strWinDir         = GetBuildfileValue("WinDir")
  strWOWX86         = GetBuildfileValue("WOWX86")

  strPath           = "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB933789\"
  objWMIReg.GetStringValue strHKLM,strPath,"Description",strStatusKB933789
  strPath           = strPathSys & "msiexec.exe"
  strVersionInst    = objFSO.GetFileVersion(strPath)
  strPath           = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v2.0.50727\"
  objWMIReg.GetStringValue strHKLM,strPath,"Version",strVersionNet2
  strPath           = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\"
  objWMIReg.GetStringValue strHKLM,strPath,"PowerShellVersion",strVersionPS

End Sub


Sub Process_Install()
  Call SetProcessId("2", strSQLVersion & " Install processing (FineBuild2InstallSQL.vbs)")
 
  Call SetUpdate("ON")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId >= "2AZ"
      ' Nothing
    Case Else
      Call SetupPreSQLTasks()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BZ"
      ' Nothing
    Case Else
      Call SetupSQLInstall()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId >= "2CZ"
      ' Nothing
    Case Else
      Call SetupPostSQLTasks()
  End Select

  Call SetUpdate("OFF")
  
  Call SetProcessId("2Z", " SQL Install processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupPreSQLTasks()
  Call SetProcessId("2A", "Pre-requisite processing")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAZ"
      ' Nothing
    Case Else
      Call SetupWindowsPreReqs()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AB"
      ' Nothing
    Case strSetupSlipstream <> "YES"
      ' Nothing
    Case Else
      Call SetupSlipstream()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACZ"
      ' Nothing
    Case Else
      Call SetupMSDTC()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AD"
      ' Nothing
    Case strSetupJRE <> "YES" 
      ' Nothing
    Case Else
      Call SetupJRE()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AE"
      ' Nothing
    Case strSetupNet3 = "YES"
      Call SetBuildFileValue("SetupNet20Status", strStatusBypassed & ", .Net 3.5 will be installed")
    Case strVersionNet2 >= "1"
      ' Nothing
    Case strSQLVersion > "SQL2005"
      ' Nothing
    Case Else
      Call SetupNet20()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AF"
      ' Nothing
    Case strSetupNet3 <> "YES"
      ' Nothing
    Case strVersionNet3 >= "3.5.30729.01"
      Call SetBuildFileValue("SetupNet3Status", strStatusPreConfig)
    Case Else
      Call SetupNet35SP1()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AG"
      ' Nothing
    Case strSetupIIS <> "YES"
      ' Nothing
    Case Else
      Call SetupIIS()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AH"
      ' Nothing
    Case GetBuildfileValue("SetupMSI45") <> "YES"
      ' Nothing
    Case strVersionInst >= "4.5.6001.22159"
      Call SetBuildFileValue("SetupMSI45Status", strStatusPreConfig)
    Case Else
      Call SetupInstaller45()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AI"
      ' Nothing
    Case GetBuildfileValue("SetupPS1") <> "YES"
      ' Nothing
    Case strSetupPS2 = "YES"
      Call SetBuildfileValue("SetupPS1Status", strStatusBypassed & ", PS2 will be installed")
    Case strVersionPS >= "1.0"
      Call SetBuildFileValue("SetupPS1Status", strStatusPreConfig)
    Case Else
      Call SetupPowerShellV1()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AJ"
      ' Nothing
    Case strSetupPS2 <> "YES"
      ' Nothing
    Case strVersionPS >= "2.0"
      Call SetBuildFileValue("SetupPS2Status", strStatusPreConfig)
    Case Else
      Call SetupPowerShellV2()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AN"
      ' Nothing
    Case GetBuildfileValue("SetupKB956250") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB956250()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AO"
      ' Nothing
    Case strSetupNet4 <> "YES"
      ' Nothing
    Case strVersionNet4 >= "4"
      Call SetBuildFileValue("SetupNet4Status", strStatusPreConfig)
    Case Else
      Call SetupNet40()
  End Select

' ProcessId 2AP available for reuse

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AQ"
      ' Nothing
    Case GetBuildfileValue("SetupNet4x") <> "YES"
      ' Nothing
    Case Not strVersionNet4 >= "4"
      Call SetBuildFileValue("SetupNet4xStatus", strStatusBypassed)
    Case Else
      Call SetupNet4x()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AR"
      ' Nothing
    Case (strSetupDRUCtlr <> "YES") And (strSetupDRUClt <> "YES")
      ' Nothing
    Case Else
      Call SetupDRU()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ASZ"
      ' Nothing
    Case GetBuildfileValue("SetupRSAT") <> "YES"
      ' Nothing
    Case strOSVersion < "6.0"
      Call SetBuildFileValue("SetupNet20Status", strStatusBypassed)
    Case Else
      Call SetupRSAT() 
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AT"
      ' Nothing
    Case GetBuildfileValue("SetupPSRemote") <> "YES"
      ' Nothing
    Case strVersionPS = ""
      Call SetBuildFileValue("SetupPSRemoteStatus", strStatusBypassed)
    Case Else
      Call SetupPSRemote() 
  End Select

  Call SetProcessId("2AZ", " Pre-requisite processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupWindowsPreReqs()
  Call SetProcessId("2AA", "Setup Windows Pre-Reqs")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAA"
      ' Nothing
    Case Else
      Call SetupCOMPreReqs()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AABZ"
      ' Nothing
    Case strOSVersion >= "6.0"
      ' Nothing
    Case Else
      Call SetupWin2003PreReqs()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AACZ"
      ' Nothing
    Case strOSVersion < "6.0"
      ' Nothing
    Case strOSVersion >= "6.1"
      ' Nothing
    Case Else
      Call SetupWin2008PreReqs()  
  End Select


  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AADZ"
      ' Nothing
    Case strOSVersion < "6.1"
      ' Nothing
    Case strOSVersion >= "6.2"
      ' Nothing
    Case Else
      Call SetupWin2008R2PreReqs()  
  End Select


  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAEZ"
      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case strOSVersion >= "6.3"
      ' Nothing
    Case Else
      Call SetupWin2012PreReqs()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAFZ"
      ' Nothing
    Case strOSVersion < "6.3"
      ' Nothing
    Case strOSVersion >= "6.3A"
      ' Nothing
    Case Else
      Call SetupWin2012R2PreReqs()  
  End Select

  Call SetProcessId("2AAZ", " Setup Windows Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupCOMPreReqs()
  Call SetProcessId("2AAA", "Setup COM PreReqs")
  Dim strCLSIdNetCon, strCLSIdRunBroker

  strCLSIdNetCon    = GetBuildfileValue("CLSIdNetCon")
  strCLSIdRunBroker = GetBuildfileValue("CLSIdRunBroker")

  Select Case True
    Case strOSVersion >= "6.0"
      ' Nothing
    Case strCLSIdNetCon = ""
      ' Nothing
    Case Else
      Call SetDCOMSecurity("AppID\{" & strCLSIdNetCon & "}\")
  End Select

  Select Case True
    Case strOSVersion <= "6.1"
      ' Nothing
    Case strCLSIdRunBroker = ""
      ' Nothing
    Case Else
      Call SetDCOMSecurity("AppID\{" & strCLSIdRunBroker & "}\")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWin2003PreReqs()
  Call SetProcessId("2AAB", "Setup Windows 2003 Pre-Reqs")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AABA"
      ' Nothing
    Case GetBuildfileValue("SetupKB925336") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB925336()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AABB"
      ' Nothing
    Case GetBuildfileValue("SetupKB933789") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB933789()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AABC"
      ' Nothing
    Case GetBuildfileValue("SetupKB937444") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB937444()  
  End Select

  Call SetProcessId("2AABZ", " Setup Windows 2003 Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupWin2008PreReqs()
  Call SetProcessId("2AAC", "Setup Windows 2008 Pre-Reqs")

  Call SetProcessId("2AACZ", " Setup Windows 2008 Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupWin2008R2PreReqs()
  Call SetProcessId("2AAD", "Setup Windows 2008R2 Pre-Reqs")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AADA"
      ' Nothing
    Case GetBuildfileValue("SetupKB4019990") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB4019990("2AADA")  
  End Select

  Call SetProcessId("2AADZ", " Setup Windows 2008 Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupWin2012PreReqs()
  Call SetProcessId("2AAE", "Setup Windows 2012 Pre-Reqs")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAEA"
      ' Nothing
    Case GetBuildfileValue("SetupKB4019990") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB4019990("2AAEA")  
  End Select

  Call SetProcessId("2AAEZ", " Setup Windows 2008 Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupKB4019990(strProcess)
  Call SetProcessId(strProcess, "Install Windows 2008R2 & Windows 2012 KB4019990 D3DCompiler Fix")
  Dim objInstParm
  Dim strKB4019990msi

  Select Case True
    Case (strOSVersion = "6.1") And (strProcArc  = "X86")                 ' Windows 2008R2 X86
      strKB4019990msi = "Windows6.1-KB4019990-x86.msu"
    Case strOSVersion = "6.1"                                             ' Windows 2008R2 X64
      strKB4019990msi = "Windows6.1-KB4019990-x64.msu"
    Case Else                                                             ' Windows 2012
      strKB4019990msi = "Windows8-RT-KB4019990-x64.msu"
  End Select

'  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Common-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB2919355~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.3.1.14") ' Yes, it does use the KB2919355 PreCon
'  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB4019990", strKB4019990msi, objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupWin2012R2PreReqs()
  Call SetProcessId("2AAF", "Setup Windows 2012 R2 Pre-Reqs")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAFA"
      ' Nothing
    Case GetBuildfileValue("SetupKB2919442") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2919442()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAFB"
      ' Nothing
    Case GetBuildfileValue("SetupKB2919355") <> "YES"
      ' Nothing
    Case Else
      Call SetupKB2919355()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2AAFC"
      ' Nothing
    Case GetBuildfileValue("SetupKB3090973") <> "YES" 
      ' Nothing
    Case Else
      Call SetupKB3090973()  
  End Select

  Call SetProcessId("2AAFZ", " Setup Windows 2012 R2 Pre-Reqs" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupKB925336()
  Call SetProcessId("2AABA", "Install Windows 2003 KB925336 Installer Fix")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB925336\Description")
  Call RunInstall("KB925336", GetBuildfileValue("KB925336exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB933789()
  Call SetProcessId("2AABB", "Install Windows 2003 KB933789 Permissions Fix")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB933789\Description")
  Call RunInstall("KB933789", GetBuildfileValue("KB933789exe"), objInstParm)

  strPath           = "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB933789\"
  objWMIReg.GetStringValue strHKLM,strPath,"Description",strStatusKB933789

  Call ProcessEnd("")

End Sub


Sub SetupKB937444()
  Call SetProcessId("2AABC", "Install Windows 2003 KB937444 Filestream Fix")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB937444\Description")
  Call RunInstall("KB937444", GetBuildfileValue("KB937444exe"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB2919442()
  Call SetProcessId("2AAFA", "Install Windows 2012 R2 KB2919442 Update 1 Pre-Req")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Common-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB2919355~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.3.1.14") ' Yes, it does use the KB2919355 PreCon
  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB2919442", GetBuildfileValue("KB2919442msu"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB2919355()
  Call SetProcessId("2AAFB", "Install Windows 2012 R2 KB2919355 Update 1")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Common-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB2919355~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.3.1.14")
  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB2919355", GetBuildfileValue("KB2919355msu"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupKB3090973()
  Call SetProcessId("2AAFC", "Install Windows 2012 R2 KB3090973 MSDTC Fix")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Package_for_KB2919355~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB3090973_RTM_GM~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.3.1.0")
  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call RunInstall("KB3090973", GetBuildfileValue("KB3090973msu"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupSlipstream()
  Call SetProcessId("2AB", "Build Slipstreamed SQL Install Media")
  Dim objTempFile
  Dim strBase, strCUFile, strSPFile, strSlipSP, strSQLMediaPCUArc, strPCUSourceTemp, strPCUSourceArc, strCUSourceTemp, strCUSourceArc, strTempFile

  strSQLMedia       = strSQLMediaBase
  strSlipSP         = "N"
  strPath           = objShell.ExpandEnvironmentStrings("%Temp%")
  strSQLMediaPCU    = strPath & "\SQLMediaPCU"
  strTempFile       = strPath & "\FBTempSlipstream.tmp"
  Select Case True
    Case strSQLMedia = ""
      Call FBLog(strProcessIdDesc & strStatusBypassed & " - SQL Media Folder not found " & strSQLMediaOrig)
      Call SetBuildfileValue("SetupSlipstreamStatus", strStatusBypassed)
      Exit Sub
    Case strPathSQLSP = ""
      Call FBLog(strProcessIdDesc & strStatusBypassed & " - Service Pack Folder not found " & strPathSQLSPOrig)
      Call SetBuildfileValue("SetupSlipstreamStatus", strStatusBypassed)
      Exit Sub
    Case strSQLMedia = strSQLMediaPCU
      Call FBLog(strProcessIdDesc & strStatusPreConfig & " - Slipstream Media already built")
      Call SetBuildfileValue("SetupSlipstreamStatus", strStatusPreConfig)
      Exit Sub
    Case objFSO.FolderExists(strSQLMediaPCU)
      objFSO.DeleteFolder strSQLMediaPCU, 1
      Wscript.Sleep strWaitShort ' Wait for NTFS Cache to catch up to avoid Permissions error
  End Select
  Select Case True
    Case strSPLevel < "RTM"
      strBase       = strSPLevel
    Case Else
      strBase       = "RTM"
  End Select
  strDebugMsg1      = "SQLMediaPCU: " & strSQLMediaPCU
  objFSO.CreateFolder(strSQLMediaPCU)
  Wscript.Sleep strWaitShort
  strSQLMediaPCU    = strSQLMediaPCU & "\"
  strCUFile         = GetBuildfileValue("CUFile")
  strSPFile         = GetBuildfileValue("SPFile")

  Select Case True
    Case (strFileArc = "X86") Or (strWOWX86 = "TRUE")
      strSQLMediaArc    = strSQLMedia & "X86" & "\"
      strSQLMediaPCUArc = strSQLMediaPCU & "X86" & "\"
      strPCUSourceTemp  = strSQLMediaPCU & strSPLevel & "\"
      strPCUSourceArc   = strPCUSourceTemp & "X86" & "\"
      strCUSourceTemp   = strSQLMediaPCU & strSPCULevel & "\"
      strCUSourceArc    = strCUSourceTemp & "X86" & "\"
      strSPFile         = Replace(strSPFile, "ENU", strSQLLanguage, 1, -1, 1)
      strSPFile         = strPathSQLSP & strSPLevel & "\" & strSPFile
      strCUFile         = Replace(strCUFile, "ENU", strSQLLanguage, 1, -1, 1)
      strCUFile         = strPathSQLSP & strSPLevel & "\" & strCUFile 
    Case Else
      strSQLMediaArc    = strSQLMedia & "X64" & "\"
      strSQLMediaPCUArc = strSQLMediaPCU & "X64" & "\"
      strPCUSourceTemp  = strSQLMediaPCU & strSPLevel & "\"
      strPCUSourceArc   = strPCUSourceTemp & "X64" & "\"
      strCUSourceTemp   = strSQLMediaPCU & strSPCULevel & "\"
      strCUSourceArc    = strCUSourceTemp & "X64" & "\"
      strSPFile         = Replace(strSPFile, "ENU", strSQLLanguage, 1, -1, 1)
      strSPFile         = strPathSQLSP & strSPLevel & "\" & strSPFile 
      strCUFile         = Replace(strCUFile, "ENU", strSQLLanguage, 1, -1, 1)
      strCUFile         = strPathSQLSP & strSPLevel & "\" & strCUFile 
  End Select 

  Call FBLog(" Extracting " & strBase & " media: " & strSQLMedia)
  strDebugMsg1      = "Target: " & strSQLMediaPCU
  Select Case True
    Case strEdition = "EXPRESS"
      strCmd        = """" & strSQLMedia & strSQLExe & """ /U /X:""" & strSQLMediaPCU & """"
      Call Util_RunExec(strCmd, "", "", 0)
    Case GetBuildfileValue("StatusRobocopy") > ""
      strCmd        = "ROBOCOPY """ & Left(strSQLMedia, Len(strSQLMedia) - 1) & """ """ & Left(strSQLMediaPCU, Len(strSQLMediaPCU) - 1) & """ *.* /E /MT:25 /XJ /ZB "
      Call Util_RunExec(strCmd, "", "", 1)
    Case Else
      objFSO.CopyFolder Left(strSQLMedia, Len(strSQLMedia) - 1), Left(strSQLMediaPCU, Len(strSQLMediaPCU) - 1), 1
  End Select

  Set objTempFile   = objFSO.OpenTextFile(strTempFile, 2, True)
  strPath           = "Microsoft.SQL.Chainer.PackageData.dll"
  objTempFile.WriteLine strPath
  objTempFile.Close
  Select Case True
    Case strSPLevel < "RTM"
      ' Nothing
    Case strSetupSP <> "YES"
      ' Nothing
    Case Left(strSPLevel, 2) <> "SP"
      ' Nothing
    Case Not objFSO.FileExists(strSPFile)
      Call FBLog(strProcessIdDesc & strStatusBypassed & " - " & strSPLevel & " file not found: " & strSPFile)
      Call SetBuildfileValue("SetupSlipstreamStatus", strStatusBypassed)
      Exit Sub
    Case Else
      Call FBLog(" Extracting SP media: " & strSPFile)
      strSlipSP     = "Y"
      strDebugMsg1  = "Source: " & strSPFile
      strDebugMsg2  = "Target: " & strPCUSourceTemp
      strCmd        = """" & strSPFile & """ /QUIET /X:""" & strPCUSourceTemp & """"
      Call Util_RunExec(strCmd, "", "", 0)
      If strSQLVersion >= "SQL2008" Then
        Call DebugLog("Replacing original media with SP media")
        strCmd      = "XCOPY """ & strPCUSourceTemp & "Setup.*"" """ & strSQLMediaPCU & "*.*"" /H /R /Y /Z"
        Call Util_RunExec(strCmd, "", "", 0)
        strCmd      = "XCOPY """ & strPCUSourceArc & "*.*"" """ & strSQLMediaPCUArc & "*.*"" /H /R /Y /Z /EXCLUDE:" & strTempFile
        Call Util_RunExec(strCmd, "", "", 0)
      End If
      strPCUSource = strPCUSourceTemp
      Call SetBuildfileValue("PCUSource",   strPCUSource)
  End Select

  Select Case True
    Case strSPLevel < "RTM"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strSetupSPCU <> "YES"
      ' Nothing
    Case Left(strSPCULevel, 2) <> "CU"
      ' Nothing
    Case strSPLevel > "SP" And strSlipSP <> "Y"
      Call DebugLog("SP requested but not found, so do not set up CU for SP")
    Case Not objFSO.FileExists(strCUFile)
      Call FBLog(strSPCULevel & " file not found: " & strCUFile)
    Case Else
      Call FBLog(" Extracting CU media: " & strCUFile)
      strDebugMsg1  = "Source: " & strCUFile
      strDebugMsg2  = "Target: " & strCUSourceTemp
      strCmd        = """" & strCUFile & """ /QUIET /X:""" & strCUSourceTemp & """"
      Call Util_RunExec(strCmd, "", "", 0)
      Call DebugLog("Replacing original media with CU media")
      strCmd        = "XCOPY """ & strCUSourceTemp & "Setup.*"" """ & strSQLMediaPCU & "*.*"" /H /R /Y /Z"
      Call Util_RunExec(strCmd, "", "", 0)
      strCmd        = "XCOPY """ & strCUSourceArc & "*.*"" """ & strSQLMediaPCUArc & "*.*"" /H /R /Y /Z /EXCLUDE:" & strTempFile
      Call Util_RunExec(strCmd, "", "", 0)
      strCUSource   = strCUSourceTemp
      Call SetBuildfileValue("CUSource", strCUSource)
      Call SetupSNAC(strSQLMediaPCUArc)
  End Select

  If objFSO.FileExists(strTempFile) Then
    objFSO.DeleteFile(strTempFile)
  End If
  strSQLMedia       = strSQLMediaPCU 
  strSQLMediaArc    = strSQLMediaPCUArc
  Call SetBuildfileValue("PathSQLMedia", strSQLMedia) 
  Call SetBuildfileValue("SQLMediaArc",  strSQLMediaArc)

  If strSQLVersion >= "SQL2008" Then 
    Call SetBuildfileValue("SPInclude",   ", included in Slipstream")
  End If

  Call SetBuildfileValue("SetupSlipstreamStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSNAC(strSQLMediaPCUArc)
  Call DebugLog("SetupSNAC:")
  Dim strPathInst, strSNACFile

  strSNACFile       = GetBuildfileValue("SNACFile")
  strPathInst       = GetPathInst(strSNACFile, strPathSQLSP & strSPLevel & "\", "")
  strDebugMsg1      = "Source: " & strPathInst
  strDebugMsg2      = "Target: " & strSQLMediaPCUArc

  Select Case True
    Case Not objFSO.FileExists(strPathInst)
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strCmd        = "%COMSPEC% /D /C COPY /B """ & strPathInst & """ """ & strSQLMediaPCUArc & "Tools\Setup\sqlncli.msi"" /Y"
      Call Util_RunExec(strCmd, "", "", 0)
    Case strSQLVersion = "SQL2008"
      strCmd        = "%COMSPEC% /D /C COPY /B """ & strPathInst & """ """ & strSQLMediaPCUArc & "Setup\" & strFileArc & "\sqlncli.msi"" /Y"
      Call Util_RunExec(strCmd, "", "", 0)
    Case Else
      ' Nothing
  End Select

End Sub


Sub SetupMSDTC()
  Call SetProcessId("2AC", "Setup MSDTC")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACA"
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case Else
      Call CheckDTCKerberos() 
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACB"
      ' Nothing
    Case strSetupDTCCID <> "YES"
      ' Nothing
    Case strSetupDTCNetAccessStatus = strStatusComplete
      If GetBuildfileValue("SetupDTCCIDStatus") = "" Then
        Call SetBuildfileValue("SetupDTCCIDStatus", strStatusBypassed)
      End If
    Case Else
      Call SetupDTCNewCID()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACC"
      ' Nothing
    Case strSetupDTCNetAccess <> "YES"
      ' Nothing
    Case strSetupDTCNetAccessStatus = strStatusComplete
      ' Nothing
    Case Else
      Call SetupDTCNetAccess()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACD"
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case (strOSVersion < "6.0") And (strDTCClusterRes <> "")
      Call SetBuildFileValue("SetupDTCClusterStatus", strStatusPreConfig)
    Case Else
      Call BuildDTCCluster()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACE"
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case (strOSVersion < "6.0") And (strDTCClusterRes <> "")
      ' Nothing
    Case Else
      Call SetDTCClusterAccess()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACF"
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case Else
      Call RegisterDTCClusterRes()  
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ACG"
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case Else
      Call CheckDTCClusterService() 
  End Select

  Call SetProcessId("2ACZ", " Setup MSDTC" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CheckDTCKerberos()
  Call SetProcessId("2ACA", "Check Kerberos Status for MSDTC") 
  Dim strClusGroups

  strClusGroups     = GetAccountAttr(strClusterName, strUserDnsDomain, "memberOf")
  Select Case True
    Case InStr(" " & strClusGroups & " ", strGroupMSA & " ") = 0
      strDebugMsg1  = "Cluster Groups: " & strClusGroups
      strDebugMsg2  = "MSA Group     : " & strGroupMSA
      Call SetBuildMessage(strMsgError, "Process Kerberos command file to allow MSDTC Cluster build to continue")
    Case GetBuildfileValue("RebootStatus") <> "Pending"
      ' Nothing
    Case GetBuildfileValue(strClusterName & "MSAGroup") <> ""
      Call SetupReboot("2ACA", "Prepare for MSDTC Cluster")
    Case Else
      ' Nothing
  End Select

End Sub


Sub SetupDTCNewCID()
  Call SetProcessId("2ACB", "Setup new CID for MSDTC") 

  strCmd            = "NET STOP MSDTC"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  Call Util_RunExec("MSDTC -UNINSTALL", "", strResponseYes, -2147023836)
  WScript.Sleep strWaitLong

  Call Util_RunExec("REG DELETE ""HKEY_LOCAL_MACHINE\System\CurrentControlSet\Services\MSDTC"" /f",   "", strResponseYes, -1)
  Call Util_RunExec("REG DELETE ""HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSDTC"" /f",   "", strResponseYes, -1)
  Call Util_RunExec("REG DELETE ""HKEY_CLASSES_ROOT\CID"" /f",   "", strResponseYes, -1)
  Call Util_RunExec("REG DELETE ""HKEY_CLASSES_ROOT\CID.Local"" /f",   "", strResponseYes, -1)

  Call Util_RunExec("MSDTC -INSTALL",   "", strResponseYes, 0)
  WScript.Sleep strWaitLong

  strReboot         = "Pending"
  Call SetBuildfileValue("RebootStatus", strReboot)

  Call SetBuildFileValue("SetupDTCCIDStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDTCNetAccess()
  Call SetProcessId("2ACC", "Setup DTC Network Access")
  Dim strDTCPort

  Select Case True
    Case strOSVersion >= "6.0"
      strCmd        = "NET START MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
    Case strClusterAction <> ""
      ' Nothing
    Case Else
      strCmd        = "NET START MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End Select

  strDTCPort        = SetupPortRange("Distributed Transaction Coordinator Local (Port)", strTCPPortDTC, 200)  

  Select Case True
    Case strOSVersion >= "6.0"
      strCmd        = "NET STOP MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case strClusterAction <> ""
      ' Nothing
    Case Else
      strCmd        = "NET STOP MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      Call DTCNetAccessRegistry("HKLM\SOFTWARE\Microsoft\MSDTC", strDTCPort)
    Case strClusterAction <> ""
      strCmd        = "SC CONFIG ""MSDTC"" START= DEMAND" 
      Call Util_RunExec(strCmd, "", "", 2)
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      strCmd        = "NET START MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case strClusterAction <> ""
      ' Nothing
    Case Else
      strCmd        = "MSDTC -RESETLOG"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "NET START MSDTC"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  strCmd            = strHKLMFB & "SetupDTCNetAccessStatus"
  Call Util_RegWrite(strCmd, strStatusComplete, "REG_SZ") 

  Call SetBuildFileValue("SetupDTCNetAccessStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub BuildDTCCluster()
  Call SetProcessId("2ACD", "Setup DTC Cluster")

  If strOSVersion < "6.0" Then
    Call ClearDefaultMSDTC()
  End If

  Call SetBuildFileValue("SetupDTCClusterStatus", strStatusProgress)
  Call BuildCluster("DTC", strClusterGroupDTC, strClusterNameDTC, strClusterNetworkDTC, "Distributed Transaction Coordinator", "", "", "", "VolDTC", "H")
'  Call BuildCluster("DTC", strClusterGroupDTC, strClusterNameDTC, strClusterNetworkDTC, "Distributed Transaction Coordinator", "", "", "Cluster\Resources\{GUID}\MSDTCPRIVATE\MSDTC\", "VolDTC", "H")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ClearDefaultMSDTC()
  Call DebugLog("ClearDefaultMSDTC:")
  Dim arrGroups,arrResources
  Dim objGroup, objResource
  Dim strGroupName, strResourceName, strResourceType

  strPath           = "Cluster\Groups"
  objWMIReg.EnumKey strHKLM, strPath, arrGroups
  For Each objGroup In arrGroups
    strPathNew      = strPath & "\" & objGroup
    objWMIReg.GetStringValue strHKLM, strPathNew, "Name",      strGroupName
    If strGroupName = "Cluster Group" Then
      objWMIReg.GetMultiStringValue strHKLM, strPathNew, "Contains", arrResources
      For Each objResource In arrResources
        strPath     = "HKLM\Cluster\Resources\" & objResource
        strResourceName = objShell.RegRead(strPath & "\Name")
        strResourceType = objShell.RegRead(strPath & "\Type")
        If strResourceType = "Distributed Transaction Coordinator" Then
          strCmd    = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResourceName  & """ /DELETE"
          Call Util_RunExec(strCmd, "", strResponseYes, 5010)
        End If
      Next
    End If
  Next

  strCmd        = "SC CONFIG ""MSDTC"" DEPEND= ""RPCSS""/""SamSS""/""ClusSvc"""
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

End Sub


Sub SetDTCClusterAccess()
  Call SetProcessId("2ACE", "Set DTC Cluster Access")
' Review http://blogs.msdn.com/b/distributedservices/archive/2010/05/04/issue-with-incoming-caller-authentication-for-ms-dtc-on-windows-2008-cluster.aspx
  Dim arrResources
  Dim objResource
  Dim strDTCPort, strMapName, strMapping, strPathReg, strResourceName, strResourceType

  Call SetResourceOn(strClusterGroupDTC, "GROUP") ' need to ensure on current node

  strDTCPort        = SetupPortRange("Distributed Transaction Coordinator " & strClusterGroupDTC  & " (Port)", strTCPPortDTC, 200)

  Call DebugLog("Setup Network Access for MSDTC Cluster")
  strPathReg        = "Cluster\Resources"
  objWMIReg.EnumKey strHKLM,strPathReg,arrResources
  For Each objResource In arrResources
    strPath         = "HKLM\" & strPathReg & "\" & objResource
    strResourceName = objShell.RegRead(strPath & "\Name")
    strResourceType = objShell.RegRead(strPath & "\Type")
    Select Case True
      Case strResourceType <> "Distributed Transaction Coordinator"
        ' Nothing
      Case strResourceName <> strClusterNameDTC
        ' Nothing
      Case strOSVersion < "6.0"
        Call DTCNetAccessRegistry(GetW2003DTCRegistry(strPath), strDTCPort)
      Case Else
        Call DTCNetAccessRegistry(strPath & "\MSDTCPRIVATE\MSDTC", strDTCPort)
    End Select
  Next

  Select Case True
    Case strOSVersion < "6.0"
      strCmd        = "MSDTC -RESETLOG"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
    Case Else
      strCmd        = "MSDTC -RESETCLUSTERTMLOG """ & strClusterNameDTC & """"
      Call Util_RunExec(strCmd, "", strResponseYes, -1) ' Ignore error -2146434815 (file already exists)
  End Select

  If strOSVersion >= "6.0" Then ' Use Local DTC for COM+ Applications
    strMapName      = "DLLHOSTMapping"
    strPath         = "Cluster\MSDTC\TMMapping\Exe\" & strMapName & "\"
    strMapping      = ""
    objWMIReg.GetStringValue strHKLM,strPath,"Name",strMapping
    If strMapping > "" Then
      objShell.RegDelete "HKLM\" & strPath
    End If
    strCmd          = "MSDTC -tmMappingSet -name """ & strMapName & """ -exe """ & strDirSys & "\System32\dllhost.exe"" -local "
    Call Util_RunCmdASync(strCmd, 0)
    Wscript.Sleep strWaitShort
    strCmd          = "WMIC process WHERE ""CommandLine LIKE '%MSDTC.EXE%'"" CALL terminate"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Function SetupPortRange(strDescription, strStartPort, intPortNum)
  Call DebugLog("SetupPortRange: " & strDescription)
  Dim objExec
  Dim arrPorts
  Dim intPort, intUBound
  Dim strPort, strPortRange, strPortTest, strReadAll, strRpcPath

  strPort           = ""
  intPort           = CLng(strStartPort)
  strPortRange      = CStr(intPort) & "-" & CStr(intPort + intPortNum)
  strPortTest       = ":" & CStr(intPort)
  strRpcPath        = "SOFTWARE\Microsoft\Rpc\Internet\"

  objWMIReg.GetMultiStringValue strHKLM, strRpcPath, "Ports", arrPorts
  strCmd            = "NETSTAT -an"
  Set objExec       = objShell.Exec(strCmd)
  strReadAll        = objExec.StdOut.ReadAll

  Select Case True
    Case strOSVersion >= "6.1"
      While strPort = ""
        Select Case True
          Case Instr(strReadAll, strPortTest) = 0
            strPort = CStr(intPort)
          Case Else
            intPort = intPort + 1
            strPortTest = ":" & CStr(intPort)
        End Select
      WEnd
      strCmd        = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""" & strDescription & """ "
      strCmd        = strCmd & "LOCALPORT=" & strPort & " PROTOCOL=TCP "
      strCmd        = strCmd & "ACTION=ALLOW PROFILE=DOMAIN DIR=IN "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
    Case Else
      strPort       = strStartPort
      Select Case True
        Case IsNull(arrPorts)
          intUBound = 0
          Redim arrPorts(intUBound)
          arrPorts(intUBound) = strPortRange
          objWMIReg.CreateKey strHKLM, strRpcPath
        Case arrPorts(0) <> strPortRange
          intUBound = UBound(arrPorts) + 1
          Redim Preserve arrPorts(intUBound)
          arrPorts(intUBound) = strPortRange
      End Select
      objWMIReg.SetMultiStringValue strHKLM, strRpcPath, "Ports", arrPorts
      strCmd        = "HKLM\" & strRpcPath & "PortsInternetAvailable"
      Call Util_RegWrite(strCmd, "Y", "REG_SZ")
      strCmd        = "HKLM\" & strRpcPath & "UseInternetPorts"
      Call Util_RegWrite(strCmd, "Y", "REG_SZ")
  End Select

  SetupPortRange    = strPort

End Function


Function GetW2003DTCRegistry(strPathReg)
  Call DebugLog("GetW2003DTCRegistry:")
  Dim arrKeys
  Dim objKey
  Dim strDTCKey, strMSDTCGUID, strDTCPath

  strDTCKey         = ""
  strDTCPath        = Mid(strPathReg, 6)
  objWMIReg.EnumKey strHKLM,strDTCPath,arrKeys
  For Each objKey In arrKeys
    objWMIReg.GetStringValue strHKLM,strDTCPath & "\" & objKey,"MSDTC",strMSDTCGUID
    If strMSDTCGUID > "" Then
      strDTCKey     = objKey
    End If
  Next

  GetW2003DTCRegistry = strPathReg & "\" & strDTCKey

End Function


Sub DTCNetAccessRegistry(strRegKey, strPort)
  Call DebugLog("DTCNetAccessRegistry: " & strRegKey)

  Call DebugLog("Enable Network DTC Access")
  strCmd            = strRegKey & "\Security\NetworkDtcAccess"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
  strCmd            = strRegKey & "\Security\NetworkDtcAccessAdmin"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
  strCmd            = strRegKey & "\Security\NetworkDtcAccessClients"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 

  Call DebugLog("Allow Inbound and Outbound transactions")
  strCmd            = strRegKey & "\Security\NetworkDtcAccessTransactions"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
  strCmd            = strRegKey & "\Security\NetworkDtcAccessInbound"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
  strCmd            = strRegKey & "\Security\NetworkDtcAccessOutbound"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD")
  strCmd            = strRegKey & "\Security\XaTransactions"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD")

  Call DebugLog("Configure authentication")
  Select Case True
    Case (strOSVersion < "6.0") And (strClusterAction <> "") And (Left(strRegKey,12) = "HKLM\Cluster") ' Configure Incoming Caller Authentication for W2003 clustered environment (See KB899191)
      strCmd        = strRegKey & "\Security\AllowOnlySecureRpcCalls"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
      strCmd        = strRegKey & "\Security\FallbackToUnsecureRPCIfNecessary"
      Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
      strCmd        = strRegKey & "\Security\TurnOffRpcSecurity"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
    Case (strOSVersion < "6.0") And (strClusterAction <> "") ' Configure Incoming Caller Authentication for W2003 clustered environment (See KB899191)
      strCmd        = strRegKey & "\AllowOnlySecureRpcCalls"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
      strCmd        = strRegKey & "\FallbackToUnsecureRPCIfNecessary"
      Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
      strCmd        = strRegKey & "\TurnOffRpcSecurity"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
    Case Else ' Configure Mutual authentication for W2003 non-clustered or any W2008 environment
      strCmd        = strRegKey & "\AllowOnlySecureRpcCalls"
      Call Util_RegWrite(strCmd, 1, "REG_DWORD") 
      strCmd        = strRegKey & "\FallbackToUnsecureRPCIfNecessary"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
      strCmd        = strRegKey & "\TurnOffRpcSecurity"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD") 
  End Select

  Select Case True
    Case strOSVersion >= "6.1" 
      Call DebugLog("Configure DTC Port")
      strCmd        = strRegKey & "\ServerTcpPort"
      Call Util_RegWrite(strCmd, strPort, "REG_DWORD")
  End Select

End Sub


Sub RegisterDTCClusterRes()
  Call SetProcessId("2ACF", "Register DTC Cluster Resource")
  Dim arrResources
  Dim objResource
  Dim strPathReg, strResourceName, strResourceType

  Call AddOwner(strClusterNameDTC)

  strPathReg        = "Cluster\Resources"
  objWMIReg.EnumKey strHKLM,strPathReg,arrResources
  For Each objResource In arrResources
    strPath         = "HKLM\" & strPathReg & "\" & objResource
    strResourceName = objShell.RegRead(strPath & "\Name")
    strResourceType = objShell.RegRead(strPath & "\Type")
    Select Case True
      Case strResourceType <> "Distributed Transaction Coordinator"
        ' Nothing
      Case strResourceName <> strClusterNameDTC
        ' Nothing
      Case strOSVersion < "6.0"
        Call SetDTCClusterRes(objResource)
      Case Else
        Call SetDTCClusterRes(objResource)
    End Select
  Next

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetDTCClusterRes(strClusterRes)
  Call DebugLog("SetDTCClusterRes: " & strClusterRes)

  If strDTCClusterRes = "" Then
    strDTCClusterRes = strClusterRes
    Call Util_RegWrite(strHKLMFB & "DTCClusterRes", strDTCClusterRes, "REG_SZ")
    Call SetBuildfileValue("DTCClusterRes", strDTCClusterRes)
  End If

  strInstDTCClusterRes = strClusterRes
  Call SetBuildfileValue("InstDTCClusterRes", strInstDTCClusterRes)

End Sub


Sub CheckDTCClusterService()
  Call SetProcessId("2ACG", "Check DTC Cluster Service")
  Dim colClusResources
  Dim objClusResource

  Select Case True
    Case strOSVersion < "6.0"
      strCmd        = "MSDTC -RESETLOG"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End Select

  Set colClusResources = GetClusterResources
  For Each objClusResource In colClusResources
    Select Case True
      Case strClusterNetworkDTC = ""
        ' Nothing
      Case objClusResource.Name <> strClusterNetworkDTC
        ' Nothing
      Case objClusResource.State <> 4
        ' Nothing
      Case Else
        Call SetBuildMessage(strMsgInfo,  "DTC Network Name has State=Failed, this may prevent SQL Cluster install succeeding")
        Call SetBuildMessage(strMsgError, "Check https://github.com/SQL-FineBuild/Common/wiki/Delegation-Of-Control before restarting FineBuild")
    End Select
  Next

  Call SetBuildFileValue("SetupDTCClusterStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupJRE()
  Call SetProcessId("2AD", "Install Java Runtime Environment")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "ParmLog",    "/l")
  Call SetXMLParm(objInstParm, "ParmReboot", "")
  Call SetXMLParm(objInstParm, "ParmSilent", "/s")
  If strSQLVersion >= "SQL2019" Then
'
  End If
  Call RunInstall("JRE", GetBuildfileValue("JREexe"), objInstParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNet20()
  Call SetProcessId("2AE", "Install .Net 2.0")
  Dim objInstParm
  Dim strDate, strPathLog

  If GetNet3Path("") <> "" Then
    strSetupNet3    = "YES"
    Call SetBuildfileValue("SetupNet3", strSetupNet3)
    Call SetBuildFileValue("SetupNet20Status", strStatusBypassed & ", .Net 3.5 will be installed")
    Call FBLog(" " & strProcessIdDesc & strStatusBypassed & ", .Net 3.5 will be installed")
    Exit Sub
  End If

  strPathLog        = Mid(strSetupLog, 2) & strProcessIdLabel & " .Net2.0Logs"
  Call SetXMLParm(objInstParm, "SetupOption", "Extract")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Select Case True
    Case strEdition = "EXPRESS"
      Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & "redist\2.0\")
    Case Else
      Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & "Servers\redist\2.0\")
  End Select
  Call SetXMLParm(objInstParm, "PathMain",    strPathAddComp & "redist\2.0\")
  Call SetXMLParm(objInstParm, "PathLog",     strPathLog & "\" & strProcessIdDesc & ".txt")
  Call SetXMLParm(objInstParm, "ParmExtract", "/C /Q /T:")
  Call SetXMLParm(objInstParm, "ParmLog",     "/L")
  Call SetXMLParm(objInstParm, "ParmSilent",  "/Q")
  Call SetXMLParm(objInstParm, "InstFile",    "install.exe")
  Call RunInstall("Net20", "dotnetfx.exe", objInstParm)

  If GetBuildfileValue("SetupNet20Status") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Collect .Net 2.0 Install Logs")
  strDate           = Right("0" & DatePart("m", Date()), 2) & "-" & Right("0" & DatePart("d", Date()), 2) & "-" & Right("0" & DatePart("yyyy", Date()), 4)
  strCmd            = "XCOPY """ & strPathTemp & "\*.txt"" """ & strPathLog & "\*.*"" /D:" & strDate & " /H /R /Y /Z"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)
  strCmd            = "XCOPY """ & strPathTemp & "\*.log"" """ & strPathLog & "\*.*"" /D:" & strDate & " /H /R /Y /Z"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  Call SetBuildFileValue("SetupNet20Status", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNet35SP1()
  Call SetProcessId("2AF", "Install .Net 3.5 SP1")
  Dim objInstParm
  Dim strFeatures, strNet3File, strNet3Path

  strNet3Path       = GetNet3Path("")
  Select Case True
    Case strNet3Path = ""
      Call SetBuildMessage(strMsgError, ".Net 3.5 SP1 install file is missing")
    Case strOSVersion < "6.1"    ' Windows 2008 or below
      strNet3File   = Right(strNet3Path, Len(strNet3Path) - InstrRev(strNet3Path, "\"))
      Call SetXMLParm(objInstParm, "PathMain",   Left(strNet3Path, InstrRev(strNet3Path, "\")))
    Case (strOSVersion >= "6.3A") And (strFileArc = "X86") ' Windows 10 X86
      strNet3File   = "DISM.exe"
      strFeatures   = "/online /enable-feature /all /LimitAccess /Source:""" & strPathAddComp & "redist\sxs\x86"" "
      strFeatures   = strFeatures & "/featurename:NetFX3 "
      Call SetXMLParm(objInstParm, "ParmXtra",   strFeatures)
    Case (strOSVersion >= "6.2") And (Instr(strOSType, "CORE") > 0) ' Windows 8 or above Server Core
      strNet3File   = "DISM.exe"
      strFeatures   = "/online /enable-feature /all /LimitAccess /Source:""" & strPathAddComp & "redist\sxs"" "
      strFeatures   = strFeatures & "/featurename:ServerCore-WOW64 "
      strFeatures   = strFeatures & "/featurename:NetFx3 "
      Call SetXMLParm(objInstParm, "ParmXtra",   strFeatures)
    Case (strOSVersion >= "6.2") And (Instr(strOSType, "CLIENT") > 0) ' Windows 8 or above Server Core
      strNet3File   = "DISM.exe"
      strFeatures   = "/online /enable-feature /all /LimitAccess /Source:""" & strPathAddComp & "redist\sxs"" "
      strFeatures   = strFeatures & "/featurename:NetFx3 "
      Call SetXMLParm(objInstParm, "ParmXtra",   strFeatures)
    Case strOSVersion >= "6.2"   ' Windows 8 or above
      strNet3File   = "DISM.exe"
      strFeatures   = "/online /enable-feature /all /LimitAccess /Source:""" & strPathAddComp & "redist\sxs"" "
      strFeatures   = strFeatures & "/featurename:NetFx3ServerFeatures "
      strFeatures   = strFeatures & "/featurename:NetFx3 "
      Call SetXMLParm(objInstParm, "ParmXtra",   strFeatures)
    Case Instr(strOSType, "CORE") > 0                  ' Windows 2008 R2 Server Core
      strNet3File   = "PKGMGR.exe"
      Call SetXMLParm(objInstParm, "ParmXtra",   "/IU:NetFx2-ServerCore;NetFx3-ServerCore;ServerCore-WOW64;NetFx2-ServerCore-WOW64;NetFx3-ServerCore-WOW64 ")
    Case Else                                          ' Windows 2008 R2
      strNet3File   = "PKGMGR.exe"
      Call SetXMLParm(objInstParm, "ParmXtra",   "/IU:NetFx3 ")
  End Select

  If strOSVersion >= "6.3A" Then ' Windows 2016
    Call SetXMLParm(objInstParm, "ParmMonitor",  "6") ' Force reboot if install hangs
  End If

  Call RunInstall("Net3", strNet3File, objInstParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Function GetNet3Path(strReturn)
  Call DebugLog("GetNet3Path:")
  Dim strPath

  strPath           = ""
  strPathAlt        = ""
  Select Case True
    Case strOSVersion = "6.1"
      strPath       = "PKGMGR"
    Case (strOSVersion = "6.2") And (strProcArc = "X86")
      strPath       = strPathAddComp & "redist\sxs\x86_microsoft-windows-netfx3-core_31bf3856ad364e35_6.2.9200.16384_none_644bdaeaab87aa21\FrameworkList.xml"
      strPathAlt    = ""
    Case strOSVersion = "6.2"
      strPath       = strPathAddComp & "redist\sxs\amd64_microsoft-windows-netfx3-core_31bf3856ad364e35_6.2.9200.16384_none_c06a766e63e51b57\FrameworkList.xml"
      strPathAlt    = ""
    Case (strOSVersion = "6.3") And (strProcArc = "X86")
      strPath       = strPathAddComp & "redist\sxs\x86_microsoft-windows-netfx3-core_31bf3856ad364e35_6.3.9600.16384_none_fc409390f5ba7a9e\FrameworkList.xml"
      strPathAlt    = ""
    Case strOSVersion = "6.3"
      strPath       = strPathAddComp & "redist\sxs\amd64_microsoft-windows-netfx3-core_31bf3856ad364e35_6.3.9600.16384_none_585f2f14ae17ebd4\FrameworkList.xml"
      strPathAlt    = ""
    Case strOSVersion >= "6.3A"
      strPath       = strPathAddComp & "redist\sxs\microsoft-windows-netfx3-ondemand-package.cab"
      strPathAlt    = ""
    Case strSQLVersion = "SQL2008"
      strPath       = strSQLMedia & strFileArc & "\redist\DotNetFrameworks\dotNetFx35setup.exe"
      strPathAlt    = strPathAddComp & "redist\DotNetFrameworks\dotNetFx35setup.exe"
    Case Else
      strPath       = strSQLMedia & "redist\DotNetFrameworks\dotNetFx35setup.exe"
      strPathAlt    = strPathAddComp & "redist\DotNetFrameworks\dotNetFx35setup.exe"
  End Select
  Select Case True
    Case strPath = ""
      ' Nothing
    Case strPath = "PKGMGR"
      ' Nothing
    Case objFSO.FileExists(strPath)
      ' Nothing
    Case objFSO.FileExists(strPathAlt)
      strPath       = strPathAlt
    Case strReturn = "Y"
      ' Nothing
    Case Else
      strPath       = ""
  End Select

  Call DebugLog("Net3 Install Path: " & strPath)
  GetNet3Path       = strPath

End Function


Sub SetupIIS()
  Call SetProcessId("2AG", "Install IIS")

  Select Case True
    Case strOSVersion < "6.0"                                   ' Windows 2003 or below
      Call InstallIISW2003
    Case strOSVersion < "6.2"                                   ' Windows Vista through Windows 2008 R2
      Call InstallIISW2008
    Case Else                                                   ' Windows 2012 R2 and above
      Call InstallIISW2012
  End Select

  If GetBuildfileValue("SetupIISStatus") <> strStatusProgress Then
    Exit Sub
  End If

  Call DebugLog("Enable IIS Remote Management")
  strCmd            = "HKLM\SOFTWARE\Microsoft\WebManagement\Server\EnableRemoteManagement"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD") 

  Call SetBuildfileValue("SetupIISStatus",  strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub InstallIISW2003()
  Call DebugLog("InstallIISW2003:")
  Dim arrSources
  Dim objInstParm, objTempFile
  Dim strParmXtra, strPathIISExe, strIISPath, strIISReg, strTempFile

  strIISPath        = strPathAddComp & "IIS6" & "\" & strFileArc
  If Not objFSO.FolderExists(strIISPath& "\I386") Then
    Call FBLog(strProcessIdDesc & strStatusBypassed & " - IIS install media not found: " & strIISPath& "\I386")
    Call SetBuildfileValue("SetupIISStatus",  strStatusBypassed)
    Exit Sub
  End If

  strIISReg         = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Setup"
  strDebugMsg1      = "Reg Path: " & strIISReg
  arrSources        = Array(strIISPath)
  objWMIReg.SetMultiStringValue strHKLM, Mid(strIISReg, 6), "Installation Sources", arrSources
  Call Util_RegWrite(strIISReg & "\SourcePath", strIISPath, "REG_SZ")
  Call Util_RegWrite(strIISReg & "\ServicePackSourcePath", strIISPath, "REG_SZ")
  strCmd            = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SourcePath"
  Call Util_RegWrite(strCmd, strIISPath, "REG_SZ")

  strTempFile       = strPathTemp & "\FBTempIIS.tmp"
  Set objTempFile   = objFSO.OpenTextFile(strTempFile, 2, True)
  objTempFile.WriteLine "[Components]"
  objTempFile.WriteLine "iis_common     = on"
  objTempFile.WriteLine "iis_ftp        = off"
  objTempFile.WriteLine "iis_www        = on"
  objTempFile.WriteLine "iis_inetmgr    = on"
  objTempFile.WriteLine "iis_asp        = on"
  objTempFile.WriteLine "aspnet         = on"
  objTempFile.WriteLine "complusnetwork = on"
  objTempFile.WriteLine "[InternetServer]"
  objTempFile.WriteLine "PathWWWRoot    = " & GetBuildfileValue("IISRoot")
  objTempFile.Close

  strParmXtra       = "/i:%windir%\sysoc.inf /u:""" & strTempFile & """ "
  Call SetXMLParm(objInstParm, "PreConKey",   "SYSTEM\CurrentControlSet\Services\W3SVC\ImagePath")
  Call SetXMLParm(objInstParm, "PathMain",    strPathSys)
  Call SetXMLParm(objInstParm, "ParmXtra",    strParmXtra)
  Call SetXMLParm(objInstParm, "ParmLog",     "")
  Call SetXMLParm(objInstParm, "ParmReboot",  "/r")
  Call SetXMLParm(objInstParm, "ParmSilent",  "")
  Call RunInstall("IIS", "SYSOCMGR.EXE", objInstParm)

  objFSO.DeleteFile strTempFile

End Sub


Sub InstallIISW2008()
  Call DebugLog("InstallIISW2008:")
' Parents must be installed before children
  Dim objInstParm

  strFeatures       = "/IU:IIS-WebServerRole"
  strFeatures       = strFeatures & ";WAS-WindowsActivationService"
  strFeatures       = strFeatures & ";WAS-ConfigurationAPI"
  strFeatures       = strFeatures & ";WAS-NetFxEnvironment"
  strFeatures       = strFeatures & ";IIS-WebServer"
  strFeatures       = strFeatures & ";IIS-CommonHttpFeatures" 
  strFeatures       = strFeatures & ";IIS-StaticContent"
  strFeatures       = strFeatures & ";IIS-DefaultDocument"
  strFeatures       = strFeatures & ";IIS-DirectoryBrowsing"
  strFeatures       = strFeatures & ";IIS-HttpErrors"
  strFeatures       = strFeatures & ";IIS-HealthAndDiagnostics"
  strFeatures       = strFeatures & ";IIS-HttpLogging"
  strFeatures       = strFeatures & ";IIS-RequestMonitor"
  strFeatures       = strFeatures & ";IIS-Performance"
  strFeatures       = strFeatures & ";IIS-HttpCompressionStatic"
  strFeatures       = strFeatures & ";IIS-Security"
  strFeatures       = strFeatures & ";IIS-WindowsAuthentication"
  strFeatures       = strFeatures & ";IIS-RequestFiltering"
  strFeatures       = strFeatures & ";IIS-WebServerManagementTools"
  strFeatures       = strFeatures & ";IIS-ManagementService"
  strFeatures       = strFeatures & ";IIS-ApplicationDevelopment"
  strFeatures       = strFeatures & ";IIS-ISAPIExtensions"
  strFeatures       = strFeatures & ";IIS-ISAPIFilter"
  strFeatures       = strFeatures & ";IIS-NetFxExtensibility"
  strFeatures       = strFeatures & ";IIS-ASPNET"
  strFeatures       = strFeatures & ";WCF-HTTP-Activation"
  strFeatures       = strFeatures & ";WCF-NonHTTP-Activation"

  If strSQLVersion = "SQL2005" Then
    strFeatures     = strFeatures & ";IIS-HttpRedirect"
    strFeatures     = strFeatures & ";IIS-HttpCompressionDynamic"
    strFeatures     = strFeatures & ";IIS-IIS6ManagementCompatibility"
    strFeatures     = strFeatures & ";IIS-Metabase"
    strFeatures     = strFeatures & ";IIS-WMICompatibility"
    strFeatures     = strFeatures & ";IIS-LegacyScripts"
    If Instr(strOSType, "CORE") = 0 Then
      strFeatures   = strFeatures & ";IIS-LegacySnapIn"
    End If
  End If

  Call SetXMLParm(objInstParm, "ParmXtra",    strFeatures)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("IIS", "PKGMGR.EXE", objInstParm)

End Sub


Sub InstallIISW2012()
  Call DebugLog("InstallIISW2012:")
' Parents can be installed before children
  Dim objInstParm

  strFeatures       = "/online /enable-feature /all "
  strFeatures       = strFeatures & "/featurename:IIS-WebServerRole "
  strFeatures       = strFeatures & "/featurename:WAS-WindowsActivationService "
  strFeatures       = strFeatures & "/featurename:WAS-ConfigurationAPI "
  strFeatures       = strFeatures & "/featurename:WAS-NetFxEnvironment "
  strFeatures       = strFeatures & "/featurename:WAS-ProcessModel "
  strFeatures       = strFeatures & "/featurename:IIS-WebServer "
  strFeatures       = strFeatures & "/featurename:IIS-CommonHttpFeatures "
  strFeatures       = strFeatures & "/featurename:IIS-StaticContent "
  strFeatures       = strFeatures & "/featurename:IIS-DefaultDocument "
  strFeatures       = strFeatures & "/featurename:IIS-DirectoryBrowsing "
  strFeatures       = strFeatures & "/featurename:IIS-HttpErrors "
  strFeatures       = strFeatures & "/featurename:IIS-HealthAndDiagnostics "
  strFeatures       = strFeatures & "/featurename:IIS-HttpLogging "
  strFeatures       = strFeatures & "/featurename:IIS-RequestMonitor "
  strFeatures       = strFeatures & "/featurename:IIS-Performance "
  strFeatures       = strFeatures & "/featurename:IIS-HttpCompressionStatic "
  strFeatures       = strFeatures & "/featurename:IIS-Security "
  strFeatures       = strFeatures & "/featurename:IIS-WindowsAuthentication "
  strFeatures       = strFeatures & "/featurename:IIS-RequestFiltering "
  strFeatures       = strFeatures & "/featurename:IIS-WebServerManagementTools "
  strFeatures       = strFeatures & "/featurename:IIS-ManagementService "
  strFeatures       = strFeatures & "/featurename:IIS-ApplicationDevelopment "
  strFeatures       = strFeatures & "/featurename:IIS-ISAPIExtensions "
  strFeatures       = strFeatures & "/featurename:IIS-ISAPIFilter "
  strFeatures       = strFeatures & "/featurename:IIS-NetFxExtensibility "
  Select Case True
    Case strSQLVersion <= "SQL2016"
      strFeatures   = strFeatures & "/featurename:IIS-ASPNET "
    Case Else
      strFeatures   = strFeatures & "/featurename:IIS-ASPNET45 "
  End Select
  strFeatures       = strFeatures & "/featurename:WCF-HTTP-Activation "
  strFeatures       = strFeatures & "/featurename:WCF-NonHTTP-Activation "

  If strSetupMDS = "YES" Then
    strFeatures     = strFeatures & "/featurename:IIS-HttpCompressionDynamic "
    strFeatures     = strFeatures & "/featurename:IIS-ManagementScriptingTools "
    strFeatures     = strFeatures & "/featurename:IIS-NetFxExtensibility45 "
    strFeatures     = strFeatures & "/featurename:IIS-ASPNET45 "
    strFeatures     = strFeatures & "/featurename:WCF-HTTP-Activation45 "
    strFeatures     = strFeatures & "/featurename:WCF-TCP-Activation45 "
  End If

  Call SetXMLParm(objInstParm, "ParmXtra",    strFeatures)
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("IIS", "DISM.EXE", objInstParm)

End Sub


Sub SetupInstaller45()
  Call SetProcessId("2AH", "Install Windows Installer 4.5")
  Dim objInstParm
  Dim strMSI45File

  Select Case True
    Case (Instr(Ucase(strOSName), " XP") > 0) And (strProcArc  = "X86")   ' Windows XP and 32-bit
      strMSI45File = "INSTMSI45XP.EXE"
    Case strOSVersion < "6.0"                                             ' Windows 2003
      strMSI45File = "INSTMSI45.EXE"
    Case Else                                                             ' Windows 2008 or above
      strMSI45File = "INSTMSI45.MSU"
  End Select

  Select Case True
    Case strSQLVersion = "SQL2008"
      Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & strFileArc & "\redist\Windows Installer\" & strFileArc & "\")
    Case strSQLVersion = "SQL2008R2"
      Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & strFileArc & "\redist\Windows Installer\")
    Case strSQLVersion >= "SQL2012"
      Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & "redist\Windows Installer\")
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot", "YES")
  Call SetXMLParm(objInstParm, "PathMain", strPathAddComp & "redist\Windows Installer\" & strFileArc & "\")
  Call RunInstall("MSI45", strMSI45File, objInstParm)

  Select Case True
    Case strSQLVersion <= "SQL2008"
      ' Nothing
    Case objFSO.GetFileVersion(strPathSys & "msiexec.exe") < "4.5.6001.22159"
      Call SetBuildMessage(strMsgError, "MSI 4.5 install file is missing")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPowerShellV1()
  Call SetProcessId("2AI", "Install PowerShell V1")
  Dim objInstParm
  Dim strPS1File

  If GetPS2Path() <> "" Then
    strSetupPS2     = "YES"
    Call SetBuildfileValue("SetupPS2", strSetupPS2)
    Call SetBuildfileValue("SetupPS1Status", strStatusBypassed & ", PS2 will be installed") 
    Call FBLog(" " & strProcessIdDesc & strStatusBypassed & ", PS2 will be installed")
    Exit Sub
  End If

  strPS1File        = GetBuildfileValue("PS1File")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Select Case True
    Case strPS1File = "PKGMGR"
      strFeatures   = "/IU:MicrosoftWindowsPowerShell"
      If strProcArc <> "X86" Then
        strFeatures = strFeatures & ";MicrosoftWindowsPowerShell-WOW64"
      End If
      Call SetXMLParm(objInstParm, "ParmXtra",    strFeatures)
      Call RunInstall("PS1", "PKGMGR.EXE", objInstParm)
    Case Else
      Call SetXMLParm(objInstParm, "PathMain",    strPathAddComp & "redist\Powershell\" & strFileArc & "\")
      Select Case True
        Case strSQLVersion = "SQL2008"
          Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & strFileArc & "\redist\Powershell\" & strFileArc & "\")
        Case strSQLVersion = "SQL2008R2"
          Call SetXMLParm(objInstParm, "PathAlt", strSQLMedia & "1033_ENU_LP\" & strFileArc & "\redist\Powershell\" & strFileArc & "\")
      End Select
      Call RunInstall("PS1", strPS1File, objInstParm)
  End Select

  Call DebugLog("Check PowerShell version")
  strPath           = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\"
  objWMIReg.GetStringValue strHKLM,strPath,"PowerShellVersion",strVersionPS
  Select Case True
    Case strVersionPS > ""
      Call SetBuildFileValue("SetupPS1Status", strStatusComplete)
    Case strSQLVersion >= "SQL2008"
      Call SetBuildfileValue("SetupPS1Status", strStatusBypassed)
      Call SetBuildMessage(strMsgError, "PowerShell V1 install file not found")
    Case Else       
      Call SetBuildfileValue("SetupPS1Status", strStatusBypassed)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPowerShellV2()
  Call SetProcessId("2AJ", "Install PowerShell V2")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Select Case True
    Case GetPS2Path() = "PKGMGR"
      strFeatures   = "/IU:MicrosoftWindowsPowerShell"
      If strProcArc <> "X86" Then
        strFeatures = strFeatures & ";MicrosoftWindowsPowerShell-WOW64"
      End If
      Call SetXMLParm(objInstParm, "ParmXtra",    strFeatures)
      Call RunInstall("PS2", "PKGMGR.EXE", objInstParm)
    Case Else
      Call RunInstall("PS2", GetBuildfileValue("PS2File"), objInstParm)
  End Select

  Call DebugLog("Check PowerShell version")
  strPath           = "SOFTWARE\Microsoft\PowerShell\1\PowerShellEngine\"
  objWMIReg.GetStringValue strHKLM,strPath,"PowerShellVersion",strVersionPS
  Select Case True
    Case strVersionPS > ""
      Call SetBuildFileValue("SetupPS2Status", strStatusComplete)
    Case strSQLVersion >= "SQL2008"
      Call SetBuildfileValue("SetupPS2Status", strStatusBypassed)
      Call SetBuildMessage(strMsgError, "PowerShell V2 install file not found")
    Case Else       
      Call SetBuildfileValue("SetupPS2Status", strStatusBypassed)
  End Select

  Call SetBuildFileValue("SetupPS2Status", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Function GetPS2Path()
  Call DebugLog("GetPS2Path:")
  Dim strFile, strPath

  strFile           = GetBuildfileValue("PS2File")
  Select Case True
    Case strFile = "PKGMGR"
      strPath       = strFile
    Case Else
      strPath       = GetPathInst(strFile, strPathAddComp, "")
  End Select

  Call DebugLog("PowerShell V2 Install Path: " & strPath)
  GetPS2Path        = strPath

End Function


Sub SetupKB956250()
  Call SetProcessId("2AN", "Install Windows 2008 KB956250 for .Net3 SP1")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\Package_for_KB956250~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.1.6001.18242")
  Call SetXMLParm(objInstParm, "PreConType",  "DWORD")
  Call SetXMLParm(objInstParm, "ParmLog",     "")
  Call RunInstall("KB956250", GetBuildfileValue("KB956250msu"), objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupNet40()
  Call SetProcessId("2AO", "Install .Net 4.0")
  Dim objInstParm
  Dim strDate, strInstFile, strPathLog
 
  Select Case True
    Case Instr(strOSType, "CORE") > 0
      strInstFile   = "dotNetFx40_Full_x86_x64_SC.exe"
    Case Else
      strInstFile   = "dotNetFx40_Full_x86_x64.exe"
  End Select

  Select Case True
    Case strOSVersion >= "6.2"
      strInstFile   = "PKGMGR.EXE"
      Call SetXMLParm(objInstParm, "ParmXtra",    "/IU:NetFX4")
    Case strSQLVersion = "SQL2005"
      Call SetXMLParm(objInstParm, "PathMain",    strSQLMedia & "Servers\redist\DotNetFrameworks\")
      Call SetXMLParm(objInstParm, "PathAlt",     strPathAddComp & "redist\DotNetFrameworks\")
    Case strSQLVersion = "SQL2008"
      Call SetXMLParm(objInstParm, "PathMain",    strSQLMedia & strFileArc & "\redist\DotNetFrameworks\")
      Call SetXMLParm(objInstParm, "PathAlt",     strPathAddComp & "redist\DotNetFrameworks\")
    Case strSQLVersion >= "SQL2008R2"
      Call SetXMLParm(objInstParm, "PathMain",    strSQLMedia & "redist\DotNetFrameworks\")
      Call SetXMLParm(objInstParm, "PathAlt",     strPathAddComp & "redist\DotNetFrameworks\")
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot", "YES")
  Call RunInstall("Net4", strInstFile, objInstParm)

  Select Case True
    Case CheckStatus("Net4")
      Call DebugLog(" Get new .Net V4 status")
      strPath       = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"
      objWMIReg.GetStringValue strHKLM,strPath,"Version",strVersionNet4
      If strOSVersion < "6.2" Then
        Call DebugLog(" Collect .Net 4.0 Install Logs")
        strDate     = Right("0" & DatePart("m", Date()), 2) & "-" & Right("0" & DatePart("d", Date()), 2) & "-" & Right("0" & DatePart("yyyy", Date()), 4)
        strPathLog  = GetPathLog("")
        strPathLog  = Left(strPathLog, Len(strPathLog) - 5)
        strCmd      = "XCOPY """ & strPathTemp & "\*.txt""  " & strPathLog & "\*.*"" /D:" & strDate & " /H /R /Y /Z"
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        strCmd      = "XCOPY """ & strPathTemp & "\*.log""  " & strPathLog & "\*.*"" /D:" & strDate & " /H /R /Y /Z"
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
        strCmd      = "XCOPY """ & strPathTemp & "\*.html"" " & strPathLog & "\*.*"" /D:" & strDate & " /H /R /Y /Z"
        Call Util_RunExec(strCmd, "", strResponseYes, 0)
      End If
      If Left(strOSVersion, 1) < "6" Then
        strReboot   = "Pending"
        Call SetBuildfileValue("RebootStatus", strReboot)
      End If
    Case strSQLVersion >= "SQL2012"
      Call SetBuildMessage(strMsgError, ".Net 4.0 install not successful")
  End Select

  Call ProcessEnd("")

End Sub


Sub SetupNet4x()
  Call SetProcessId("2AQ", "Install .Net 4.x")
  Dim objInstParm
  Dim strNetExe, strNetLevel

  strNetExe         = GetBuildfileValue("Net4Xexe")
  Select Case True
    Case strNetexe = "Unknown"
      strNetLevel   = ""
    Case strNetexe > ""
      strNetLevel   = Mid(strNetexe, 4)
      strNetLevel   = Left(strNetLevel, Instr(strNetLevel, "-") - 1)
    Case Else
      strNetEexe    = "dotnetfx45_full_x86_x64.exe"
      strNetLevel   = ""
  End Select

  Select Case True
    Case strNetLevel > "48"
      strNetLevel   = "999999"
    Case strNetLevel = "48"
      strNetLevel   = "528040"
    Case strNetLevel = "472"
      strNetLevel   = "461808"
    Case strNetLevel = "471"
      strNetLevel   = "461308"
    Case strNetLevel = "47"
      strNetLevel   = "460798"
    Case strNetLevel = "462"
      strNetLevel   = "394802"
    Case strNetLevel = "461"
      strNetLevel   = "394254"
    Case strNetLevel = "46"
      strNetLevel   = "393295"
    Case strNetLevel = "452"
      strNetLevel   = "379893"
    Case strNetLevel = "451"
      strNetLevel   = "378675"
    Case Else
      strNetLevel   = "378389"
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot",   "YES")
  Call SetXMLParm(objInstParm, "PreConKey",   "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\")
  Call SetXMLParm(objInstParm, "PreConValue", strNetLevel)

  If strOSVersion >= "6.3A" Then ' Windows 2016
    Call SetXMLParm(objInstParm, "ParmMonitor",  "12") ' Force reboot if install hangs
  End If

  Call RunInstall("Net4x", strNetExe, objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupDRU()
  Call SetProcessId("2AR", "Setup Distributed Replay")

  Select Case True
    Case strSetupDRUCtlr <> "YES" 
      ' Nothing
    Case Ucase(strCltSvcAccount) <> strNTAuthAccount
      ' Nothing 
    Case Else 
      strCmd        = "NET LOCALGROUP """ & strGroupDistComUsers & """ """ & strCltSvcAccount & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End Select

  Select Case True
    Case strSetupDRUCtlr <> "YES" 
      ' Nothing
    Case Else
      strCmd        = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""SQL DRU Controller"" "
      strCmd        = strCmd & "PROGRAM=""" & strDirProgX86 & "\" & strSQLVersionNum & "\Tools\DReplayController\DReplayController.exe"" "
      strCmd        = strCmd & "ACTION=ALLOW PROFILE=DOMAIN DIR=IN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetBuildFileValue("SetupDRUCtlrStatus", strStatusProgress)
  End Select

  Select Case True
    Case strSetupDRUClt <> "YES"
      ' Nothing
    Case Else
      strCmd        = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""SQL DRU Client"" "
      strCmd        = strCmd & "PROGRAM=""" & strDirProgX86 & "\" & strSQLVersionNum & "\Tools\DReplayClient\DReplayClient.exe"" "
      strCmd        = strCmd & "ACTION=ALLOW PROFILE=DOMAIN DIR=IN ENABLE=YES"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetBuildFileValue("SetupDRUCltStatus", strStatusProgress)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRSAT()
  Call SetProcessId("2AS", "Setup Remote Server Administration Tools (RSAT)")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ASA"
      ' Nothing
    Case Else
      Call InstallRSAT() 
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ASB"
      ' Nothing
    Case GetBuildfileValue("SetupRSATStatus") <>  strStatusProgress
      ' Nothing
    Case Else
'      Call SetupLangenUS() 
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2ASC"
      ' Nothing
    Case GetBuildfileValue("SetupRSATStatus") <>  strStatusProgress
      ' Nothing
    Case Else
      Call EnableRSAT() 
  End Select

  Call SetProcessId("2ASZ", " Setup Remote Server Administration Tools (RSAT)" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub InstallRSAT()
  Call SetProcessId("2ASA", "Install RSAT Components")
  Dim objInstParm
  Dim strRSATFile

  Select Case True
    Case strOSVersion = "6.0"
      strRSATFile   = "Windows6.0-KB941314-"  & strFileArc & "_en-US.msu"
      Call SetXMLParm(objInstParm, "ParmLog",  "")
    Case strOSVersion = "6.1"
      strRSATFile   = "Windows6.1-KB958830-"  & strFileArc & "-RefreshPkg.msu"
    Case strOSVersion = "6.2"
      strRSATFile   = "Windows6.2-KB2693643-" & strFileArc & ".msu"
    Case strOSVersion = "6.3"
      strRSATFile   = "Windows8.1-KB2693643-" & strFileArc & ".msu"
    Case strOSVersion >= "6.3A"
      strRSATFile   = "WindowsTH-KB2693643-"  & strFileArc & ".msu"
  End Select

  Call SetXMLParm(objInstParm, "CleanBoot",    "YES")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("RSAT", strRSATFile, objInstParm)

  Call ProcessEnd("")

End Sub


Sub SetupLangenUS()
  Call SetProcessId("2ASB", "Setup en-US Language")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",     strPathSys)
  Call SetXMLParm(objInstParm, "ParmXtra",     "/i en-US")
  Call SetXMLParm(objInstParm, "ParmLog",      "")
  Call SetXMLParm(objInstParm, "ParmReboot",   "/r")
  Call SetXMLParm(objInstParm, "ParmSilent",   "/s")
  Call SetXMLParm(objInstParm, "StatusOption", strStatusProgress)
  Call RunInstall("RSAT", "LPKSETUP.EXE", objInstParm)

  Call ProcessEnd("")

End Sub


Sub EnableRSAT()
  Call SetProcessId("2ASC", "Enable RSAT")
  Dim objInstParm

  Select Case True
    Case strOSVersion = "6.0"
      Call SetXMLParm(objInstParm, "ParmXtra",    "/IU:RemoteServerAdministrationTools")
      Call RunInstall("RSAT", "PKGMGR.EXE", objInstParm)
    Case Else
      Call SetXMLParm(objInstParm, "ParmXtra",    "/online /enable-feature /featurename:RemoteServerAdministrationTools")
      Call RunInstall("RSAT", "DISM.EXE", objInstParm)
  End Select

  Call ProcessEnd("")

End Sub


Sub SetupPSRemote()
  Call SetProcessId("2AT", "Setup PowerShell Remote Access")

  strCmd            = strCmdPS & " -Command Enable-PSRemoting -Force"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strCmd            = strCmdPS & " Set-Item WSMan:\localhost\Client\TrustedHosts * -Force"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strCmd            = strCmdPS & " Restart-Service -Name WinRM -Force"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strCmd            = strCmdPS & " Set-ExecutionPolicy RemoteSigned"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  strCmd            = strCmdPS & " -ExecutionPolicy Bypass -File """ & strPathFBScripts & "Set-SessionConfig.ps1"" " & strGroupDBA
  Call Util_RunCmdAsync(strCmd, 0)

  If strGroupDBANonSA <> "" Then
    strCmd          = strCmdPS & " -ExecutionPolicy Bypass -File """ & strPathFBScripts & "Set-SessionConfig.ps1"" " & strGroupDBANonSA
    Call Util_RunCmdAsync(strCmd, 0)
  End If

  Call SetBuildFileValue("SetupPSRemoteStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLInstall()
  Call SetProcessId("2B", "Setup SQL Server Install")

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BAZ"
      ' Nothing
    Case Else
      Call SetupSQLEnv()
  End Select 

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BBZ"
      ' Nothing
    Case Else
      Call ProcessSQLInstall("2BB", "DB", strActionSQLDB)
  End Select

  Select Case True ' Install Analysis Services for Cluster
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BCZ"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call ProcessSQLInstall("2BC", "AS", strActionSQLAS)
  End Select

  Select Case True ' Install Reporting Services for Cluster
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BDZ"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call ProcessSQLInstall("2BD", "RS", "INSTALL")
  End Select

  Select Case True ' Install External Services for Cluster
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BEZ"
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call ProcessSQLInstall("2BE", "EX", "INSTALL")
  End Select

  Call SetProcessId("2BZ", " Setup SQL Server Install Process" & strStatusComplete)
  Call ProcessEnd("")

End Sub



Sub SetupSQLEnv()
  Call SetProcessId("2BA", "Setup SQL Install Environment")

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BAA"
      ' Nothing
    Case Else
      Call SetSQLCLSId()
  End Select 

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BAB"
      ' Nothing
    Case Else
      Call SetSQLCompatFlags()
  End Select 

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2BAC"
      ' Nothing
    Case Else
      Call CheckSQLKerberos()
  End Select 

  Call SetProcessId("2BAZ", " Setup SQL Install Environment" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetSQLCLSId()
  Call SetProcessId("2BAA", "Set SQL CLSId Values")

  Select Case True
    Case strOSVersion <= strNativeOS
      ' Nothing
    Case Else
      If strCLSIdSQL > "" Then
        strCmd      = strCompatFlags & "{" & strCLSIdSQL & "}"
        Call Util_RegWrite(strCmd, 4, "REG_DWORD")
      End If
      If strCLSIdSQLSetup > "" Then
        strCmd      = strCompatFlags & "{" & strCLSIdSQLSetup & "}" 
        Call Util_RegWrite(strCmd, 4, "REG_DWORD")
      End If
      If strCLSIdVS > "" Then
        strCmd      = strCompatFlags & "{" & strCLSIdVS & "}"
        Call Util_RegWrite(strCmd, 4, "REG_DWORD")
      End If
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetSQLCompatFlags()
  Call SetProcessId("2BAB", "Set Application Compatibility flags")
  Dim arrProducts
  Dim objProduct
  Dim strPathReg, strProduct

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strClusterAction <> "" ' KB912998 fix for Cluster install
      strCmd        = "HKLM\SYSTEM\CurrentControlSet\Control\Lsa\disabledomaincreds"
      Call Util_RegWrite(strCmd, 0, "REG_DWORD")
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      ' Nothing
    Case Else ' KB2918614 has broken 'minor upgrade' on W2003, workaround is Product must be removed from registry
      strPathReg    = "Installer\Products"
      objWMIReg.EnumKey strHKCR, strPathReg, arrProducts
      For Each objProduct In arrProducts
        strPath     = "HKCR\" & strPathReg & "\" & objProduct
        strProduct  = objShell.RegRead(strPath & "\ProductName")
        If Left(strProduct, 40) = "Microsoft SQL Server Setup Support Files" Then
          strCMD    = "REG DELETE """ & strPath & """ /f"
          Call Util_RunExec(strCmd, "", strResponseYes, 0)
        End If
      Next
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CheckSQLKerberos()
  Call SetProcessId("2BAC", "Check SQL Kerberos Status")
  Dim strServerGroups

  strServerGroups   = GetAccountAttr(strServer, strUserDnsDomain, "memberOf")
  Select Case True
    Case strGroupMSA = ""
      ' Nothing
    Case InStr(" " & strServerGroups & " ", strGroupMSA & " ") = 0
      Call SetBuildfileValue("RebootStatus", "Pending")
      strDebugMsg1  = "Server Groups: " & strServerGroups
      strDebugMsg2  = "MSA Group    : " & strGroupMSA
      Call SetBuildMessage(strMsgError, "Process Kerberos command file to allow SQL install to continue")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub ProcessSQLInstall(strLabel, strSetupType, strSetupAction)
  Call SetProcessId(strLabel, strSQLVersion & " install for build " & strType)
  Dim objInstParm
  Dim strLabelWork, strSetupSQLAS, strSetupSQLDB, strSetupSQLIS, strSetupSQLRS, strSetupSQLTools, strSetupEXFeat, strSetupNCFeat

  Call GetSetup(strSetupType, strSetupAction, objInstParm)
  strSetupSQLAS     = GetXMLParm(objInstParm, "SQLSetupAS",     "NO")
  strSetupSQLDB     = GetXMLParm(objInstParm, "SQLSetupDB",     "NO")
  strSetupSQLIS     = GetXMLParm(objInstParm, "SQLSetupIS",     "NO")
  strSetupSQLRS     = GetXMLParm(objInstParm, "SQLSetupRS",     "NO")
  strSetupSQLTools  = GetXMLParm(objInstParm, "SQLSetupTools",  "NO")
  strSetupEXFeat    = GetXMLParm(objInstParm, "SQLSetupEXFeat", "NO")
  strSetupNCFeat    = GetXMLParm(objInstParm, "SQLSetupNCFeat", "NO")
  strClusterOptions = ""
  strExpressOptions = ""
  strFeatures       = ""
  strListDir        = ""
  strListOpts       = ""
  strListFirstOpts  = ""
  strUpgradeOptions = ""
  objInstParm       = ""

' The following routines must all run with the same ProcessId so that Restart processing can work correctly
  strLabelWork      = strLabel & "A"

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLASParms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLDBParms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLISParms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLRSParms(strSetupSQLDB)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupSQLTools <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLToolsParms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupNCFeat <> "YES"
      ' Nothing
    Case Else
      Call SetupNCFeatures()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupEXFeat <> "YES"
      ' Nothing
    Case Else
      Call SetupEXFeatures()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case Left(strSetupAction, 7) <> "UPGRADE"
      ' Nothing
    Case Else
      Call SetupSQLUpgradeParms()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case Else
      Call SetupSQLCommonParams(strSetupAction, strSetupSQLDB, strSetupSQLAS, strSetupSQLRS)
  End Select 

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strFeatures = ""
      ' Nothing
    Case Else
      Call RunSQLInstall(strLabelWork, strSetupAction)
  End Select

' End of routines that must run with same ProcessId

  strLabelWork      = strLabel & "B"
  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strFeatures = ""
      ' Nothing
    Case Else
      Call SQLLogShortcut(strLabelWork, strInstLog & strProcessIdLabel & " " & strProcessIdDesc)
  End Select

  strLabelWork      = strLabel & "C"
  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strFeatures = ""
      ' Nothing
    Case Else
      Call SaveSQLMenuNames(strLabelWork)
  End Select 

  strLabelWork      = strLabel & "Z"
  Call SetProcessId(strLabelWork, " " & strSQLVersion & " install for build " & strType & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub GetSetup(strSetupType, strSetupAction, objInstParm)
  Call DebugLog("GetSetup: " & strSetupType & ":" & strSetupAction)
  Dim strRSSetup, strSQLDBAction, strSQLDBSetup

  strSQLDBAction    = GetBuildfileValue("ActionSQLDB")
  strSQLDBSetup     = GetBuildfileValue("SetupSQLDB")
  Select Case True
    Case strSetupPowerBI = "YES"
      strRSSetup    = "NO"
    Case strSQLVersion >= "SQL2017"
      strRSSetup    = "NO"
    Case Else
      strRSSetup    = GetBuildfileValue("SetupSQLRS")
  End Select

  Select Case True
    Case strType = "CLIENT"
      Call SetXMLParm(objInstParm,       "SQLSetupTools",  GetBuildfileValue("SetupSQLTools"))
    Case strSetupType = "AS"
      Select Case True
        Case strSQLDBSetup <> "YES"
          Call SetXMLParm(objInstParm,   "SQLSetupAS",     GetBuildfileValue("SetupSQLAS"))
        Case Instr(" " & strActionClusInst & " ADDNODE ", strSQLDBAction & " ") > 0
          Call SetXMLParm(objInstParm,   "SQLSetupAS",     GetBuildfileValue("SetupSQLAS"))
        Case strSetupAction <> strSQLDBAction
          Call SetXMLParm(objInstParm,   "SQLSetupAS",     GetBuildfileValue("SetupSQLAS"))
      End Select
    Case strSetupType = "DB"
      Select Case True
        Case strSQLDBSetup <> "YES"
          ' Nothing
        Case Else
          Call SetXMLParm(objInstParm,   "SQLSetupDB",     GetBuildfileValue("SetupSQLDB"))
          Call SetXMLParm(objInstParm,   "SQLSetupIS",     GetBuildfileValue("SetupSQLIS"))
          Call SetXMLParm(objInstParm,   "SQLSetupTools",  GetBuildfileValue("SetupSQLTools"))
          If Instr(" " & strActionClusInst & " ADDNODE ", strSQLDBAction & " ") = 0 Then
            Call SetXMLParm(objInstParm, "SQLSetupAS",     GetBuildfileValue("SetupSQLAS"))
            Call SetXMLParm(objInstParm, "SQLSetupRS",     strRSSetup)
            Call SetXMLParm(objInstParm, "SQLSetupNCFeat", "YES")
          End If
          If strSQLSharedMR = "YES" Then
            Call SetXMLParm(objInstParm, "SQLSetupEXFeat", "YES")
          End If
      End Select
    Case strSetupType = "RS"
      Select Case True
        Case strSQLDBSetup <> "YES"
          Call SetXMLParm(objInstParm,   "SQLSetupRS",     strRSSetup)
          Call SetXMLParm(objInstParm,   "SQLSetupIS",     GetBuildfileValue("SetupSQLIS"))
          Call SetXMLParm(objInstParm,   "SQLSetupTools",  GetBuildfileValue("SetupSQLTools"))
        Case Instr(" " & strActionClusInst & " ADDNODE ", strSQLDBAction & " ") > 0
          Call SetXMLParm(objInstParm,   "SQLSetupRS",     strRSSetup)
          Call SetXMLParm(objInstParm,   "SQLSetupNCFeat", "YES")
      End Select
    Case strSetupType = "EX"
      Select Case True
        Case strSQLDBSetup <> "YES"
          Call SetXMLParm(objInstParm,   "SQLSetupEXFeat", "YES")
        Case strSQLSharedMR <> "YES"
          Call SetXMLParm(objInstParm,   "SQLSetupEXFeat", "YES")
      End Select
  End Select

End Sub


Sub SetupSQLASParms()
  Call DebugLog("SetupSQLASParms:")

  Select Case True
    Case strSetupSQLASCluster <> "YES"
      ' Nothing
    Case strClusterGroupAS = ""
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      Call MoveToNode(strClusterGroupAS, GetPrimaryNode(strClusterGroupAS))
      Call SetResourceOn(strClusterGroupAS, "GROUP")
      strFailoverClusterIPAddresses = GetClusterIPAddresses(strClusterNameAS, "AS", "ALL")
    Case Else
      strVolList    = "VolDataAS VolLogAS VolTempAS VolBackupAS"
      Call BuildCluster("AS", strClusterGroupAS, "", "", "", "", "", "", strVolList, "")
      Call SetResourceOn(strClusterGroupAS, "GROUP")
      strFailoverClusterIPAddresses = GetClusterIPAddresses(strClusterNameAS, "AS", "ALL")
  End Select

  strFailoverClusterDisks = GetBuildfileValue("FailoverClusterDisks")
  Select Case True
    Case strSetupSQLASCluster <> "YES"
      ' Nothing
    Case strClusterGroupAS = ""
      ' Nothing
    Case (strActionSQLAS = "ADDNODE") And (strSQLVersion = "SQL2005")
      strClusterOptions   = strClusterOptions & " VS=""" & strClusterNameAS & """ InstallVS=""Analysis_Server"" AddNode=""" & strServer & """ Group=""" & strClusterGroupAS & """ "
    Case (strActionSQLAS = "ADDNODE") And (strSQLVersion >= "SQL2012")
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameAS & """ /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
    Case strActionSQLAS = "ADDNODE"
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameAS & """ "
    Case strSQLVersion = "SQL2005"
      strClusterOptions   = strClusterOptions & " VS=""" & strClusterNameAS & """ Group=""" & strClusterGroupAS & """ IP=" & strFailoverClusterIPAddresses & " InstallVS=""Analysis_Server"" AddNode=""" & strServer & """ "
      If strASDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " ASClusterGroup=""" & strASDomainGroup & """"
      End If
    Case strSQLVersion >= "SQL2012"
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameAS & """ /FailoverClusterGroup=""" & strClusterGroupAS & """ /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
      If strFailoverClusterDisks <> "" Then
        strClusterOptions = strClusterOptions & " /FailoverClusterDisks=" & strFailoverClusterDisks
      End If
    Case Else
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameAS & """ /FailoverClusterGroup=""" & strClusterGroupAS & """ /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
      If strASDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " /ASDomainGroup=""" & strASDomainGroup & """"
      End If
      If strFailoverClusterDisks <> "" Then
        strClusterOptions = strClusterOptions & " /FailoverClusterDisks=" & strFailoverClusterDisks
      End If
  End Select

  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " ASACCOUNT="""  & strAsAccount  & """"
      If strAsPassword <> "" Then
        strListOpts = strListOpts & " ASPASSWORD=""" & strAsPassword & """"
      End If
    Case Else
      strListOpts   = strListOpts & " /ASSVCACCOUNT="""  & strAsAccount  & """"
      If strAsPassword <> "" Then
        strListOpts = strListOpts & " /ASSVCPASSWORD=""" & strAsPassword & """"
      End If
  End Select

  If strUserDNSDomain <> "" Then
    strAsSysadmin   = """" & strDomain & "\" & strUserName & """"
  Else
    strAsSysadmin   = """" & strUserName & """"
  End If
  If strSetupStdAccounts = "YES" Then
    strAsSysadmin   = strAsSysadmin & " """ & strGroupDBA & """"
  End If
  If strSSASAdminAccounts <> "" Then
    strAsSysadmin   = strAsSysadmin & " " & strSSASAdminAccounts
  End If
  Select Case True
    Case Left(strAsAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      strAsSysadmin = strAsSysadmin & " """ & strAsAccount & """"
  End Select
  Select Case True
    Case strSqlAccount = ""
      ' Nothing
    Case strSqlAccount = strAsAccount
      ' Nothing
    Case Left(strSqlAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      strAsSysadmin = strAsSysadmin & " """ & strSqlAccount & """"
  End Select
  Select Case True
    Case strAgtAccount = ""
      ' Nothing
    Case strAgtAccount = strAsAccount
      ' Nothing
    Case strAgtAccount = strSqlAccount
      ' Nothing
    Case Left(strAgtAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      strAsSysadmin = strAsSysadmin & " """ & strAgtAccount & """"
  End Select
  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /ASSYSADMINACCOUNTS=" & strAsSysadmin & " "
  End Select

  strCollationAS    = GetBuildfileValue("CollationAS")
  Select Case True
    Case strCollationSQL = ""
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " ASCOLLATION=""" & strCollationAS & """ "
    Case Else
      strlistOpts   = strListOpts & " /ASCOLLATION=""" & strCollationAS & """"
  End Select

  strASProviderMSOlap = GetbuildfileValue("ASProviderMSOlap")
  Select Case True
    Case strSQLVersion = "SQL2005"
      strListDir    = strListDir & " INSTALLASDATADIR=""" & strDirDataAS & """ "
      strListOpts   = strListOpts & " ASAUTOSTART=" & strAsSvcStartuptype & " " 
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case Else
      strListDir    = strListDir & " /ASCONFIGDIR=""" & strDirDataAS & "\Config"" /ASDATADIR=""" & strDirDataAS & """ /ASLOGDIR=""" & strDirLogAS & """ /ASTEMPDIR=""" & strDirTempAS & """ /ASBACKUPDIR=""" & strDirBackupAS & "\AdHocBackup" & """ "
      strListOpts   = strListOpts & " /ASSVCSTARTUPTYPE=" & strASSvcStartupType
      If strASProviderMSOlap <> "" Then
        strListOpts = strListOpts & " /ASPROVIDERMSOLAP=""" & strASProviderMSOlap & """"
      End If
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "Analysis_Server,"
    Case strActionSQLAS = "ADDNODE"
      strFeatures   = strFeatures & "AS,"
    Case strSQLVersion = "SQL2008"
      strFeatures   = strFeatures & "AS,"
    Case strRole = ""
      strFeatures   = strFeatures & "AS,"
    Case Else
      strListOpts   = strListOpts & " /ROLE=""" & strRole & """"
      If UCase(strRole) = UCase("SPI_AS_NewFarm") Then
        strListOpts = strListOpts & " /FARMACCOUNT="""    & strFarmAccount    & """"
        strListOpts = strListOpts & " /FARMPASSWORD="""   & strFarmPassword   & """"
        strListOpts = strListOpts & " /FarmAdminIPort=""" & strFarmAdminIPort & """"
        strListOpts = strListOpts & " /PASSPHRASE="""     & strPassphrase     & """"
      End If
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case strAsServerMode = ""
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /ASSERVERMODE=""" & strAsServerMode & """"
  End Select

End Sub


Sub SetupSQLDBParms()
  Call DebugLog("SetupSQLDBParms:")

  Select Case True
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strClusterGroupSQL = ""
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      If strSQLVersion = "SQL2005" Then
        strClusterOptions = strClusterOptions & " VS=""" & strClusterNameSQL & """ InstallVS=""SQL_Engine"" AddNode=""" & strServer & """ Group=""" & strClusterGroupSQL & """ "
      End If
      Call MoveToNode(strClusterGroupSQL, GetPrimaryNode(strClusterGroupSQL))
      Call SetResourceOn(strClusterGroupSQL, "GROUP")
      strFailoverClusterIPAddresses = GetClusterIPAddresses(strClusterNameSQL, "DB", "ALL")
    Case Else
      strVolList    = "VolSysDB VolData VolLog VolTemp VolLogTemp VolBackup VolDataFT"
      Select Case True
        Case strSetupSQLDBFS <> "YES"
          ' Nothing
        Case strFSLevel < "2"
          ' Nothing
        Case Else
          strVolList    = strVolList & " VolDataFS"
      End Select
      Call BuildCluster("DB", strClusterGroupSQL, "", "", "", "", "", "", strVolList, "")
      Call SetResourceOn(strClusterGroupSQL, "GROUP")
      strFailoverClusterIPAddresses = GetClusterIPAddresses(strClusterNameSQL, "DB", "ALL")
  End Select

  strFailoverClusterDisks = GetBuildfileValue("FailoverClusterDisks")
  Select Case True
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strClusterGroupSQL = ""
      ' Nothing
    Case (strActionSQLDB = "ADDNODE") And (strSQLVersion >= "SQL2012")
      strClusterOptions   = strClusterOptions & " /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strClusterOptions   = strClusterOptions & " VS=""" & strClusterNameSQL & """ Group=""" & strClusterGroupSQL & """ IP=" & strFailoverClusterIPAddresses & " InstallVS=""SQL_Engine"" AddNode=""" & strServer & """ "
      If strAgtDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " AGTClusterGroup=""" & strAgtDomainGroup & """"
      End If
      If strFTSDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " FTSClusterGroup=""" & strFTSDomainGroup & """"
      End If
      If strSQLDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " SQLClusterGroup=""" & strSQLDomainGroup & """"
      End If
    Case strSQLVersion >= "SQL2012"
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameSQL & """ /FailoverClusterGroup=""" & strClusterGroupSQL & """ /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
      If strFailoverClusterDisks <> "" Then
        strClusterOptions = strClusterOptions & " /FailoverClusterDisks=" & strFailoverClusterDisks
      End If
    Case Else
      strClusterOptions   = strClusterOptions & " /FailoverClusterNetworkName=""" & strClusterNameSQL & """ /FailoverClusterGroup=""" & strClusterGroupSQL & """ /FailoverClusterIPAddresses=" & strFailoverClusterIPAddresses & " "
      If strAgtDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " /AgtDomainGroup=""" & strAgtDomainGroup & """"
      End If
      If strSQLDomainGroup <> "" Then
        strClusterOptions = strClusterOptions & " /SQLDomainGroup=""" & strSQLDomainGroup & """"
      End If
      If strFailoverClusterDisks <> "" Then
        strClusterOptions = strClusterOptions & " /FailoverClusterDisks=" & strFailoverClusterDisks
      End If
  End Select

  Select Case True
    Case strSetupDQ <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "DQ,"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strSetupMDS = "YES" 
      strFeatures   = strFeatures & "MDS,"
  End Select

  Select Case True
    Case strSetupPolyBase <> "YES"
      ' Nothing
    Case (Not Checkstatus("JRE")) AND (strSQLVersion < "SQL2019")
      Call SetBuildfileValue("SetupPolyBaseStatus", strStatusBypassed & ", no JRE")
    Case strSetupJRE <> "YES"
      Call SetBuildfileValue("SetupPolyBaseStatus", strStatusBypassed & ", no JRE")
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      If strSQLVersion >= "SQL2019" Then
        strFeatures = strFeatures & "PolyBaseCore,PolyBaseJava,"
      End If
      strFeatures   = strFeatures & "PolyBase,"
      strListOpts = strListOpts & " /PBENGSVCACCOUNT="""  & strPBEngSvcAccount & """ "
      If strPBEngSvcPassword <> "" Then
        strListOpts = strListOpts & " /PBENGSVCPASSWORD=""" & strPBEngSvcPassword & """"
      End If
      strListOpts   = strListOpts & " /PBENGSVCSTARTUPTYPE=" & strPBEngSvcStartup
      strListOpts = strListOpts & " /PBDMSSVCACCOUNT="""  & strPBDMSSvcAccount & """ "
      If strPBDMSSvcPassword <> "" Then
        strListOpts = strListOpts & " /PBDMSSVCPASSWORD=""" & strPBDMSSvcPassword & """"
      End If
      strListOpts   = strListOpts & " /PBDMSSVCSTARTUPTYPE=" & strPBDMSSvcStartup
      strListOpts   = strListOpts & " /PBPORTRANGE=" & strPBPortRange
      strListOpts   = strListOpts & " /PBSCALEOUT=" & strPBScaleout
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "SQL_Engine,SQL_Data_Files,"
      If strSetupSQLDBRepl = "YES" Then
        strFeatures = strFeatures & "SQL_Replication,"
      End If
      If strSetupSQLDBFT = "YES" Then
        strFeatures = strFeatures & "SQL_FullText,"
      End If
      strListDir    = strListDir & "INSTALLSQLDATADIR=""" & strDirSysDB & """ "
    Case strActionSQLDB = "ADDNODE"
      strFeatures   = strFeatures & "SQLEngine,"
    Case Else
      strFeatures   = strFeatures & "SQLEngine,"
      If strSetupSQLDBRepl = "YES" Then
        strFeatures = strFeatures & "Replication,"
      End If
      If strSetupSQLDBFT = "YES" Then
        strFeatures = strFeatures & "FullText,"
      End If
      strListDir    = strListDir & " /INSTALLSQLDATADIR=""" & strDirSysDB & """ /SQLUSERDBDIR=""" & strDirData & """ /SQLUSERDBLOGDIR=""" & strDirLog & """ /SQLTEMPDBDIR=""" & strDirTemp & """ /SQLTEMPDBLOGDIR=""" & strDirLogTemp & """ /SQLBACKUPDIR=""" & strDirBackup & "\AdHocBackup" & """ "
  End Select

  strSecurityMode   = GetBuildfileValue("SecurityMode")
  Select Case True
    Case strSecurityMode = ""
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " SECURITYMODE=""" & strSecurityMode & """"
    Case Else
      strListOpts   = strListOpts & " /SECURITYMODE=""" & strSecurityMode & """"
  End Select

  Select Case True
    Case strsaPwd = ""
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts = strListOpts & " SAPWD=""" & strsaPwd & """"
    Case Else
      strListOpts = strListOpts & " /SAPWD=""" & strsaPwd & """"
  End Select

  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts = strListOpts & " SQLACCOUNT="""  & strSqlAccount  & """ "
      If strSqlPassword <> "" Then
        strListOpts = strListOpts & " SQLPASSWORD=""" & strSqlPassword & """"
      End If
    Case Else
      strListOpts = strListOpts & " /SQLSVCACCOUNT="""  & strSqlAccount & """ "
      If strSqlPassword <> "" Then
        strListOpts = strListOpts & " /SQLSVCPASSWORD=""" & strSqlPassword & """"
      End If
  End Select
 
  Select Case True
    Case strSetupSQLDBCluster = "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " SQLAUTOSTART=" & strSQLSvcStartupType 
    Case Else
      strListOpts   = strListOpts & " /SQLSVCSTARTUPTYPE=" & strSQLSvcStartupType
  End Select

  Select Case True
    Case strSetupSQLDBAG <> "YES"
      ' Nothing
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts = strListOpts & " AGTACCOUNT="""  & strAgtAccount  & """"
      If strAgtPassword <> "" Then
        strListOpts = strListOpts & " AGTPASSWORD=""" & strAgtPassword & """"
      End If
    Case Else
      strListOpts = strListOpts & " /AGTSVCACCOUNT="""  & strAgtAccount & """"
      If strAgtPassword <> "" Then
        strListOpts = strListOpts & " /AGTSVCPASSWORD=""" & strAgtPassword & """"
      End If
  End Select

  Select Case True
    Case strSetupSQLDBAG <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " AGTAUTOSTART=" & strAGTSvcStartupType
    Case Else
      strListOpts   = strListOpts & " /AGTSVCSTARTUPTYPE=" & strAGTSvcStartupType
  End Select

  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strSetupSQLDBFT <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /FTSVCACCOUNT="""  & strFTAccount & """"
      If strFtPassword <> "" Then
        strListOpts   = strListOpts & " /FTSVCPASSWORD=""" & strFTPassword & """"
      End If
  End Select

  If strUserDNSDomain <> "" Then
    strSqlSysadmin  = """" & strDomain & "\" & strUserName & """"
  Else
    strSqlSysadmin  = """" & strUserName & """"
  End If
  If strSetupStdAccounts = "YES" Then
    strSqlSysadmin  = strSqlSysadmin & " """ & strGroupDBA & """"
  End If
  If strSQLAdminAccounts <> "" Then
    strSqlSysadmin  = strSqlSysadmin & " " & strSQLAdminAccounts
  End If
  Select Case True
    Case Left(strActionSQLDB, 7) = "UPGRADE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /SQLSYSADMINACCOUNTS=" & strSqlSysadmin
  End Select

  strCollationSQL   = GetBuildfileValue("CollationSQL")
  Select Case True
    Case strCollationSQL = ""
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " SQLCOLLATION=""" & strCollationSQL & """"
    Case Else
      strlistOpts   = strListOpts & " /SQLCOLLATION=""" & strCollationSQL & """"
  End Select

  Select Case True
    Case strSetupSQLDBFS <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strFSInstLevel <> strFSLevel
      strListOpts   = strListOpts & " /FILESTREAMLEVEL=" & strFSInstLevel
    Case Else
      strListOpts   = strListOpts & " /FILESTREAMLEVEL=" & strFSLevel
      strListOpts   = strListOpts & " /FILESTREAMSHARENAME=""" & strFSShareName & """"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /SQLTEMPDBFILECOUNT=1"
      strListOpts   = strListOpts & " /SQLTEMPDBFILESIZE=" & GetMBSize(strtempdbFile)
      strListOpts   = strListOpts & " /SQLTEMPDBFILEGROWTH=" & GetMBSize(strtempdbFile)
      strListOpts   = strListOpts & " /SQLTEMPDBLOGFILESIZE=" & GetMBSize(strtempdbLogFile)
      strListOpts   = strListOpts & " /SQLTEMPDBLOGFILEGROWTH=" & GetMBSize(strtempdbLogFile)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /SQLSVCINSTANTFILEINIT=true"
  End Select

  Select Case True
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strActionSQLDB <> "ADDNODE"
      ' Nothing
    Case strClusSubnet = "M" 
      strlistOpts   = strListOpts & " /CONFIRMIPDEPENDENCYCHANGE=0 "
  End Select

  Select Case True
    Case strActionSQLDB = strActionClusInst
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("ClusterSQLFound") <> "Y"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /SkipRules=StandaloneInstall_HasClusteredOrPreparedInstanceCheck"
  End Select

End Sub


Function GetMBSize(strSize)
  Call DebugLog("GetMBSize: " & strSize)
  Dim strMBSize, strUnits

  strMBSize         = Replace(strSize, " ", "")
  strUnits          = Right(strMBSize, 2)
  Select Case True
    Case strUnits = "KB"
      strMBSize     = Left(strMBSize, Len(strMBSize) - 2)
      strMBSize     = CStr((CInt(strMBSize) / 1024) + 1)
    Case strUnits = "MB"
      strMBSize     = Left(strMBSize, Len(strMBSize) - 2)
    Case strUnits = "GB"
      strMBSize     = Left(strMBSize, Len(strMBSize) - 2)
      strMBSize     = CStr(CInt(strMBSize) * 1024)
    Case strUnits = "TB"
      strMBSize     = Left(strMBSize, Len(strMBSize) - 2)
      strMBSize     = CStr(CInt(strMBSize) * 1024 * 1024)
    Case Else
      ' Nothing
  End Select

  GetMBSize         = strMBSize

End Function


Sub SetupSQLISParms()
  Call DebugLog("SetupSQLISParms:")

  Select Case True
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "SQL_DTS,"
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "IS,"
      strListOpts   = strListOpts & " /ISSVCSTARTUPTYPE=" & strIsSvcStartupType
      strListOpts   = strListOpts & " /ISSVCACCOUNT="""  & strIsAccount & """"
      If strIsPassword <> "" Then
        strListOpts = strListOpts & " /ISSVCPASSWORD=""" & strIsPassword & """"
      End If
  End Select

  Select Case True
    Case strSetupISMaster <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "IS_Master,"
      strListOpts   = strListOpts & " /ISMasterSVCSTARTUPTYPE=" & strIsMasterStartupType
      strListOpts   = strListOpts & " /ISMasterSVCPORT=" & strIsMasterPort
      Select Case True
        Case strIsMasterThumbprint <> ""
          strListOpts = strListOpts & " /ISMasterSVCTHUMBPRINT=""" & strIsMasterThumbprint & """"
        Case Else
 '         strListOpts = strListOpts & " /ISMasterSVCSSLCertCN=""CN=" & GetBuildfileValue("DNSNameIM") & """
      End Select
      strListOpts   = strListOpts & " /ISMasterSVCACCOUNT="""  & strIsMasterAccount & """"
      If strIsMasterPassword <> "" Then
        strListOpts = strListOpts & " /ISMasterSVCPASSWORD=""" & strIsMasterPassword & """"
      End If
  End Select

  Select Case True
    Case strSetupISWorker <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "IS_Worker,"
      strListOpts   = strListOpts & " /ISWorkerSVCSTARTUPTYPE=" & strIsWorkerStartupType
      If strIsWorkerCert <> "" Then
        strListOpts = strListOpts & " /ISWorkerSVCCERT=""" & strIsWorkerCert & """"
      End If
      If strIsWorkerMaster <> "" Then
        strListOpts = strListOpts & " /ISWorkerSVCMASTER=""" & FormatServer(strIsWorkerMaster, "HTTPS") & ":" & strISMasterPort & """"
      End If
      strListOpts   = strListOpts & " /ISWorkerSVCACCOUNT="""  & strIsWorkerAccount & """"
      If strIsWorkerPassword <> "" Then
        strListOpts = strListOpts & " /ISWorkerSVCPASSWORD=""" & strIsWorkerPassword & """"
      End If
  End Select

End Sub


Sub SetupSQLRSParms(strSetupSQLDB)
  Call DebugLog("SetupSQLRSParms:")

  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts  = strListOpts  & " RSACCOUNT="""  & strRsAccount  & """"
      If strRsPassword <> "" Then
        strListOpts  = strListOpts  & " RSPASSWORD=""" & strRsPassword & """"
      End If
    Case Else
      strListOpts  = strListOpts  & " /RSSVCACCOUNT="""  & strRsAccount & """"
      If strRsPassword <> "" Then
        strListOpts  = strListOpts  & " /RSSVCPASSWORD=""" & strRsPassword & """"
      End If
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2017"
      strFeatures   = strFeatures & "RS,"
    Case strSQLVersion >= "SQL2012"
      strFeatures   = strFeatures & "RS,"
      If strRSShpInstallMode <> "" Then
        strFeatures = strFeatures & "RS_SHP,RS_SHPWFE,"
      End If
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "RS_Server,RS_Web_Interface,RS_SharedTools,"
    Case Else
      strFeatures   = strFeatures & "RS,"
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " RSAUTOSTART=" & strRSSvcStartupType
    Case Else
      strListOpts   = strListOpts & " /RSSVCSTARTUPTYPE=" & strRSSvcStartupType
  End Select

  strRSActualMode   = strRSInstallMode
  Select Case True
    Case strSQLVersion = "SQL2005"
      strRSSQLLocal    = GetBuildfileValue("RSSQLLocal")
      Select Case True
        Case strSetupSQLDB <> "YES"
          strRSActualMode  = "FilesOnly"
          strRSSQLLocal    = 0
        Case strActionSQLRS = "ADDNODE"
          strRSActualMode  = "FilesOnly"
          strRSSQLLocal    = 0
        Case strCatalogServerName = strServer
          ' Nothing
        Case strCatalogServerName = strClusterName
          strRSActualMode  = "FilesOnly"
          strRSSQLLocal    = 0
        Case Else
          strRSActualMode  = "FilesOnly"
          strRSSQLLocal    = 0
      End Select
      strListOpts   = strListOpts & " RSCONFIGURATION=""" & strRSActualMode & """ RSSQLLOCAL=""" & strRSSQLLocal & """"
    Case Else
      Select Case True
        Case strSetupSQLDB <> "YES"
          strRSActualMode  = "FilesOnlyMode"
        Case strActionSQLRS = "ADDNODE"
          strRSActualMode  = "FilesOnlyMode"
        Case strCatalogServerName = strServer
          ' Nothing
        Case strCatalogServerName = strClusterName
          strRSActualMode  = "FilesOnlyMode"
        Case Else
          strRSActualMode  = "FilesOnlyMode"
      End Select
      strListOpts   = strListOpts & " /RSINSTALLMODE=""" & strRSActualMode & """"
      Select Case True
        Case strSQLVersion >= "SQL2017"
          ' Nothing
        Case strSQLVersion >= "SQL2012" 
          strListOpts = strListOpts & " /RSSHPINSTALLMODE= """ & strRSShpInstallMode & """"
      End Select
  End Select
  Call SetBuildfileValue("RSActualMode", strRSActualMode)

End Sub


Sub SetupSQLToolsParms()
  Call DebugLog("SetupSQLToolsParms:")

  Select Case True
    Case strSQLVersion = "SQL2005" 
      strFeatures   = strFeatures & "Connectivity,"
    Case Else
      strFeatures   = strFeatures & "CONN,"
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strEdition = "EXPRESS"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "SQLXML,"
  End Select

  Select Case True
    Case strSetupBOL <> "YES" 
      ' Nothing
    Case strSQLVersion = "SQL2005" And strEdition = "EXPRESS"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "SQL_BooksOnline,"
    Case strSQLVersion <= "SQL2016"
      strFeatures   = strFeatures & "BOL,"
    Case Else
      ' Nothing
  End Select

  Select Case True
    Case strSetupBIDS <> "YES" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "SQL_WarehouseDevWorkbench,"
    Case Else
      strFeatures   = strFeatures & "BIDS," 
  End Select

  Select Case True
    Case strSetupSSMS <> "YES"
      ' Nothing
    Case strUseFreeSSMS = "YES"
      ' Nothing
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strEdition <> "EXPRESS"
      strFeatures = strFeatures & "SQL_Tools90,"
    Case strExpVersion = "Basic"
      ' Nothing
    Case Else
      strFeatures = strFeatures & "SQL_SSMSEE,"
  End Select

  Select Case True
    Case strSetupSSMS <> "YES"
      ' Nothing
    Case strUseFreeSSMS = "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strEdition <> "EXPRESS"
      strFeatures   = strFeatures & "SSMS,ADV_SSMS," 
    Case strExpVersion = "Basic"
      ' Nothing
    Case strExpVersion = "With Tools"
      strFeatures   = strFeatures & "SSMS,"
    Case Else
      strFeatures   = strFeatures & "SSMS,ADV_SSMS,"
  End Select

  Select Case True
    Case strSetupSQLBC <> "YES" 
      ' Nothing
    Case (strSQLVersion = "SQL2005") And (strEdition = "EXPRESS")
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strFeatures   = strFeatures & "Tools_Legacy,"
    Case Else
      ' Nothing
  End Select

  Select Case True
    Case strSetupDQC <> "YES"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "DQC,"
  End Select

End Sub


Sub SetupNCFeatures()
  Call DebugLog("SetupNCFeatures:")
  Dim strRule

  Select Case True
    Case strSetupDRUCtlr <> "YES"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "DREPLAY_CTLR,"
      strListOpts   = strListOpts & " /CTLRSTARTUPTYPE=""" & strCtlrStartupType & """"
      If strUserDNSDomain <> "" Then
        strListOpts = strListOpts & " /CTLRUSERS=" & strUserAccount 
      End If
      strListOpts   = strListOpts & " /CTLRSVCACCOUNT=""" & strCtlrSvcAccount & """"
      If strCtlrPassword <> "" Then
        strListOpts = strListOpts & " /CTLRSVCPASSWORD=""" & strCtlrSvcPassword & """"
      End If
  End Select

  Select Case True
    Case strSetupDRUClt <> "YES"
      ' Nothing
    Case Else
      strFeatures   = strFeatures & "DREPLAY_CLT,"
      strListOpts   = strListOpts & " /CLTSTARTUPTYPE=""" & strCltStartupType & """ /CLTCTLRNAME=""" & strManagementServer & """ /CLTWORKINGDIR=""" & strDirDRU & "\DRU.Work" & """ /CLTRESULTDIR=""" & strDirDRU & "\DRU.Result" & """"
      strListOpts   = strListOpts & " /CLTSVCACCOUNT=""" & strCltSvcAccount & """"
      If strCltPassword <> "" Then
        strListOpts = strListOpts & " /CLTSVCPASSWORD=""" &  strCltPassword & """"
      End If
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strSetupSQLNS = "YES" 
      strFeatures   = strFeatures & "Notification_Services,NS_Engine,NS_Client,"
  End Select

End Sub


Sub SetupExFeatures()
  Call DebugLog("SetupEXFeatures:")
  Dim strRule

  Select Case True
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strSQLVersion >= "SQL2019"
      strFeatures   = strFeatures & "AdvancedAnalytics,"
      strListOpts   = strListOpts & " /MRCACHEDIRECTORY=""" & Left(strPathAddComp, Len(strPathAddComp) - 1) & """ "
      strListOpts   = strListOpts & " /EXTSVCACCOUNT="""  & strExtSvcAccount & """ "
      If strExtSvcPassword <> "" Then
        strListOpts = strListOpts & " /EXTSVCPASSWORD=""" & strExtSvcPassword & """"
      End If
    Case strSetupSQLDBCluster <> "YES"
      strFeatures   = strFeatures & "AdvancedAnalytics,"
      strListOpts   = strListOpts & " /MRCACHEDIRECTORY=""" & Left(strPathAddComp, Len(strPathAddComp) - 1) & """ "
      strListOpts   = strListOpts & " /EXTSVCACCOUNT="""  & strExtSvcAccount & """ "
      If strExtSvcPassword <> "" Then
        strListOpts = strListOpts & " /EXTSVCPASSWORD=""" & strExtSvcPassword & """"
      End If
    Case Else
      strFeatures   = strFeatures & "AdvancedAnalytics,"
      strListOpts   = strListOpts & " /MRCACHEDIRECTORY=""" & Left(strPathAddComp, Len(strPathAddComp) - 1) & """ "
  End Select

  strRule           = "StandaloneInstall_HasClusteredOrPreparedInstanceCheck"
  Select Case True
    Case strSQLVersion >= "SQL2019"
      ' Nothing
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case Instr(strListOpts, strRule) > 0
      ' Nothing
    Case GetBuildfileValue("ClusterSQLFound") = "Y"
      strListOpts   = strListOpts & " /SkipRules=" & strRule
    Case strSetupSQLDBCluster = "YES"
      strListOpts   = strListOpts & " /SkipRules=" & strRule
  End Select

  Select Case True
    Case strSetupPython <> "YES"
      ' Nothing
    Case strSQLSharedMR = "YES"
      strFeatures   = strFeatures & "SQL_SHARED_MPY,"
      strListOpts   = strListOpts & " /IACCEPTPYTHONLICENSETERMS"
'      strListOpts   = strListOpts & " /MPYCACHEDIRECTORY=""" & Left(strPathAddComp, Len(strPathAddComp) - 1) & """ "
    Case Else
      strFeatures   = strFeatures & "SQL_INST_MPY,"
      strListOpts   = strListOpts & " /IACCEPTPYTHONLICENSETERMS"
'      strListOpts   = strListOpts & " /MPYCACHEDIRECTORY=""" & Left(strPathAddComp, Len(strPathAddComp) - 1) & """ "
  End Select

  Select Case True
    Case strSetupRServer <> "YES"
      ' Nothing
    Case strSQLSharedMR = "YES" 
      strFeatures   = strFeatures & "SQL_SHARED_MR,"
      strListOpts   = strListOpts & " /IACCEPTROPENLICENSETERMS"
    Case Else
      If strSQLVersion >= "SQL2017" Then
        strFeatures = strFeatures & "SQL_INST_MR,"
      End If
      strListOpts   = strListOpts & " /IACCEPTROPENLICENSETERMS"
  End Select

  Select Case True
    Case (strSetupAnalytics <> "YES") And (strSetupPython <> "YES") And (strSetupRServer <> "YES")
      ' Nothing
    Case strSQLSharedMR = "YES"
      ' Nothing
    Case Else
      strListOpts = " /INSTANCENAME=""" & GetBuildfileValue("InstMR") & """ " & strListOpts
  End Select

  Select Case True
    Case strSQLVersion < "SQL2019"
      ' Nothing
    Case strSetupJRE <> "YES"
      ' Nothing
    Case Else
      strFeatures = strFeatures & "SQL_INST_JAVA,"  
  End Select

End Sub


Sub SetupSQLUpgradeParms()
  Call DebugLog("SetupSQLUpgradeParms:")

  Select Case True
    Case strFailoverClusterRollOwnership = "" 
      ' Nothing
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /FailoverClusterRollOwnership=""" & strFailoverClusterRollOwnership & """"
  End Select

  Select Case True
    Case strFTUpgradeOption = "" 
      ' Nothing
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /FTUpgradeOption=""" & strFTUpgradeOption & """"
  End Select

  Select Case True
    Case strRSDBAccount = "" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strUpgradeOptions = strUpgradeOptions & " RSUpgradeDatabaseAccount=""" & strRSDBAccount & """"
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /RSUpgradeDatabaseAccount=""" & strRSDBAccount & """"
  End Select

  Select Case True
    Case strAllowUpgradeForRSSharePointMode <> "YES"
      ' Nothing
    Case strSQLVersion <> "SQL2012"
      ' Nothing
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /ALLOWUPGRADEFORRSSHAREPOINTMODE"
  End Select

  Select Case True
    Case strRSDBPassword = "" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strUpgradeOptions = strUpgradeOptions & " RSUpgradePassword=""" & strRSDBPassword & """"
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /RSUpgradePassword=""" & strRSDBPassword & """"
  End Select

  Select Case True
    Case strUseSysDB = "" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strUpgradeOptions = strUpgradeOptions & " UseSysDB=""" & strUseSysDB & """"
    Case Else
      strUpgradeOptions = strUpgradeOptions & " /UseSysDB=""" & strUseSysDB & """"
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case Else
      strCmd        = ""
      If strSetupSQLDB = "YES" Then
        strCmd      = strCmd & "SQL_Engine,"
      End If
      If strSetupSQLAS = "YES" Then
        strCmd      = strCmd & "Analysis_Server,"
      End If
      If strSetupSQLRS = "YES" Then
        strCmd      = strCmd & "RS_Server,"
      End If
      strCmd        = Left(strCmd, Len(strCmd) - 1)
      strUpgradeOptions = strUpgradeOptions & " UPGRADE=""" & strCmd & """"
  End Select

End Sub


Sub SetupSQLCommonParams(strAction, strSetupSQLDB, strSetupSQLAS, strSetupSQLRS)
  Call DebugLog("SetupSQLCommonParams:")
  Dim strPathUpdate

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case Else
      strListOpts   = " /ACTION=""" & strAction & """ " & strListOpts
  End Select

  strPathUpdate       = strUpdateSource
  Select Case True
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strPathUpdate = ""
      Select Case True
        Case strSetupSP = "YES"
          ' Nothing
        Case strSetupSPCU = "YES"
          ' Nothing
        Case Else
          strListOpts = strListOpts & " /UPDATEENABLED=FALSE "
      End Select
    Case Else
      If strPathUpdate = strPathSQLSP Then
        strPathUpdate = strPathUpdate & strSPLevel
      End If
      If Right(strPathUpdate, 1) = "\" Then
        strPathUpdate = Left(strPathUpdate, Len(strPathUpdate) - 1)
      End If
      If objFSO.FolderExists(strPathUpdate) Then
        strListOpts = strListOpts & " /UPDATESOURCE=""" & strPathUpdate & """ /UPDATEENABLED=TRUE "
        Call SetBuildfileValue("SPInclude",   ", included in SQL install")
      End If
   End Select

  If strUserFeatures <> "" Then
    strFeatures     = strFeatures & strUserFeatures & ","
  End If
  If strUserOptions <> "" Then
    strListOpts     = strListOpts & " " & strUserOptions
  End If

  Select Case True
    Case strAction = strActionClusInst
      ' Nothing
    Case strAction = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " DISABLENETWORKPROTOCOLS=" & GetBuildfileValue("DisableNetworkProtocols") & " "
    Case Else
      strListOpts   = strListOpts & " /NPENABLED=" & GetBuildfileValue("NPEnabled") & " /TCPENABLED=" & GetBuildfileValue("TCPEnabled")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strSetupSQLASCluster = "YES" 
      strListOpts   = strListOpts & " ADMINPASSWORD=""" & strAdminPassword & """ "
    Case strSetupSQLDBCluster = "YES" 
      strListOpts   = strListOpts & " ADMINPASSWORD=""" & strAdminPassword & """ "
  End Select
  
  Select Case True
    Case strAction = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " ERRORREPORTING=" & strErrorReporting & " SQMREPORTING=" & strSQMReporting
    Case Else
      strListOpts   = strListOpts & " /ERRORREPORTING=" & strErrorReporting & " /SQMREPORTING=" & strSQMReporting
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strEnu = "YES"
      strListOpts   = strListOpts & " /ENU"
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strAction = "ADDNODE"
      ' Nothing
    Case strSKUUpgrade <> ""
      strListOpts   = strListOpts & " SKUUPGRADE=" & strSKUUpgrade
  End Select

  Select Case True
    Case strPID = ""
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListOpts   = strListOpts & " PIDKEY=""" & Replace(strPID, "-", "") & """"
    Case Else
      strListOpts   = strListOpts & " /PID=""" & strPID & """"
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strWOWX86 <> "TRUE"
      ' Nothing
    Case strFileArc = "X64" 
      strListOpts   = strListOpts & " /X86" 
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008R2"
      ' Nothing
    Case strIAcceptLicenseTerms = "YES"
      strListOpts   = strListOpts & " /IAcceptSQLServerLicenseTerms"
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      Select Case True
        Case strSetupSQLDB <> "YES"
          ' Nothing
        Case Else
          strListOpts = strListOpts & " SQLBROWSERAUTOSTART=" & strSqlBrowserStartup
          strListOpts = strListOpts & " SQLBROWSERACCOUNT="""  & strSqlBrowserAccount  & """"
          If strSqlBrowserPassword <> "" Then
            strListOpts = strListOpts & " SQLBROWSERPASSWORD=""" & strSqlBrowserPassword & """"
          End If
      End Select 
    Case strClusterAction = strActionClusInst
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case Else
      strListOpts = strListOpts & " /BROWSERSVCSTARTUPTYPE=" & strSqlBrowserStartup 
      strListOpts = strListOpts & " /BROWSERSVCUSERNAME=""" & strSqlBrowserAccount  & """"
      If strSqlBrowserPassword <> "" Then
        strListOpts = strListOpts & " /BROWSERSVCPASSWORD=""" & strSqlBrowserPassword & """"
      End If
  End Select

  Select Case True
    Case (strSetupSlipstream <> "YES") And (strSetupSlipstream <> "DONE")
      ' Nothing
    Case strPCUSource = ""
      ' Nothing
    Case strSQLVersion = "SQL2005" ' Syntax is described in KB910070 
      strListFirstOpts = " HOTFIXPATCH=""" & strPCUSource & "hotfixas\files\SQlrun_as.msp;" & strPCUSource & "hotfixsql\files\SQlrun_sql.msp;" & strPCUSource & "hotfixdts\files\SQlrun_dts.msp;" & strPCUSource & "hotfixns\files\SQlrun_ns.msp;" & strPCUSource & "hotfixrs\files\SQlrun_rs.msp;" & strPCUSource & "hotfixtools\files\SQlrun_tools.msp""" & strListFirstOpts
    Case Else
      strListOpts   = strListOpts & " /PCUSOURCE=""" & Left(strPCUSource, Len(strPCUSource) - 1) & """"
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case (strSetupSlipstream <> "YES") And (strSetupSlipstream <> "DONE")
      ' Nothing
    Case strCUSource = ""
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /CUSOURCE=""" & Left(strCUSource, Len(strCUSource) - 1) & """"
  End Select

  Select Case True
    Case strType <> "CLIENT" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case Else
      strListOpts   = strListOpts & " /INSTANCEID=CLIENT"
  End Select

  Select Case True
    Case strSetupSQLDB = "YES"
      If strSQLVersion = "SQL2005" Then
        strListOpts = " INSTANCENAME=""" & strInstance & """ " & strListOpts
      Else
        strListOpts = " /INSTANCENAME=""" & strInstance & """ " & strListOpts
      End If
    Case strSetupSQLAS = "YES"
      If strSQLVersion = "SQL2005" Then
        strListOpts = " INSTANCENAME=""" & strInstASSQL & """ " & strListOpts
      Else
        strListOpts = " /INSTANCENAME=""" & strInstASSQL & """ " & strListOpts
      End If
    Case strSetupSQLRS = "YES"
      If strSQLVersion = "SQL2005" Then
        strListOpts = " INSTANCENAME=""" & strInstRSSQL & """ " & strListOpts
      Else
        strListOpts = " /INSTANCENAME=""" & strInstRSSQL & """ " & strListOpts
      End If
  End Select

  Select Case true
    Case strEdition <> "EXPRESS" 
      ' Nothing
    Case strEnableRANU = "" 
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strExpressOptions = strExpressOptions & " ENABLERANU=" & strEnableRANU
    Case Else
      strExpressOptions = strExpressOptions & " /ENABLERANU=" & strEnableRANU
  End Select

  Select Case True
    Case strAction = ""
      ' Nothing
    Case strAction = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListDir    = strListDir & "INSTALLSQLDIR=""" & strDirProg & """ "  
    Case strWOWX86 = "TRUE"
      strListDir    = strListDir & " /INSTANCEDIR=""" & strDirProgX86 & """ "
    Case Else
      strListDir    = strListDir & " /INSTANCEDIR=""" & strDirProg & """ "
  End Select

  Select Case True
    Case strAction = ""
      ' Nothing
    Case strAction = "ADDNODE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      strListDir    = strListDir & " INSTALLSQLSHAREDDIR=""" & strDirProg & """ "  
    Case strWOWX86 = "TRUE"
      strListDir    = strListDir & " /INSTALLSHAREDDIR=""" & strDirProg & """ "
    Case Else
      strListDir    = strListDir & " /INSTALLSHAREDDIR=""" & strDirProg & """ "
      If strFileArc = "X64" Then
        strListDir  = strListDir & " /INSTALLSHAREDWOWDIR=""" & strDirProgX86 & """ "
      End If
  End Select

  Select Case True
    Case strAction = ""
      ' Nothing
    Case strSQLVersion = "SQL2005"
      ' Nothing
    Case strSQLVersion >= "SQL2012"
      ' Nothing
    Case Left(strOSVersion, 1) <= "5"   ' Installing on W2003 or below
      ' Nothing
    Case Instr(strOSType, "CORE") = 0   ' Not installing on Core OS
      ' Nothing
    Case strClusterReport <> ""         ' Cluster Validation Report found
      ' Nothing
    Case Else
      strClusterOptions = strClusterOptions & " /SkipRules=Cluster_VerifyForErrors"
  End Select

  Select Case True
    Case Else
      If strSetupDRUCtlr = "YES" Then
        strListOpts   = strListOpts & ""
      End If
      If strSetupDRUClt = "YES" Then
        strListOpts   = strListOpts & ""
      End If
  End Select

End Sub


Function GetInstallTimestamp()
  Call DebugLog("GetInstallTimestamp:")
  Dim strTimestamp
' In some error situations, the SQL install does not generate its own timestamp
' When this happens, the install eventually fails giving an error about the missing timestamp and masking the real error
' Therefore a /TIMESTAMP parameter is supplied so that any underlying error is shown

  strTimestamp      = Replace(Replace(Replace(GetStdDateTime(""), "/", ""), " ", "_"), ":", "")
  GetInstallTimestamp = " /TIMESTAMP=" & strTimestamp

End Function


Sub RunSQLInstall(strLabelWork, strAction)
  Call SetProcessId(strLabelWork, "SQL Server Install")
  Dim strSQLCmd, strFeatureParm, strMsgXtra

  strMsgXtra        = "For more advice see https://github.com/SQL-FineBuild/Common/wiki/SQL-Server-Installation-Problems"

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case CheckReboot() = "Pending"
      Call SetupReboot(strLabelWork, "Prepare for SQL Install")
  End Select

  Select Case True 
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > strLabelWork
      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case Else
      Call AlignDTCCluster(strLabelWork)
  End Select 

  Select Case True
    Case (strSQLVersion >= "SQL2014") And (strEdition = "EXPRESS") 
      strCmd        = """" & strSQLMedia & strSQLExe & """ " & "/U /X:""" & strPathTemp & "\" & strSQLVersion & "ExpressMedia"""
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strSQLCmd     = """" & strPathTemp & "\" & strSQLVersion & "ExpressMedia\Setup.exe" & """ "
    Case Else
      strSQLCmd     = """" & strSQLMedia & strSQLExe & """ "
  End Select

  Select Case True
    Case strFeatures = ""
      strFeatureParm = ""
    Case strSQLVersion = "SQL2005"
      strFeatureParm = " ADDLOCAL=""" & Left(strFeatures, Len(strFeatures) - 1) & """"
    Case strAction = "ADDNODE"
      strFeatureParm = ""
    Case Else
      strFeatureParm = " /FEATURES=""" & Left(strFeatures, Len(strFeatures) - 1) & """"
  End Select
 
  strSetupFlag      = ""
  Select Case True
    Case strMode = "ACTIVE"
      strSetupFlag  = ""
    Case strSQLVersion = "SQL2005"
      strSetupFlag  = "/QB "
    Case Else
      strSetupFlag  = "/QUIETSIMPLE "
  End Select

  strCmd            = strListOpts
  Select Case True
    Case Left(strAction, 7) = "UPGRADE"
      strCmd        = strCmd & " " & strUpgradeOptions & " " & strClusterOptions & " " & strExpressOptions
    Case strAction = "INSTALL"
      strCmd        = strFeatureParm & " " & strCmd & " " & strListDir & " " & strExpressOptions
    Case Else
      strCmd        = strFeatureParm & " " & strCmd & " " & strListDir & " " & strClusterOptions
  End Select
  strCmd            = strSQLCmd & strSetupFlag & " " & strListFirstOpts & " " & strCmd & " " & GetInstallTimestamp()

  Call FBLog(" SQL Server install command: " & strCmd)
  Call SetBuildfileValue("RebootLoop", "0")
  Call Util_RunExec(strCmd, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 3010
      strReboot     = "Pending"
      Call SetBuildfileValue("RebootStatus", strReboot)
    Case intErrSave =  -1073741818
      Call FBLog("Retrying SQL Server Install - Scenario Engine workaround")
      Call Util_RunExec("%COMSPEC% /D /C DIR """ & GetBuildfileValue("FBPathLocal") & "\""", "", "", 0) 
      Call Util_RunExec(strCmd & GetInstallTimestamp(), "", strResponseYes, 0) 
    Case (intErrSave = -2054422494) And (strSQLVersion >= "SQL2012") ' Known bug - should not recurr on rerun.
      Call FBLog("Retrying SQL Server Install - /MRCACHEDIRECTORY workaround")
      Call Util_RunExec(strCmd & GetInstallTimestamp(), "", strResponseYes, 0)
    Case intErrSave = -2067529676
      Call SetBuildfileValue("MsgXtra", strMsgXtra)
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " Windows OS software level for SQL Server lower than required for " & strSQLVersion)
    Case intErrSave = -2067723326
      ' Nothing
    Case intErrSave = -2067919934
      Call SetupReboot(strLabelWork, "Retry SQL Install - Pending file updates found")
    Case intErrSave = -2068578304
      Call FBLog("Retrying SQL Server Install - network cannot be reached workaround")
      WScript.Sleep strWaitLong
      Call Util_RunExec(strCmd & GetInstallTimestamp(), "", strResponseYes, 0)
    Case intErrSave = -2068643838 ' No components to install
      ' Nothing
    Case (intErrSave = 5) And (strOSVersion >= "6.1") ' Known bug - should not recurr on rerun
      Call FBLog("Retrying SQL Server Install - Windows handle workaround")
      Call Util_RunExec(strCmd & GetInstallTimestamp(), "", strResponseYes, 0)
    Case intErrSave =  -2067919934
      Call Util_RunExec(strSQLMedia & strSQLSupportMsi & " /norestart /passive", "", strResponseYes, 0) 
      Call FBLog("Retrying SQL Server Install - Fusion ATL workaround")
      Call Util_RunExec(strCmd & GetInstallTimestamp(), "", strResponseYes, 0)     
    Case Else
      Call SetBuildfileValue("MsgXtra", strMsgXtra)
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

  strPathLog        = ""
  strPath           = Mid(strHKLMSQL, 6) & strSQLVersionNum & "\Bootstrap\"
  objWMIReg.GetStringValue strHKLM,strPath,"BootstrapDir",strPathLog
  Select Case True
    Case IsNull(strPathLog)
      Call SetBuildfileValue("MsgXtra", strMsgXtra)
      Call SetBuildMessage(strMsgError, "SQL Server Install could not be run")
    Case strPathLog = ""
      Call SetBuildfileValue("MsgXtra", strMsgXtra)
      Call SetBuildMessage(strMsgError, "SQL Server Install Log can not be found")
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case (strSQLVersion = "SQL2005") Or (strAction = "ADDNODE")
      strPath       = strHKLMSQL & "Instance Names\SQL\"
      strInstRegSQL = objShell.RegRead(strPath & strInstance)
      Call SetBuildfileValue("InstRegSQL",     strInstRegSQL)
      strPath       = strHKLMSQL & strInstRegSQL & "\MSSQLServer\BackupDirectory"
      Call Util_RegWrite(strPath, GetBuildfileValue("DirBackup"), "REG_SZ") 
  End Select

  Select Case True
    Case (strSQLVersion >= "SQL2014") And (strEdition = "EXPRESS") 
      strPath       = strPathTemp & "\" & strSQLVersion & "ExpressMedia"
      strDebugMsg1  = "Deleting: " & strPath
      Set objFolder = objFSO.GetFolder(strPath)
      objFolder.Delete(1)
    Case Else
      ' Nothing
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub AlignDTCCluster(strLabelWork)
  Call DebugLog("Align DTC Cluster for SQL Install")

  Select Case True
    Case strClusterGroupDTC = ""
      ' Nothing
    Case strSQLVersion = "SQL2005"
      Call MoveToNode(strClusterGroupDTC, "")
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case Else
      Call MoveToNode(strClusterGroupDTC, "")
  End Select

End Sub


Sub SQLLogShortcut(strLabelWork, strDescription)
  Call SetProcessId(strLabelWork, "Make shortcut to SQL Setup Log files")
  Dim objLogFolder, objShortcut
  Dim strPathDest, strPathTempLog

  strPathLog        = strPathLog & "Log"
  strDebugMsg1      = "SQL Log Folder " & strPathLog
  Set objFolder     = objFSO.GetFolder(strPathLog)
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
      strPathDest    = strPathLog & "\" & strSQLLog & "\"
      strPathTempLog = strPathTemp & "\SqlSetup.log"
      If objFSO.FileExists(strPathTempLog) Then
        Call objFSO.CopyFile(strPathTempLog, strPathDest, True)
        Call objFSO.DeleteFile(strPathTempLog, True)
      End If
      strPathTempLog = strPathTemp & "\SqlSetup_Local.log"
      If objFSO.FileExists(strPathTempLog) Then
        Call objFSO.CopyFile(strPathTempLog, strPathDest, True)
        Call objFSO.DeleteFile(strPathTempLog, True)
      End If
  End Select

  strPath           = strPathLog & "\" & strSQLLog
  Call FBLog(" '" & strDescription & "' log files located in: " & strPath)
  Call SetBuildfileValue("SQLLog", strPath)
  Set objShortcut   = objShell.CreateShortcut(Mid(strSetupLog, 2) & strDescription & ".lnk")
  objShortcut.TargetPath  = strPath
  objShortcut.Description = strSQLLog
  objShortcut.Save()

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SaveSQLMenuNames(strLabelWork)
  Call SetProcessId(strLabelWork,"Save SQL Server Menu Names")
  Dim colFiles
  Dim strFileName

  strSQLLog         = GetBuildfileValue("SQLLog")
  Set objFolder     = objFSO.GetFolder(strSQLLog)
  Set colFiles      = objFolder.Files

  For Each objFile In colFiles
    strFileName     = RTrim(objFile.Name)
    Select Case True
      Case (StrComp(Right(strFileName, 9), "BOL_1.log") = 0) And (strSQLVersion = "SQL2005")
        Call SaveMenuName(objFile.Path, "MenuSQLDocs",      "DocsMenuFolder")
      Case (StrComp(Left(strFileName, 7), "BOL_loc") = 0) And (strSQLVersion >= "SQL2008")
        Call SaveMenuName(objFile.Path, "MenuSQLDocs",      "DocsMenuFolder")
      Case (StrComp(Right(strFileName, 9), "Tools.log") = 0) And (strSQLVersion = "SQL2005")
        Call SaveMenuName(objFile.Path, "MenuSQL",         "SqlMenuFolder")
        Call SaveMenuName(objFile.Path, "MenuPerfTools",   "PerfToolsMenuFolder")
        Call SaveMenuName(objFile.Path, "MenuConfigTools", "ConfigurationToolsMenuFolder")
      Case (StrComp(Left(strFileName, 13), "sql_tools_loc") = 0) And (strSQLVersion >= "SQL2008")
        Call SaveMenuName(objFile.Path, "MenuPerfTools",   "PerfToolsMenuFolder")
      Case (StrComp(Left(strFileName, 11), "SqlSupport_") = 0) And (StrComp(Right(strFileName, 19), "ComponentUpdate.log") = 0) And (strSQLVersion >= "SQL2008")
        Call SaveMenuName(objFile.Path, "MenuSQL",         "SqlMenuFolder")
        Call SaveMenuName(objFile.Path, "MenuConfigTools", "ConfigurationToolsMenuFolder")
      Case (StrComp(Left(strFileName, 11), "SqlSupport_") = 0) And (StrComp(Right(strFileName, 6), "_1.log") = 0) And (Instr(strFileName, "Katmai") = 0) And (strSQLVersion >= "SQL2014")
        Call SaveMenuName(objFile.Path, "MenuSQL",         "SqlMenuFolder")
        Call SaveMenuName(objFile.Path, "MenuConfigTools", "ConfigurationToolsMenuFolder")
      Case (StrComp(Left(strFileName, 16), "sql_ssms_loc_Cpu") = 0) And (strSQLVersion >= "SQL2008") And (strSPLevel < "RTM") 
        Call SaveMenuString(objFile.Path, "MenuSSMS",        "ProductName = ")
      Case (StrComp(Left(strFileName, 14), "sql_as_loc_Cpu") = 0) And (strSQLVersion >= "SQL2008") 
        Call SaveMenuString(objFile.Path, "DirASDLL",      "SQLSETUPARPWRAPPER = ")
    End Select
  Next

  strMenuSQL          = GetBuildfileValue("MenuSQL")
  Call SetBuildfileValue("Menu" & strSQLVersion, strMenuSQL)

  Set colFiles      = Nothing
  Set objFolder     = Nothing
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SaveMenuName(strPath, strMenuName, strMenuKey)
  Call DebugLog("SaveMenuName: " & strMenuName & " for " & strPath)
  Dim intIdx
  Dim objTextFile
  Dim strText, strMenuPath
  
  Set objTextFile   = objFSO.OpenTextFile(strPath, 1, 0, True)
  strText           = Replace(objTextFile.ReadAll, vbCr, "")
  intIdx            = Instr(strText, "Adding " & strMenuKey)
  If intIdx > 0 Then
    strMenuPath     = Mid(strText, intIdx)
    intIdx          = Instr(strMenuPath, vbLf)
    strMenuPath     = Left(strMenuPath, intIdx - 1)
    intIdx          = Instr(strMenuPath, "'")
    strMenuPath     = Mid(strMenuPath, intIdx + 1)
    intIdx          = Instr(strMenuPath, "'")
    strMenuPath     = Left(strMenuPath, intIdx - 1)
    If Right(strMenuPath, 1) = "\" Then
      strMenuPath   = Left(strMenuPath, Len(strMenuPath) - 1)
    End If
    intIdx          = InstrRev(strMenuPath, "\")
    strMenuPath     = Mid(strMenuPath, intIdx + 1)
    If strMenuPath > "" Then
      Call SetBuildfileValue(strMenuName, strMenuPath)
    End If
  End If

  objTextFile.Close
  Set objTextFile   = Nothing

End Sub


Sub SaveMenuString(strPath, strMenuName, strMenuKey)
  Call DebugLog("SaveMenuString: " & strMenuName & " for " & strPath)
  Dim intIdx
  Dim objTextFile
  Dim strText, strMenuPath
  
  Set objTextFile   = objFSO.OpenTextFile(strPath, 1, 0, True)
  strText           = Replace(objTextFile.ReadAll, vbCr, "")
  intIdx            = Instr(strText, strMenuKey)
  If intIdx > 0 Then
    intIdx          = intIdx + Len(strMenuKey)
    strMenuPath     = Mid(strText, intIdx)
    intIdx          = Instr(strMenuPath, vbLf)
    strMenuPath     = Left(strMenuPath, intIdx - 1)
    If Right(strMenuPath, 4) = ".exe" Then
      intIdx        = InstrRev(strMenuPath, "\")
      strMenuPath   = Left(strMenuPath, intIdx)
    End If
    If Right(strMenuPath, 1) = "\" Then
      strMenuPath   = Left(strMenuPath, Len(strMenuPath) - 1)
    End If
    intIdx          = InstrRev(strMenuPath, "\")
    If intIdx > 0 Then
      strMenuPath   = Mid(strMenuPath, intIdx + 1)
    End If
    If strMenuPath > "" Then
      Call SetBuildfileValue(strMenuName, strMenuPath)
    End If
  End If

  objTextFile.Close
  Set objTextFile   = Nothing

End Sub


Sub SetupPostSQLTasks()
  Call SetProcessId("2C", "SQL Server post-install tasks")
  
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAZ"
      ' Nothing
    Case Else
      Call SQLPostInstall()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CBZ"
      ' Nothing
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case Else
      Call SetupAnalytics()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCZ"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case Else
      Call CheckChildClusters()
  End Select

  Call SetProcessId("2CZ", " SQL Server post-install tasks" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SQLPostInstall()
  Call SetProcessId("2CA","Perform SQL Post-Install tasks")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAA"
      ' Nothing
    Case Else
      Call CheckSQLComponents()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAB"
      ' Nothing
    Case Else
      Call SaveEditionData()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAC"
      ' Nothing
    Case Else
      Call RefreshEditionData()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAD"
      ' Nothing
    Case Else
      Call SaveSQLRegistryPaths()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAE"
      ' Nothing
    Case strSetupSQLDBFS <> "YES"

      ' Nothing

    Case strFSInstLevel = strFSLevel
      Call SetBuildfileValue("SetupSQLDBFSStatus", strStatusComplete)
    Case Else
      Call EnableFilestream()
  End Select



  Select Case True

    Case err.Number <> 0

      ' Nothing

    Case strProcessId > "2CAFZ"

      ' Nothing
    Case GetBuildfileValue("SetupClusterShares") <> "YES"
      ' Nothing

    Case strClusterAction = ""

      ' Nothing

    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case Else

      Call SetupClusterShares()

  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAG"
      ' Nothing
    Case Else
      Call SetupFilePermissions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAH"
      ' Nothing
    Case Else
      Call SetupDBARegistryPermissions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAI"
      ' Nothing
    Case Else
      Call SetupWMIPermissions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAJ"
      ' Nothing
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case Else
      Call SetupServiceDependencies()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAK"
      ' Nothing
    Case strSetupSQLDBAG <> "YES"
      ' Nothing
    Case Else
      Call CheckSQLAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAL"
      ' Nothing
    Case Else
      Call CheckSQLBrowser()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAMZ"
      ' Nothing
    Case strClusterHost <> "YES"
      ' Nothing
    Case Else
      Call SetupClusterBindings()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAN"
      ' Nothing
   Case (strManagementServerName <> strClusterNameSQL) And (strClusterNameSQL <> "")
      ' Nothing
    Case (strManagementServerName <> strServer) And (strClusterNameSQL = "")
      ' Nothing
    Case strManagementInstance <> strInstance
      ' Nothing
    Case Else
      Call RegisterManagementServer()
  End Select

  Call SetProcessId("2CAZ", " Check SQL Server Edition data" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CheckSQLComponents()
  Call SetProcessId("2CAA","Check required SQL components were installed")
  Dim strServiceName

  If strSetupSQLTools = "YES" Then
    Call SetBuildfileValue("SetupSQLToolsStatus", strStatusProgress)
  End If

  Call DebugLog("Checking Service: " & strInstAS)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstAS & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupSQLASStatus", strStatusComplete)
    Case Else
      Call SetBuildMessage(strMsgError, "SQL Analysis Services was not installed")
  End Select 

  Call DebugLog("Checking Service: " & strInstSQL)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstSQL & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupSQLDBStatus", strStatusComplete)
      If strSetupSQLDBRepl = "YES" Then
        Call SetBuildfileValue("SetupSQLDBReplStatus", strStatusComplete)
      End If
      If strSetupSQLDBFT = "YES" Then
        Call SetBuildfileValue("SetupSQLDBFTStatus",   strStatusProgress)
      End If
    Case Else
      Call SetBuildMessage(strMsgError, "SQL Database Services was not installed")
  End Select  
    
  Call DebugLog("Checking Service: " & strInstAgent)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstAgent & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBAG <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupSQLDBAGStatus", strStatusComplete)
    Case Else
      Call SetBuildMessage(strMsgError, "SQL Agent was not installed")
  End Select 

  Call DebugLog("Checking Service: " & strInstIS)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstIS & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupSQLISStatus", strStatusComplete)
    Case Else
      Call SetBuildMessage(strMsgError, "SQL Integration Services was not installed")
  End Select 

  Call DebugLog("Checking Service: " & strInstAnal)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstAnal & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupAnalyticsStatus", strStatusProgress)
    Case Else
      Call SetBuildfileValue("SetupAnalyticsStatus", strStatusBypassed)
  End Select

  strServiceName    = "SQL Server Distributed Replay Controller"
  Call DebugLog("Checking Service: " & strServiceName)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strServiceName & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupDRUCtlr <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupDRUCtlrStatus", strStatusComplete)
    Case Else
      Call SetBuildMessage(strMsgError, "DRU Controller was not installed")
  End Select 

  strServiceName    = "SQL Server Distributed Replay Client"
  Call DebugLog("Checking Service: " & strServiceName)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strServiceName & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupDRUClt <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupDRUCltStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupDRUCltStatus", strStatusBypassed)
  End Select 

  Call DebugLog("Checking Service: " & strInstPE)
  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstPE & "\"
  strPathNew        = ""
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strPathNew
  Select Case True
    Case strSetupPolyBase <> "YES"
      ' Nothing
    Case strPathNew > ""
      Call SetBuildfileValue("SetupPolyBaseStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupPolyBaseStatus", strStatusBypassed)
  End Select

  strPathNew        = strPathVS & "IDE\devenv.exe"
  Select Case True
    Case strSetupBIDS <> "YES"
      ' Nothing
    Case objFSO.FileExists(strPathNew)
      Call SetBuildfileValue("SetupBIDSStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupBIDSStatus", strStatusBypassed)
  End Select

  strPathNew        = strDirProgX86 & "\" & strSQLVersionNum & "\Tools\Binn\DQ\DataQualityServices.exe"
  Select Case True
    Case strSetupDQC <> "YES"
      ' Nothing
    Case objFSO.FileExists(strPathNew)
      Call SetBuildfileValue("SetupDQCStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupDQCStatus", strStatusBypassed)
  End Select

  strPathNew        = strDirProgX86 & "\80\Tools\binn\sqldmo.dll"
  Select Case True
    Case strSetupSQLBC <> "YES"
      ' Nothing
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case objFSO.FileExists(strPathNew)
      Call SetBuildfileValue("SetupSQLBCStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupSQLBCStatus", strStatusBypassed)
  End Select

  strPathNew        = strDirProgX86 & "\90\NotificationServices\9.0.242\Bin\nscontrol.exe"
  Select Case True
    Case strSetupSQLNS <> "YES"
      ' Nothing
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case objFSO.FileExists(strPathNew)
      Call SetBuildfileValue("SetupSQLNSStatus", strStatusComplete)
    Case Else
      Call SetBuildfileValue("SetupSQLNSStatus", strStatusBypassed)
  End Select 

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SaveEditionData()
  Call SetProcessId("2CAB","Save SQL Edition Data")
  Dim strEditionOrig

  Select Case True
    Case strType = "CLIENT"
      strInstRegSQL = strSQLVersionNum & "\Tools"
      Call SetBuildfileValue("InstRegSQL",     strInstRegSQL)
      strSQLBinRoot = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLPath")
      Call SetBuildfileValue("SQLBinRoot",     strSQLBinRoot)
      strSQLVersionFull = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\Version")
      Call SetBuildfileValue("SQLVersionFull", strSQLVersionFull)
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call DebugLog("Save installed Edition name")
      strPath       = strHKLMSQL & "Instance Names\SQL\"
      strInstRegSQL = objShell.RegRead(strPath & strInstance)
      Call SetBuildfileValue("InstRegSQL",     strInstRegSQL)
      strSQLBinRoot = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLBinRoot")
      Call SetBuildfileValue("SQLBinRoot",     strSQLBinRoot)
      strSQLVersionFull = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\Version")
      Call SetBuildfileValue("SQLVersionFull", strSQLVersionFull)
      strEdition    = objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\Edition")
      strEdition    = Ucase(Left(strEdition, InStrRev(strEdition, "Edition") - 2))
      Call SetBuildfileValue("AuditEdition",   strEdition)
      strEditionOrig  = GetBuildfileValue("EditionOrig")
      If strEditionOrig <> strEdition Then
        Call SetBuildMessage(strMsgWarning, "Installed Edition " & strEdition & " does not match requested Edition " & strEditionOrig)
      End If
      If strEdType = strStatusAssumed Then
        Call SetBuildfileValue("EdType",       "")
      End If
  End Select

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      strPath       = strHKLMSQL & "Instance Names\OLAP\"
      objWMIReg.GetStringValue strHKLM,Mid(strPath,6),strInstASSQL,strInstRegAS
      Call SetBuildfileValue("InstRegAS",      strInstRegAS)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub

Sub RefreshEditionData()
  Call SetProcessId("2CAC","Refresh Buildfile values for updated Edition")

  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupMDS       = GetBuildfileValue("SetupMDS")
  strSetupSQLNS     = GetBuildfileValue("SetupSQLNS")
  strSetupStreamInsight = GetBuildfileValue("SetupStreamInsight")

  Select Case True
    Case strEdition = "WEB"
      Call SetParam("SetupSQLAS",         strSetupSQLAS,         "NO",  "Analysis Services can not be installed for "     & strEdition & "Edition", "")
      Call SetParam("SetupSQLIS",         strSetupSQLIS,         "NO",  "Integration Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupMDS",           strSetupMDS,           "NO",  "Master Data Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupSQLNS",         strSetupSQLNS,         "NO",  "Notification Services can not be installed for " & strEdition & "Edition", "")
    Case strEdition = "WORKGROUP"
      Call SetParam("SetupSQLAS",         strSetupSQLAS,         "NO",  "Analysis Services can not be installed for "     & strEdition & "Edition", "")
      Call SetParam("SetupSQLIS",         strSetupSQLIS,         "NO",  "Integration Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupMDS",           strSetupMDS,           "NO",  "Master Data Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupSQLNS",         strSetupSQLNS,         "NO",  "Notification Services can not be installed for " & strEdition & "Edition", "")
    Case strEdition = "EXPRESS"
      Call SetParam("SetupSQLAS",         strSetupSQLAS,         "NO",  "Analysis Services can not be installed for "     & strEdition & "Edition", "")
      Call SetParam("SetupSQLDBAG",       strSetupSQLDBAG,       "NO",  "SQL Agent can not be installed for "             & strEdition & "Edition", "")
      Call SetParam("SetupSQLIS",         strSetupSQLIS,         "NO",  "Integration Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupMDS",           strSetupMDS,           "NO",  "Master Data Services can not be installed for "  & strEdition & "Edition", "")
      Call SetParam("SetupSQLNS",         strSetupSQLNS,         "NO",  "Notification Services can not be installed for " & strEdition & "Edition", "")
      Call SetParam("SetupStreamInsight", strSetupStreamInsight, "NO",  "Stream Insight can not be installed for "        & strEdition & "Edition", "")
  End Select

  Call SetBuildfileValue("SetupSQLAS",         strSetupSQLAS)
  Call SetBuildfileValue("SetupSQLDBAG",       strSetupSQLDBAG)
  Call SetBuildfileValue("SetupSQLIS",         strSetupSQLIS)
  Call SetBuildfileValue("SetupMDS",           strSetupMDS)
  Call SetBuildfileValue("SetupSQLNS",         strSetupSQLNS)
  Call SetBuildfileValue("SetupStreamInsight", strSetupStreamInsight)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SaveSQLRegistryPaths()
  Call SetProcessId("2CAD","Save Paths to SQL Registry Items")

  strCmdSQL         = GetCmdSQL()

  Select Case True
    Case strSetupSSMS <> "YES"
      ' Nothing
    Case strUseFreeSSMS = "YES"
      ' Nothing
    Case Else
      strPath       = GetBuildfileValue("RegTools") & "SQLPath"
      strPathSSMS   = objShell.RegRead(strPath)
      If strPathSSMS > "" Then
        strPathSSMS = strPathSSMS & "\"
        Call SetBuildfileValue("PathSSMS", strPathSSMS)
        Call SetBuildfileValue("SetupSSMSStatus", strStatusComplete)
      End If
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub EnableFilestream()
  Call SetProcessId("2CAE", "Enable Filstream")
  ' Code based on https://sqlsrvengine.codeplex.com/wikipage?title=FileStreamEnable
  Dim objSQLConfig, objInParam, objOutParam
  Dim strLocalFSLevel

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'show advanced options',           '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)
  Wscript.Sleep strWaitShort

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'remote access',                   '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)
  Wscript.Sleep strWaitShort

  strPath           = "winmgmts:{impersonationLevel=impersonate}!\\" & strServer & "\ROOT\Microsoft\SqlServer\ComputerManagement" & Left(strSQLVersionNum, 2)
  Set objSQLConfig  = GetObject(strPath & ":FilestreamSettings='" & strInstance & "'")
  Set objInParam    = objSQLConfig.Methods_("EnableFilestream").inParameters.SpawnInstance_()
  objInParam.AccessLevel = strFSLevel
  objInParam.ShareName   = strFSShareName
  Set objOutParam   = objSQLConfig.ExecMethod_("EnableFilestream", objInParam)
  If objOutParam.returnValue <> 0 Then
    Call SetBuildMessage(strMsgWarning, "Unable to change SQL Server Configuration to enable Filestream")
    Call SetBuildfileValue("SetupSQLDBFSStatus", strStatusBypassed)
    Exit Sub
  End If 

  strLocalFSLevel   = CStr(CInt(strFSLevel) - 1)
  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'filestream_access_level','" & strLocalFSLevel & "';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  Call SetBuildfileValue("SetupSQLDBFSStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupClusterShares()
  Call SetProcessId("2CAF", "Setup Cluster Shares")

  Select Case True
    Case strProcessId > "2CAFA"

      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case Else
      Call SetupSOFSRole
  End Select

  Select Case True
    Case strProcessId > "2CAFB"

      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLDBShares
  End Select

  Select Case True
    Case strProcessId > "2CAFC"

      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLASShares
  End Select

  Select Case True
    Case strProcessId > "2CAFD"

      ' Nothing
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case Else
      Call SetupDTCShares
  End Select

  Call SetBuildfileValue("SetupSharesStatus", strStatusComplete)
  Call SetProcessId("2CAFZ", " Setup Cluster Shares" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSOFSRole()
  Call SetProcessId("2CAFA", "Setup SOFS Cluster Role " & strClusterGroupFS)

  strCmd            = strCmdPS & " Add-ClusterScaleOutFileServerRole -Name '" & strClusterGroupFS & "' -Cluster '" & strClusterName & "'"
  Call Util_RunExec(strCmd, "", strResponseYes, 5010)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLDBShares()
  Call SetProcessId("2CAFB", "Setup SQLDB Shares " & strClusterGroupSQL)

  Call PrepareShare("VolData",       strVolData,     strClusterGroupSQL, strClusterNameSQL, "DirData")
  Call PrepareShare("VolLog",        strVolLog,      strClusterGroupSQL, strClusterNameSQL, "DirLog")
  Call PrepareShare("VolSysDB",      strVolSysDB,    strClusterGroupSQL, strClusterNameSQL, "DirSysDB")
  Call PrepareShare("VolTemp",       strVolTemp,     strClusterGroupSQL, strClusterNameSQL, "DirTemp")
  Call PrepareShare("VolLogTemp",    strVolLogTemp,  strClusterGroupSQL, strClusterNameSQL, "DirLogTemp")
  Call PrepareShare("VolBackup",     strVolBackup,   strClusterGroupSQL, strClusterNameSQL, "DirBackup")

  Select Case True
    Case strSetupBPE <> "YES"
      ' Nothing
    Case Else
      Call PrepareShare("VolBPE",    strVolBPE,      strClusterGroupSQL, strClusterNameSQL, "DirBPE")
  End Select

  Select Case True
    Case strSetupSQLDBFS <> "YES"
      ' Nothing
    Case GetBuildfileValue("SetupSQLDBFSStatus") = strStatusBypassed
      ' Nothing
    Case strFSLevel < "2"
      ' Nothing
    Case Else
      Call PrepareShare("VolDataFS", strVolDataFS,   strClusterGroupSQL, strClusterNameSQL, "DirDataFS")
  End Select

  Select Case True
    Case strSetupSQLDBFT <> "YES"
      ' Nothing
    Case Else
      Call PrepareShare("VolDataFT", strVolDataFT,   strClusterGroupSQL, strClusterNameSQL, "DirDataFT")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLASShares()
  Call SetProcessId("2CAFC", "Setup SQLAS Shares " & strClusterGroupAS)

  Call PrepareShare("VolDataAS",     strVolDataAS,   strClusterGroupAS, strClusterNameAS, "DirDataFS")
  Call PrepareShare("VolLogAS",      strVolLogAS,    strClusterGroupAS, strClusterNameAS, "DirLogAS")
  Call PrepareShare("VolTempAS",     strVolTempAS,   strClusterGroupAS, strClusterNameAS, "DirTempAS")
  Call PrepareShare("VolBackupAS",   strVolBackupAS, strClusterGroupAS, strClusterNameAS, "DirBackupAS")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDTCShares()
  Call SetProcessId("2CAFD", "Setup DTC Shares " & strClusterGroupDTC)

  Call PrepareShare("VolDTC",        strVolDTC,      strClusterGroupDTC, strClusterName, "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub PrepareShare(strVolVar, strVolList, strClusterGroup, strClusterName, strDirectory)
  Call DebugLog("PrepareShare: " & strVolList)
  Dim arrVolumes
  Dim intIdx, intVol
  Dim strVol, strVolLabel, strVolSource, strVolumes, strVolType, strShareName

  strVolSource      = GetBuildfileValue(strVolVar & "Source")
  arrVolumes        = Split(strVolList, ",")

  Select Case True
    Case strVolSource = "D" 
      For intVol = 0 To UBound(arrVolumes)
        strVolumes     = Trim(arrVolumes(intVol))
        For intIdx = 1 To Len(strVolumes)
          strVol       = Mid(strVolumes, intIdx, 1)
          strVolLabel  = GetBuildfileValue("Vol" & strVol & "Label")
          strVolType   = GetBuildfileValue("Vol" & strVol & "Type")
          strShareName = GetBuildfileValue("Vol" & strVol & "Share")
          If strVolType = "C" Then
            Call SetupClusterShare(strVol, strVolLabel, strShareName, strClusterGroup, strClusterName)
          End If
        Next
      Next
    Case (strVolSource = "C") And (strDirectory <> "")
      strVol           = GetBuildfileValue(strDirectory)
      strVolLabel      = GetBuildfileValue("Lab" & Mid(strDirectory, 4))
      strShareName     = Mid(strDirectory, 4)
      Call SetupClusterShare(strVol, strVolLabel, strShareName, strClusterGroup, strClusterName) ' this needs troubleshooting
  End Select

End Sub


Sub SetupClusterShare(strVol, strVolLabel, strShareName, strClusterGroup, strClusterName)
  Call DebugLog("SetupClusterShare: " & strVol & " for " & strShareName)
  Dim strVolNew

  strVolNew         = strVol
  Select Case True
    Case strOSVersion >= "6.2"
      If Len(strVolNew) = 1 Then
        strVolNew = strVolNew & ":\"
      End If
      strCmd        = strCmdPS & " New-SmbShare -Name '" & strShareName & "' -Path '" & strVolNew & "' -ScopeName '" & strClusterGroupFS & "' -Description '" & strShareName & " share' -FolderEnumerationMode AccessBased -FullAccess 'Administrators' -ReadAccess 'Users'"
      Call Util_RunExec(strCmd, "", strResponseYes, 1)
    Case Left(strOSVersion, 1) >= "6"
      If Len(strVolNew) = 1 Then
        strVolNew = strVolNew & ":"
      End If
      strCmd        = "NET SHARE """ & strShareName & """=""" & strVolNew & """ /Grant:Administrators,FULL /GRANT:Users,READ"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
    Case Else
      If Len(strVolNew) = 1 Then
        strVolNew = strVolNew & ":\"
      End If
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strShareName  & """ /CREATE /GROUP:""" & strClusterGroup & """ /TYPE:""File Share"" /PROP DESCRIPTION=""" & strShareName & " share"""
      Call Util_RunExec(strCmd, "", strResponseYes, 5010)
      Call SetResourceOff(strShareName, "") ' Ensure Resource is offline in case it already exists
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strShareName  & """ /ADDDEP:""" & strVolLabel & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 5003)
      Call SetResourceOn(strVolLabel, "")
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strShareName  & """ /PRIV SHARENAME=""" & strShareName & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 5010)
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strShareName  & """ /PRIV PATH=""" & strVolNew & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 5010)
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strShareName  & """ /PRIV SECURITY=Administrators,set,F,Users,set,R:security"
      Call Util_RunExec(strCmd, "", strResponseYes, 5010)
      Call SetResourceOff(strShareName, "")
  End Select

End Sub


Sub SetupFilePermissions()
  Call SetProcessId("2CAG", "Reset File Permissions")

  Select Case True
    Case strSetupSQLAS <> "YES" 
      ' Nothing
    Case Else
      Call DebugLog("Reset permissions on AS folders")
      Call ResetDBAFilePerm(strDirProg)
      Call ResetDBAFilePerm(objShell.RegRead(strHKLMSQL & strInstRegAS & "\Setup\SQLPath"))
      Call ResetFilePerm(strDirProg, strASAccount)
      strPath       = strWinDir & "\System32\LogFiles\Sum"
      If objFSO.FolderExists(strPath) Then
        Call ResetFilePerm(strPath, strASAccount) ' KB2811566 Fix
      End If
      If strActionSQLAS <> "ADDNODE" Then
        Call ResetDBAFilePerm(strDirProg & "\" & strInstRegAS) 
        Call ResetDBAFilePerm(strDirDataAS) 
        Call ResetDBAFilePerm(strDirLogAS) 
        Call ResetDBAFilePerm(strDirTempAS) 
        Call ResetDBAFilePerm(strDirBackupAS) 
        Call ResetFilePerm(strDirLogAS,    strASAccount)
        Call ResetFilePerm(strDirBackupAS, strASAccount)
        Call ResetFilePerm(strDirBackupAS, strAgtAccount)
      End If
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES" 
      ' Nothing
    Case Else
      Call DebugLog("Reset permissions on SQL DB folders")
      Call ResetDBAFilePerm(strDirProg)
      Call ResetDBAFilePerm(objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLPath")) 
      Call ResetDBAFilePerm(objShell.RegRead(strHKLMSQL & strInstRegSQL & "\Setup\SQLPath") & "\FTData")
      Call ResetFilePerm(strDirProg, strSQLAccount)
'      Call ResetDBAFilePerm(strPathVS)
      If strSetupBPE = "YES" Then
        Call ResetFilePerm(strDirBPE, strSQLAccount)
      End If
      strPath       = strWinDir & "\System32\LogFiles\Sum"
      If objFSO.FolderExists(strPath) Then
        Call ResetFilePerm(strPath, strSQLAccount) ' KB2811566 Fix
      End If
      If strActionSQLDB <> "ADDNODE" Then
        Call ResetDBAFilePerm(strDirProg & "\" & strInstRegSQL)
        Call ResetDBAFilePerm(strDirData) 
        Call ResetDBAFilePerm(strDirLog) 
        Call ResetDBAFilePerm(strDirLogTemp) 
        Call ResetDBAFilePerm(Left(strDirTemp, Len(strDirTemp) - 7)) 
        Call ResetDBAFilePerm(strDirBackup)
        Call ResetDBAFilePerm(strDirSysDB)
        Call ResetFilePerm(strDirBackup, strSQLAccount)
        Call ResetFilePerm(strDirBackup, strAgtAccount) 
        Select Case True
          Case strSetupSQLDBFS <> "YES"
            ' Nothing
          Case strFSLevel < "2"
            ' Nothing
          Case Else
            Call ResetDBAFilePerm(strDirDataFS)
            Call ResetFilePerm(strDirDataFS, strSQLAccount)
        End Select 
        Select Case True
          Case strSetupSQLDBFT <> "YES"
            ' Nothing
          Case Else
            Call ResetDBAFilePerm(strDirDataFT)
            Call ResetFilePerm(strDirDataFT, strSQLAccount)
        End Select
      End If
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDBARegistryPermissions()
  Call SetProcessId("2CAH", "Setup Registry Permissions for DBAs")

  Call SetRegPerm("HKEY_LOCAL_MACHINE\" & Mid(strHKLMSQL, 6),            strGroupDBA,      "F")
  Call SetRegPerm("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSSQLServer",   strGroupDBA,      "F")

  If strGroupDBANonSA <> "" Then
    Call SetRegPerm("HKEY_LOCAL_MACHINE\" & Mid(strHKLMSQL, 6),          strGroupDBANonSA, "R")
    Call SetRegPerm("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\MSSQLServer", strGroupDBANonSA, "R")
  End If

  Call ProcessEnd(strStatusComplete)
End Sub


Sub SetupWMIPermissions()
  Call SetProcessId("2CAI", "Setup WMI Permissions for DBAs")

  Select Case True
    Case strOSVersion < "6.0" 
      ' Nothing
    Case Else
      Call SetWMIPerm(strGroupDBA,      "F")
  End Select

  Select Case True
    Case Left(strOSVersion, 1) < "6" 
      ' Nothing
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      Call SetWMIPerm(strGroupDBANonSA, "R")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetWMIPerm(strName, strAccess)
  Call DebugLog("SetWMIPerm: " & strName)
  ' Code based on example posted by Bart Van Hecke on http://www.sapien.com/forums/viewtopic.php?f=19&t=5919
  Dim objSS, objSD

  Set objSS         = objWMI.Get("__SystemSecurity=@")
  objSS.GetSecurityDescriptor objSD
  Set objSD         = AddACE(objSD, strName, strAccess)
  objSS.SetSecurityDescriptor objSD

  Set objSD         = Nothing
  Set objSS         = Nothing

End Sub


Function AddACE(objSD, strName, strAccess)
  Call DebugLog("AddACE: " & strName)
  Dim colAccount
  Dim objAccount, objAccountSID, objACE, objTrustee
  Dim arrDACL
  Dim strACEAccessAllow, strACEFullWrite, strACEPropogate, strACEEnableAccount, strACEInherit, strACERemoteEnable
  Dim strLocal, strSID, strQueryDomain, strQueryName
  Dim intIdx, intUBound
  Dim SE_DACL_PROTECTED

  strACEAccessAllow = 0
  strACEPropogate   = 2
  strACEEnableAccount = 1
  strACEInherit       = 16
  strACERemoteEnable  = 32
  strACEFullWrite   = 4
  SE_DACL_PROTECTED = &h0001
  strLocal          = ""

  intIdx            = InStr(strName, "\")
  Select Case True
    Case intIdx = 0
      strQueryDomain  = strDomain
      strQueryName    = strName
    Case Left(strName, intIdx - 1) = strDomain
      strQueryDomain  = strDomain
      strQueryName    = Mid(strName, intIdx + 1)
    Case Else
      strQueryDomain  = strServer
      strQueryName    = Mid(strName, intIdx + 1)
      strLocal        = " AND LocalAccount=True"
  End Select

  arrDACL           = objSD.DACL
  intUBound         = UBound(arrDACL)
  Set objACE        = arrDACL(intUBound)
  Select Case True
    Case strAccess = "F"
      objACE.AccessMask = strACERemoteEnable + strACEFullWrite
    Case Else
      objACE.AccessMask = strACERemoteEnable + strACEEnableAccount
  End Select
  objACE.ACEType    = strACEAccessAllow
  objACE.ACEFlags   = strACEInherit + strACEPropogate

  Set objTrustee    = objACE.Trustee
  Set colAccount    = objWMI.ExecQuery("SELECT * FROM Win32_Group WHERE Domain='" & strQueryDomain & "' AND Name='" & strQueryName & "'" & strLocal)
  For Each objAccount in colAccount
    Set objAccountSID = objWMI.Get("Win32_SID.SID='" & objAccount.SID &"'") 
    Exit For
    Next
  objTrustee.Domain = strQueryDomain 
  objTrustee.Name   = strQueryName 
  objTrustee.SID    = objaccountSID.BinaryRepresentation 
  objTrustee.SIDLength  = objaccountSID.SIDLength
  objTrustee.SIDString  = objAccount.SID
  objACE.Trustee    = objTrustee

  intUBound         = intUBound + 1
  ReDim Preserve arrDACL(intUBound)
  Set arrDACL(intUBound) = objACE
  objSD.DACL        = arrDACL

  If (objSD.ControlFlags And SE_DACL_PROTECTED) = SE_DACL_PROTECTED Then
    objSD.ControlFlags = objSD.ControlFlage Xor SE_DACL_PROTECTED
  End If

  Set AddACE        = objSD

End Function


Sub SetupServiceDependencies()
  Call SetProcessId("2CAJ", "Setup Service Dependencies")
  Dim arrDepends
  Dim intDepend, intIdx, intIdxNew

  strPath     = "SYSTEM\CurrentControlSet\Services\msftesql\"
  objWMIReg.GetMultiStringValue strHKLM, strPath, "DependOnService", arrDepends
  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBFT <> "YES"
      ' Nothing
    Case Not IsArray(arrDepends)
      ' Nothing
    Case UBound(arrDepends) = 0
      ' Nothing
    Case Else
      intIdxNew     = -1
      intDepend     = Ubound(arrDepends)
      ReDim arrDependsNew(intDepend)
      For intIdx = 0 To intDepend
        If UCase(arrDepends(intIdx)) <> "NTLMSSP" Then
          intIdxNew = intIdxNew + 1
          arrDependsNew(intIdxNew) = arrDepends(intIdx)
        End If
      Next
      If intIdxNew < 0 Then
        intIdxNew   = 0
        arrDependsNew(intIdxNew) = vbNullChar
      End If
      ReDim Preserve arrDependsNew(intIdxNew)
      objWMIReg.SetMultiStringValue strHKLM, strPath, "DependOnService", arrDependsNew
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CheckSQLAccounts()
  Call SetProcessId("2CAK", "Check SQL Server Accounts")

  Select Case True
    Case Ucase(strSqlAccount) = Ucase(strNTAuthOSName )
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strNTAuthAccount & "', 'sysadmin';""", 1)
    Case Else
      strCmd        = "CREATE LOGIN [" & strSqlAccount & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
      Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strSqlAccount & "', 'sysadmin';""", 1)
  End Select

  Select Case True
    Case strAgtAccount = strSqlAccount
      ' Nothing
    Case Ucase(strAgtAccount) = Ucase(strNTAuthOSName )
      strCmd        = "CREATE LOGIN [" & strNTAuthAccount & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
      If strSQLVersion = "SQL2005" Then
        Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strNTAuthAccount & "', 'sysadmin';""", 1)
      End If
    Case Else
      strCmd        = "CREATE LOGIN [" & strAgtAccount & "] FROM WINDOWS;"
      Call Util_ExecSQL(strCmdSQL & "-Q", """" & strCmd & """", 1)
      If strSQLVersion = "SQL2005" Then
        Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_addsrvrolemember '" & strAgtAccount & "', 'sysadmin';""", 1)
      End If
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CheckSQLBrowser()
  Call SetProcessId("2CAL", "Check SQL Browser service")

  Select Case True
    Case strSQLVersion = "SQL2005" And strSqlBrowserStartup = "1"
      Call Util_RunExec("NET START ""SQLBrowser""", "", "", 2)
    Case strSQLVersion > "SQL2005" And UCase(strSqlBrowserStartup) = "Automatic"
      Call Util_RunExec("NET START ""SQLBrowser""", "", "", 2)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupClusterBindings()
  Call SetProcessId("2CAM", "Setup Cluster bindings")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAMA"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strSetupSQLASCluster = "YES"
      Call SetupASClusterBindings()
    Case GetBuildfileValue("ClusterASFound") = "Y"
      Call RemoveOwner("SQL Network Name (" & strClusterNameAS & ")", "")
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAMB"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      Call SetupDBClusterBindings()
    Case GetBuildfileValue("ClusterSQLFound") = "Y"
      Call RemoveOwner("SQL Network Name (" & strClusterNameSQL & ")", "")
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAMC"
      ' Nothing
    Case strSetupDTCCluster = "YES"
      Call SetupDTCClusterBindings()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CAMD"
      ' Nothing
    Case Else
      Call SetupClusterProperties()
  End Select

  Call SetProcessId("2CAMZ", " Setup Cluster bindings" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupASClusterBindings()
  Call SetProcessId("2CAMA", "Setup SSAS Cluster Bindings")

  Call SetResourceOff("SQL Network Name (" & strClusterNameAS & ")", "")
  Select Case True
    Case strActionSQLAS = "ADDNODE"
      Call AddOwner(strClusterNetworkAS)
      Call AddOwner("SQL IP Address 1 (" & strClusterNameAS & ")")
      Call SetOwnerNode(strClusterGroupAS)
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Network Name (" & strClusterNameAS & ")"" /PRIV RequireKerberos=1:DWORD "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetSQLVolDep(strClusterGroupAS, "Analysis Services" & strResSuffixAS)
      Call MSDTCBind(strClusterNameDTC, strClusterNameAS, strInstAS)
      Call SetOwnerNode(strClusterGroupAS)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDBClusterBindings()
  Call SetProcessId("2CAMB", "Setup SQL DB Cluster Bindings")

  Select Case True
    Case strActionSQLDB = "ADDNODE"
      Call AddOwner(strClusterNetworkSQL)
      Call SetOwnerNode(strClusterGroupSQL)
    Case Else
      Call SetResourceOff("SQL Network Name (" & strClusterNameSQL & ")", "")
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Network Name (" & strClusterNameSQL & ")"" /PROP PendingTimeout=600000:DWORD "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Network Name (" & strClusterNameSQL & ")"" /PRIV RequireDNS=1:DWORD "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE ""SQL Network Name (" & strClusterNameSQL & ")"" /PRIV RequireKerberos=1:DWORD "
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetSQLVolDep(strClusterGroupSQL, "SQL Server" & strResSuffixDB)
      Call MSDTCBind(strClusterNameDTC, strClusterNameSQL, strInstSQL)
      Call SetOwnerNode(strClusterGroupSQL)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupDTCClusterBindings()
  Call SetProcessId("2CAMC", "Setup DTC Cluster Bindings")

  Select Case True
    Case strClusterAction = "ADDNODE"
      Call SetOwnerNode(strClusterGroupDTC)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupClusterProperties()
  Call SetProcessId("2CAMD", "Setup Cluster Properties")

  Select Case True
    Case strOSVersion < "6" 
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESTYPE " & GetResType("Network Name") & " /PRIV DeleteVcoOnResCleanup=1:DWORD "
      Call Util_RunExec(strCmd, "", strResponseYes, 13)
  End Select

  If strSetupSQLASCluster = "YES" Then
    Call SetResourceOn(strClusterGroupAS, "GROUP")
    Call SetBuildfileValue("SetupSQLASClusterStatus", strStatusComplete)
  End If

  If strSetupSQLDBCluster = "YES" Then
    Call SetResourceOn(strClusterGroupSQL, "GROUP")
    Call SetBuildfileValue("SetupSQLDBClusterStatus", strStatusComplete)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Function GetResType(strResType)
  Call DebugLog("GetResType: " & strResType)
  Dim strPathReg, strRestypeName

  strPath           = "HKLM\Cluster\ResourceTypes\" & strResType & "\Name"
  strRestypeName    = objShell.RegRead(strPath)
  GetRestype        = """" & strRestypeName & """"

End Function


Sub SetSQLVolDep(strClusterGroup, strClusterService)
  Call DebugLog("SetSQLVolDep: " & strClusterGroup)
  Dim colClusGroups, colClusResources
  Dim objClusGroup, objClusResource

  Set colClusGroups = GetClusterGroups
  For Each objClusGroup In colClusGroups
    If objClusGroup.Name = strClusterGroup Then                   
      Set colClusResources = objClusGroup.Resources
      For Each objClusResource In colClusResources
        If objClusResource.TypeName = "Physical Disk" Then
          strCmd    = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterService & """ /ADDDEP:""" & objClusResource.Name & """"
          Call Util_RunExec(strCmd, "", strResponseYes, 5003)
        End If
      Next
    End If
  Next
  Set colClusResources = Nothing
  Set colClusGroups    = Nothing 

End Sub


Sub MSDTCBind(strClusterNameDTC, strCluster, strInst)
  Call DebugLog("MSDTCBind: " & strClusterNameDTC & " to " & strCluster & " instance " & strInst)
  Dim strMapping

  If strOSVersion >= "6" Then
    strPath         = "Cluster\MSDTC\TMMapping\Service\" & strCluster & "\"
    strMapping      = ""
    objWMIReg.GetStringValue strHKLM,strPath,"Name",strMapping
    If strMapping > "" Then
      objShell.RegDelete "HKLM\" & strPath
    End If
    strCmd          = "MSDTC -tmMappingSet -name """ & strCluster & """ -service """ & strInst & """ -ClusterResourceName """ & strClusterNameDTC & """ "
    Call Util_RunCmdASync(strCmd, 0)
    Wscript.Sleep strWaitShort
    strCmd          = "WMIC process WHERE ""CommandLine LIKE '%MSDTC.EXE%'"" CALL terminate"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

End Sub


Sub RegisterManagementServer()
  Call SetProcessId("2CAN", "Register Management Server Name in Registry")

  If strManagementServerRes = "" Then
    strManagementServerRes = strManagementServer
    Call Util_RegWrite(strHKLMFB & "ManagementServerRes", strManagementServerRes, "REG_SZ")
    Call SetBuildfileValue("ManagementServerRes", strManagementServerRes)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAnalytics()
  Call SetProcessId("2CB","Setup R Analytics")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CBA"
      ' Nothing
    Case Else
      Call SetupAnalConfig()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CBB"
      ' Nothing
    Case strSetupRServer <> "YES"
      ' Nothing
    Case Else
      Call SetupRServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CBC"
      ' Nothing
    Case strSetupRServer <> "YES"
      ' Nothing
    Case Else
      Call SetupRUsers()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CBD"
      ' Nothing
    Case strSetupPython <> "YES"
      ' Nothing
    Case Else
      Call SetupPython()
  End Select

  Call SetBuildFileValue("SetupAnalyticsStatus", strStatusComplete)
  Call SetProcessId("2CBZ", " Setup R Analytics" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupAnalConfig()
  Call SetProcessId("2CBA", "Setup SQL Config for Analytics")

  Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'show advanced options',           '1';""", 0)
  Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)

  If strSQLSharedMR = "YES" Then
    Wscript.Sleep strWaitShort
    Call Util_ExecSQL(strCmdSQL & "-Q", """EXEC sp_configure 'external scripts',              '1';""", 0)
    Call Util_ExecSQL(strCmdSQL & "-Q", """RECONFIGURE WITH OVERRIDE;""", 0)
  End If

  Call SetBuildFileValue("SetupAnalyticsStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRServer()
  Call SetProcessId("2CBB", "Setup R Server")
  Dim strSQLCmd1, strSQLCmd2, strSQLRUserGroup

  strSQLRUserGroup  = strServer & "\" & "SQLRUserGroup"
  strSQLCmd1        = "CREATE LOGIN [" & strSQLRUserGroup & "] FROM WINDOWS;"
  strSQLCmd2        = "GRANT CONNECT SQL TO [" & strSQLRUserGroup & "];"

  Select Case True
    Case strActionSQLDB = "ADDNODE"
      Call SetBuildMessage(strMsgInfo, "Run the following when SQL running on Node " & strServer & ": " & strSQLCmd1)
      Call SetBuildMessage(strMsgInfo, "Run the following when SQL running on Node " & strServer & ": " & strSQLCmd2)
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-r -Q", """" & strSQLCmd1 & """", 1)
      Call Util_ExecSQL(strCmdSQL & "-d ""master"" -Q", """" & strSQLCmd2 & """", 1)
  End Select

  Call SetBuildFileValue("SetupRServerStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRUsers()
  Call SetProcessId("2CBC", "Setup R Users")
  Dim objGroup, objMember

  If strSQLVersion < "SQL2019" Then
    Set objGroup    = GetObject("WinNT://" & strServer & "/SQLRUserGroup,group")
    For Each objMember In objGroup.Members
      strCmd        = "NET LOCALGROUP """ & strGroupUsers & """ """ & strServer & "\" & objMember.Name & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
    Next
  End If

  Call SetBuildFileValue("SetupRServerStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupPython()
  Call SetProcessId("2CBD", "Setup Python Support")

' No actions needed

  Call SetBuildFileValue("SetupPythonStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub CheckChildClusters()
  Call SetProcessId("2CC", "Check SQL DB Child Clusters")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCAZ"
      ' Nothing
    Case strSetupPolyBaseCluster <> "YES"
      ' Nothing
    Case strSQLVersion >= "SQL2016"
      Call SetBuildFileValue("SetupPolyBaseClusterStatus", strStatusComplete)
    Case Else
      Call SetupPolyBaseCluster()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCBZ"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case strSetupSSISCluster <> "YES"
      ' Nothing
    Case Else
      Call SetupSSISCluster()
  End Select

  Call SetProcessId("2CCZ", " Check SQL DB Child Clusters" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupPolyBaseCluster()
  Call SetProcessId("2CCA", "Setup PolyBase Cluster")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCAA"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call CreatePBEngineCluster()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCAB"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call CreatePBDMCluster()
  End Select

  Call SetBuildfileValue("SetupPolyBaseClusterStatus", strStatusComplete)
  Call SetProcessId("2CCAZ", " Setup PolyBase Cluster" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CreatePBEngineCluster()
  Call SetProcessId("2CCAA", "Setup PolyBase Engine Cluster")

  Call AddChildCluster("PolyBaseCluster", strClusterGroupSQL, strClusterNamePE, "SQL Network Name (" & strClusterNameSQL & ")", strInstPE, "SQL Server PolyBase Engine (" & strInstance & ")", "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreatePBDMCluster()
  Call SetProcessId("2CCAB", "Setup PolyBase Data Movment Cluster")

  Call AddChildCluster("PolyBaseCluster", strClusterGroupSQL, strClusterNamePM, "SQL Network Name (" & strClusterNameSQL & ")", strInstPM, "SQL Server PolyBase Data Movement (" & strInstance & ")", "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSSISCluster()
  Call SetProcessId("2CCB", "Setup SSIS Cluster")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCBA"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call CreateSSISCluster()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCBB"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case Else
      Call UpdateSSISConfig()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "2CCBC"
      ' Nothing
    Case Else
      Call AddSSISClusterNode()
  End Select 

  Call SetProcessId("2CCBZ", " Setup SSIS Cluster" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub CreateSSISCluster()
  Call SetProcessId("2CCBA", "Create SSIS Cluster")
  Dim strPathSSIS, strSSISConfig

  Call SetBuildfileValue("SetupSSISClusterStatus", strStatusProgress)
  strSSISConfig     = objShell.RegRead(strRegSSIS)
  If strSSISConfig = "" Then
    strSSISConfig   = objShell.RegRead(strRegSSISSetup) & "Binn\MsDtsSrvr.ini.xml"
    Call Util_RegWrite(strRegSSIS, strSSISConfig, "REG_SZ")
  End If

  Call AddChildCluster("IS", strClusterGroupSQL, strClusterNameIS, "SQL Network Name (" & strClusterNameSQL & ")", strInstIS, "SQL Server Integration Services", "")

  Call ProcessEnd(strStatusComplete)

End Sub


Sub UpdateSSISConfig()
  Call SetProcessId("2CCBB", "Update SSIS Config File")
  Dim colSSISConfig
  Dim objSSISConfig, objSSISFolder
  Dim strSSISConfig, strSSISFolder, strSSISNode

  Call SetResourceOff(strClusterNameIS, "")

  Call DebugLog("Move SSIS Config File")
  strSSISConfig     = objShell.RegRead(strRegSSIS)
  Set objFile       = objFSO.GetFile(strSSISConfig)
  strPathNew        = strDirDataIS & "\" & objFile.Name
  objFile.Copy strPathNew
  Call Util_RegWrite(strRegSSIS, strPathNew, "REG_SZ")

  Call DebugLog("Update SSIS Config File")
  strSSISConfig     = objShell.RegRead(strRegSSIS)
  Set objSSISConfig = CreateObject ("Microsoft.XMLDOM") 
  objSSISConfig.async = "false"
  objSSISConfig.load(strSSISConfig)
  Set colSSISConfig = objSSISConfig.getElementsByTagName("Folder")
  For Each objSSISFolder In colSSISConfig
    strSSISFolder   = objSSISFolder.getAttribute("xsi:type")
    Select Case True
      Case strSSISFolder = "SqlServerFolder" 
        Set strSSISNode  = objSSISFolder.selectSingleNode("./ServerName")
        If Left(strSSISNode.Text, 1) = "." Then
          strSSISNode.Text = strServInst
        End If
      Case strSSISFolder = "FileSystemFolder" 
        Set strSSISNode  = objSSISFolder.selectSingleNode("./StorePath")
        If strSSISNode.Text = "..\Packages" Then
          strSSISNode.Text = strDirDataIS & "\Packages"
        End If
    End Select
  Next
  objSSISConfig.save(strSSISConfig)
  Set colSSISConfig = Nothing
  Set objSSISConfig = Nothing

  strCmd            = "CLUSTER """ & strClusterName & """ RESOURCE """ & strClusterNameIS  & """ /ADDCHECKPOINTS:""" & Mid(strRegSSIS, 6) & """" 
  Call Util_RunExec(strCmd, "", strResponseYes, 183)

  Call SetBuildfileValue("SetupSSISClusterStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub AddSSISClusterNode()
  Call SetProcessId("2CCBC", "Add Node to SSIS Cluster")

  Call AddChildNode("IS", strClusterNameIS)

  Call SetBuildfileValue("SetupSSISClusterStatus", strStatusComplete)
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


End Class