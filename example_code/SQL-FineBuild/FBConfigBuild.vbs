''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FBConfigBuild.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      Build FineBuild Configuration 
'
'  Author:       Ed Vassie
'
'  Date:         23 Sep 2008
'
'  Change History
'  Version  Author        Date         Description
'  1.0      Ed Vassie     23 Sep 2008  Initial version
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim arrClusInstances(26)
Dim colArgs, colBuild, colFiles, colFlags, colGlobal, colStrings, colSysEnvVars, colUsrEnvVars
Dim objAccount, objADOConn, objADRoot, objConfig, objDrive, objFile, objFolder, objFSO, objNetwork, objRE, objShell, objStatefile, objSysInfo, objWMI, objWMIDNS, objWMIReg
Dim intIdx, intProcNum, intSQLMemory, intSpeedTest, intTimer
Dim strAction, strActionAO, strActionDAG, strActionDTC, strActionSQLDB, strActionSQLAS, strActionSQLIS, strActionSQLRS, strActionSQLTools, strActionClusInst, strAdminPassword, strADRoot, strAGName, strAGDagName, strAGDagNodes, strAgentJobHistory, strAgentMaxHistory, strAllowUpgradeForRSSharePointMode, strAllUserProf, strAllUserDTop, strAlphabet, strAnyKey, strASProviderMSOlap, strAsServerMode, strAsSvcStartuptype, strAGTSvcStartuptype, strWriterSvcStartupType, strAuditLevel, strAutoLogonCount, strAVCmd
Dim strAgtAccount, strASAccount, strSqlBrowserAccount, strCmdShellAccount, strCtlrAccount, strCltAccount, strFtAccount, strIsAccount, strMDWAccount, strRsAccount, strRsExecAccount, strRsShareAccount, strSqlAccount, strSQLAcDomain, strSQLAgentStart, strLocalAdmin
Dim strAgtPassword, strASPassword, strSqlBrowserPassword, strCmdShellPassword, strCtlrPassword, strCltPassword, strFtPassword, strIsPassword, strMDWPassword, strRsPassword, strRsExecPassword, strRsSharePassword, strSqlPassword, strsaName, strsaPwd
Dim strBackupStart, strBackupRetain, strBackupDiffRetain, strBackupLogFreq, strBackupLogRetain, strBPEFile, strBuiltinDom
Dim strCatalogServer, strCatalogServerName, strCatalogInstance, strCheckRegPerm, strCollationAS, strCollationSQL, strCompatFlags, strCSVFound, strCSVRoot, strConfig, strConfirmIPDependencyChange, strCmd, strCmdPS, strCmdRS
Dim strCLSIdDTExec, strCLSIdNetCon, strCLSIdRunBroker, strCLSIdSQL, strCLSIdSQLSetup, strCLSIdSSIS, strCLSIdVS
Dim strClusGroups, strClusNetworkId, strClusStorage, strClusSubnet, strClusSuffix, strClusterAction, strClusterBase, strClusterSQLFound, strClusterHost, strClusterName, strClusterNameAS, strClusterNameDTC, strClusterNameIS, strClusterNamePE, strClusterNamePM, strClusterNameRS, strClusterNetworkAS, strClusterNetworkDTC, strClusterNetworkSQL, strClusterPath, strClusterGroupAO, strClusterGroupAS, strClusterGroupDTC, strClusterGroupFS, strClusterGroupRS, strClusterGroupSQL, strClusterNameSQL, strClusterNode, strClusterReport, strClusterTCP, strClusAASuffix, strClusAOSuffix, strClusASSuffix, strClusDTCSuffix, strClusDBSuffix, strClusFSSuffix, strClusIMSuffix, strClusISSuffix, strClusMRSuffix, strClusPESuffix, strClusPMSuffix, strClusWinSuffix, strClusRSSuffix, strClusIPAddress, strClusIPVersion, strClusIPV4Address, strClusIPV4Mask, strClusIPV4Network, strClusIPV6Address, strClusIPV6Mask, strClusIPV6Network, strClusterPassive
Dim strClusterAOFound, strClusterASFound, strClusterDTCFound, strDTCClusterRes, strDTCMultiInstance, strDC, strDfltDoc, strDfltProf, strDfltRoot, strDQPassword, strCltStartupType, strCtlrStartupType
Dim strDBA_DB, strDBAEmail, strDefaultAccount, strDefaultPassword, strDefaultUser, strDiscoverFile, strDiscoverFolder, strDNSIPIM, strDNSNameIM, strDomain, strDomainComputers, strDomainUsers, strDomainSID, strDirASDLL, strDirDBA, strDirDRU, strDirProg, strDirProgSys, strDirProgSysX86, strDirProgX86, strDirServInst, strDirSQL, strDirSQLBootstrap, strDirSys, strDirSysData, strDirTempWin, strDirTempUser, strDisableNetworkProtocols, strDistDatabase, strDistPassword, strDriveList
Dim strEdition, strEditionEnt, strEdType, strEnableRANU, strEncryptAO, strENU, strErrorReporting, strExpVersion, strExtSvcAccount, strExtSvcPassword
Dim strFailoverClusterRollOwnership, strFarmAccount, strFarmPassword, strFarmAdminIPort, strFBCmd, strFBParm, strFeatures, strFileArc, strFilePerm, strFineBuildStatus, strFirewallStatus, strFSInstLevel, strFSLevel, strFSShareName, strFTUpgradeOption, strGroupAO
Dim strPBDMSSvcAccount, strPBDMSSvcPassword, strPBDMSSvcStartup, strPBEngSvcAccount, strPBEngSvcPassword, strPBEngSvcStartup, strPBPortRange, strPBScaleout, strPID, strPowerBIexe, strPowerBIPID, strPreferredOwner, strProfDir, strProfileName, strProgCacls, strProgNTRights, strProgSetSPN, strProgReg, strPSInstall
Dim strGroupAdmin, strGroupDBA, strGroupDBANonSA, strGroupMSA, strDBMailProfile, strDBOwnerAccount
Dim strGroupDistComUsers, strGroupIISIUsers, strGroupPerfLogUsers, strGroupPerfMonUsers, strGroupRDUsers, strGroupUsers
Dim strHKLM, strHKLMFB, strHKU, strHTTP, strHistoryRetain
Dim strIsInstallDBA, strInstance, strInstAO, strInstMR, strInstRegAS, strInstRegSQL, strInstRegRS, strInstPE, strInstPM, strInstRS, strInstRSDir, strInstRSHost, strInstRSSQL, strInstRSURL, strInstRSWMI, strInstADHelper, strIsSvcStartuptype
Dim strInstAgent, strInstAnal, strInstAS, strInstASCon, strInstASSQL, strInstFT, strInstIS, strInstISMaster, strInstISWorker, strInstNode, strInstNodeAS, strInstNodeIS, strInstLog, strInstTel, strIISAccount, strIISRoot, strIsMasterAccount, strIsMasterPassword, strIsMasterStartupType, strIsMasterPort, strIsMasterThumbprint, strIsWorkerAccount, strIsWorkerPassword, strIsWorkerStartupType, strIsWorkerMaster, strIsWorkerCert
Dim strJobCategory, strStartJobPassword, strKeepAliveCab
Dim strLabBackup, strLabBackupAS, strLabBPE, strLabData, strLabDataAS, strLabDataFS, strLabDataFT, strLabDTC, strLabLog, strLabLogAS, strLabLogTemp, strLabPrefix, strLabProg, strLabSysDB, strLabSystem, strLabTemp, strLabTempAS, strLabTempWin, strLabDBA, strLanguage, strLocalDomain, strLogFile, strListAddNode, strListType, strListCluster, strListCompliance, strListCore, strListEdition, strListMain, strListMSA, strListOSVersion, strListSQLAS, strListSQLDB, strListSQLRS, strListSQLVersion, strListSQLTools, strListSSAS, strListSSIS
Dim strMainInstance, strSQLMaxDop, strMailServer, strMailServerType, strManagementAlias, strManagementDW, strManagementServer, strManagementServerRes, strManagementServerName, strManagementInstance, strMDSAccount, strMDSPassword, strMDSDB, strMDSPort, strMDSSite, strMembersDBA, strMode, strMountRoot, strMSSupplied
Dim strNativeOS, strNetNameSource, strNet4Xexe, strNetworkGUID, strNPEnabled, strNTAuth, strNTAuthAccount, strNTAuthEveryone, strNTAuthNetwork, strNTAuthOSName, strNTService, strNumErrorLogs, strNumLogins, strNumTF
Dim strOptions, strOSBuild, strOSName, strOSLanguage, strOSLevel, strOSRegPath, strOSType, strOSVersion, strOUPath, strOUCName
Dim strPath, strPathBOL, strPathAddComp, strPathAddCompOrig, strPathAutoConfig, strPathAutoConfigOrig, strPathCScript, strPathFB, strPathFBStart, strPathFBScripts, strPathNew, strPathPS, strPathSQLDefault, strPathSQLMedia, strPathSQLMediaOrig, strPathSQLSP, strPathSQLSPOrig, strPathSSIS, strPathSSMS, strPathSSMSX86, strPathSys, strPathVS, strPassphrase, strProcArc
Dim strRebootLoop, strRegasmExe, strRemoteRoot, strReportOnly, strReportViewerVersion, strRoleDBANonSA, strRSAlias, strRSName, strRsFxVersion, strRSDBAccount, strRSDBPassword, strRSDBName, strRSEmail, strRSFullURL, strRSInstallMode, strRSShpInstallMode, strRSSQLLocal, strRSURLSuffix, strRSVersion, strRole, strResSuffixAS, strResSuffixDB, strRsSvcStartuptype, strRSVersionNum, strRunCount
Dim strSecMain, strSecDBA, strSecTemp, strSecurityMode, strSIDDistComUsers, strSIDIISIUsers, strSNACFile, strSSISDB, strSSISPassword, strSSISRetention, strSSMSexe, strSPFile, strSPLevel, strSPCULevel, strStatefile, strStreamInsightPID
Dim strStatusAssumed, strStatusBypassed, strStatusComplete, strStatusFail, strStatusManual, strStatusPreConfig, strStatusProgress, strStatusKB2919355, strStatusRobocopy, strStatusXcopy, strStopAt, strSQLAdminAccounts, strSSASAdminAccounts
Dim strSetCLREnabled, strSetCostThreshold, strSetLowMemLimit, strSetHardMemLimit, strSetTotalMemLimit, strSetVertiMemLimit, strSetHeaderLength, strSetMemOptHybridBP, strSetMemOptTempdb, strSQLMaxMemory, strSQLMinMemory, strSetOptimizeForAdHocWorkloads, strSetRemoteAdminConnections, strSetRemoteProcTrans, strSetROLAPDimensionProcessingEffort, strSetWorkingSetMaximum
Dim strSetupAlwaysOn, strSetupAOAlias, strAOAliasOwner, strSetupAOProcs, strSetupAODB, strSetupAPCluster, strSetupAutoConfig, strSetupCmdshell, strSetupCompliance, strSetupClusterShares, strSetupDBAManagement, strSetupDBOpts, strSetupDisableSA, strSetupFT, strSetupNetBind, strSetupNetName, strSetupNoDefrag, strSetupNonSAAccounts, strSetupNoSSL3, strSetupNoTCPNetBios, strSetupNoTCPOffload, strSetupNoWinGlobal, strSetupOLAP, strSetupOLAPAPI, strSetupSAAccounts, strSetupSAPassword, strSetupServices, strSetupServiceRights, strSetupSnapshot, strSetupSQLAgent, strSetupSQLDebug, strSetupSQLInst, strSetupSQLPowershell, strSetupSQLServer, strSetupSSL, strSetupStdAccounts, strSetupSysDB, strSetupSysIndex, strSetupSysManagement, strSetupAnalytics, strSetupPowerBI, strSetupPowerBIDesktop, strSetupPSRemote, strSetupPython, strSetupRServer, strSetupRSAdmin, strSetupRSAlias, strSetupRSDB, strSetupRSExec, strSetupRSIndexes, strSetupRSKeepAlive, strSetupRSShare, strSetupRSAT
Dim strSetupABE, strSetupKerberos, strSetupPDFReader, strSetupPerfDash, strSetupPolyBase, strSetupPolyBaseCluster, strSetupProcExp, strSetupProcMon, strSetupRawReader, strSetupRptTaskPad, strSetupRSScripter, strSetupSamples, strSetupSemantics, strSetupIntViewer, strSetupISMaster, strSetupISMasterCluster, strSetupISWorker, strSetupJavaDBC, strSetupStartJob, strSetupJRE, strSetupMDS, strSetupMDSC, strSetupMDXStudio, strInstSQL
Dim strSetupShares, strSetupSQLAS, strSetupSQLASCluster, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLDBFS, strSetupSQLDBFT, strSetupSQLDBRepl, strSetupSQLNS, strSetupSQLTools, strInstStream, strSetupStreamInsight, strSetupStretch, strSetupSystemViews
Dim strSetupBIDS, strSetupBOL, strSetupBPAnalyzer, strSetupBPE, strSetupCmd, strSetxpCmdshell
Dim strSetupDBMail, strSetupDB2OLE, strSetupDCom, strSetupDimensionSCD, strSetupDistributor, strSetupNoDriveIndex, strSetupDTCCluster, strSetupDTCNetAccess, strSetupDTCNetAccessStatus, strSetupDTSDesigner, strSetupDTSBackup, strSetupDQ, strSetupDQC, strSetupDRUCtlr, strSetupDRUClt, strSetupDTCCID
Dim strSetupFirewall, strSetupGenMaint, strSetupGovernor, strSetupIIS, strSetupKB925336, strSetupKB932232, strSetupKB933789, strSetupKB937444, strSetupKB954961, strSetupKB956250, strSetupKB2549864, strSetupKB2781514, strSetupKB2854082, strSetupKB2862966, strSetupKB2919355, strSetupKB2919442, strSetupKB3090973, strSetupKB4019990, strSetupManagementDW, strSetupMenus, strSetupMBCA, strSetupMSI45, strSetupMSMPI, strSetupMyDocs
Dim strSetupNet3, strSetupNet4, strSetupNet4x, strSetupNetTrust, strSetupNetwork, strSetupOldAccounts, strSetupParam, strSetupPBM, strSetupPlanExplorer, strSetupPlanExpAddin, strSetupPowerCfg, strSetupPS1, strSetupPS2
Dim strSetupReportViewer, strSetupRMLTools, strSetupRSLinkGen, strSetupSQLBC, strSetupSQLCE, strSetupSQLMail, strSetupSQLIS, strSetupSQLRS, strSetupSQLRSCluster, strSetupSQLNexus
Dim strSetupSlipstream, strSetupSP, strSetupSPCU, strSetupSPCUSNAC, strSetupSSDTBI, strSetupSSMS, strSetupSSISCluster, strSetupSSISDB, strSetupBIDSHelper, strSetupCacheManager
Dim strSetupTelemetry, strSetupTempDb, strSetupTempWin, strSetupTLS12, strSetupTrouble, strSetupVC2010, strSetupVS, strSetupVS2005SP1, strSetupVS2010SP1, strSetupWindows, strSetupWinAudit, strSetupXEvents, strSetupXMLNotepad, strSetupZoomIt
Dim strSKUUpgrade, strServer, strServerAO, strServerMB, strServerGroups, strServerIP, strService, strSQLBinRoot
Dim strSQLEmail, strSQLExe, strSQLJavaDir, strSQLLanguage, strSQLLogReinit, strSQLOperator, strSQLProgDir, strSQLMediaArc, strSQLTempdbFileCount, strPCUSource, strCUFile, strCUSource, strRebootStatus, strRegSSIS, strRegSSISSetup, strServInst,  strServIP, strServName, strServParm, strSQLRecoveryComplete, strSQLRSStart, strSQLSharedMR, strSQLSvcStartuptype, strSQLSupportMsi
Dim strSqlBrowserStartup, strSQLList, strSQMReporting
Dim strTallyCount, strTelSvcAcct, strTelSvcPassword, strTelSvcStartup, strtempdbFile, strtempdbLogFile, strValidate, strListSQL, strSQLVersion, strSQLVersionNet, strSQLVersionNum, strSQLVersionWMI, strTCPEnabled, strTCPPort, strTCPPortAO, strTCPPortAS, strTCPPortDAC, strTCPPortDebug, strTCPPortDTC, strTCPPortRS, strTemp, strType, strTypeList
Dim strVarName, strVersionFB, strVSVersionPath, strVSVersionNum, strWaitLong, strWaitMed, strWaitShort, strWOWX86, strXMLNode
Dim strUCServer, strUnknown, strUseFreeSSMS, strUserConfiguration, strUserConfigurationvbs, strUserPreparation, strUserPreparationvbs, strUserProfile, strUserReg, strUpdateSource, strUserDNSDomain, strUserDNSServer, strUserDTop, strUserProf, strUserAdmin, strUserSID, strUseSysDB, strUserAccount, strUserName
Dim strVersionNet3, strVersionNet4,strVolErrorList, strVolFoundList, strVolFBLog, strVolBackup, strVolBackupAS, strVolBPE, strVolData, strVolDataAS, strVolDataFS, strVolDataFT, strVolDTC, strVolLog, strVolLogAS, strVolRoot, strVolRootAS, strVolSysDB, strVolLogTemp, strVolTemp, strVolTempAS, strVolTempWin, strVolProg, strVolSys, strVolDBA, strVolDRU, strVolMDW

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()
 
  Call Process()

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case err.Number = 0 
      ' Nothing
    Case (err.Number = 3010) And (err.Description = "Reboot required")
      Call FBLog("***** Reboot in progress *****")
    Case Left(err.Description, 11) = "Stop forced"
      Call FBLog("***** " & err.Description & " *****")
    Case Else
      Call FBLog("***** Error has occurred *****")
      If err.Number > "" Then
        Call FBLog(" Error code : " & err.Number)
      End If
      If err.Source > "" Then
        Call FBLog(" Source     : " & err.Source)
      End If
      If err.Description > "" Then
        Call FBLog(" Description: " & err.Description)
      End If
      If (strDebugDesc > "") And (strDebugDesc <> err.Description) Then
        Call FBLog(" Last Action: " & strDebugDesc)
      End If
      If strDebugMsg1 <> "" Then
        Call FBLog(" " & strDebugMsg1)
      End If
      If strDebugMsg2 <> "" Then
        Call FBLog(" " & strDebugMsg2)
      End If
      Call FBLog("FineBuild Configuration failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation ()
' Perform initialisation processing

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  strProcessIdLabel = "0"
  Call SetProcessIdCode("FBCB")
  Include "FBUtils.vbs"
  Include "FBManageAccount.vbs"
  Include "FBManageBoot.vbs"
  Include "FBManageCluster.vbs"
  strSQLVersion     = GetBuildfileValue("AuditVersion")

  Set objADOConn    = CreateObject("ADODB.Connection")
  Set objConfig     = CreateObject("Microsoft.XMLDOM")
  Set objNetwork    = WScript.CreateObject("Wscript.Network")
  Set objStatefile  = CreateObject ("Microsoft.XMLDOM") 
  Set objSysInfo    = CreateObject("ADSystemInfo")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colArgs       = Wscript.Arguments.Named
  Set colSysEnvVars = objShell.Environment("System")
  Set colUsrEnvVars = objShell.Environment("User")
  Call SetBuildfileValue("ErrorConfig", "")

  strType           = GetBuildfileValue("Type")
  strTypeList       = " CLIENT CONFIG DISCOVER FIX FULL REBUILD UPGRADE WORKSTATION "
  strXMLNode        = GetBuildfileValue("TypeNode")
  strConfig         = strPathFB & GetBuildfileValue("Config")
  strDebugMsg1      = "Config: " & strConfig
  objConfig.async   = "false"
  objConfig.load(strConfig)
  Set colGlobal     = objConfig.documentElement.selectSingleNode("Global")
  Set colBuild      = objConfig.documentElement.selectSingleNode(strXMLNode)
  Set colFiles      = objConfig.documentElement.selectSingleNode("Files")
  Set colFlags      = objConfig.documentElement.selectSingleNode(strXMLNode + "/Flags")
  Set colStrings    = objConfig.documentElement.selectSingleNode("Global/Strings")

  Set objRE         = New RegExp
  objRE.Global      = True
  objRE.IgnoreCase  = True

  objADOConn.Provider            = "ADsDSOObject"
  objADOConn.Open "ADs Provider"

  strListAddNode    = ""
  strListCluster    = ""
  strListCompliance = ""
  strListCore       = ""
  strListEdition    = ""
  strListMain       = ""
  strListMSA        = ""
  strListSQLDB      = ""
  strListOSVersion  = ""
  strListSQLTools   = ""
  strListSQLRS      = ""
  strListSQLVersion = ""
  strListSSAS       = ""
  strListSSIS       = ""
  strListType       = ""

  strHKLM           = &H80000002
  strHKLMFB         = "HKLM\SOFTWARE\FineBuild\"
  strHKU            = &H80000003
  strInstance       = GetBuildfileValue("Instance")
  strsaPwd          = GetParam(Null,                  "saPwd",              "",                    "")
  strPath           = "SOFTWARE\Microsoft\Windows\CurrentVersion\explorer\Shell Folders\"
  objWMIReg.GetStringValue strHKLM,strPath,"Common Desktop",strAllUserDTop
  objWMIReg.GetStringValue strHKLM,strPath,"Common Start Menu",strAllUserProf
  strActionClusInst = "INSTALLFAILOVERCLUSTER"
  strActionDAG      = ""
  strAdminPassword  = GetParam(Null,                  "AdminPassword",      "",                    "")
  strAGName         = UCase(GetParam(Null,            "AGName",             "ClusterNameAO",       ""))
  strAGDagName      = UCase(GetParam(Null,            "AGDagName",          "",                    ""))
  strAgentJobHistory  = GetParam(colGlobal,           "AgentJobHistory",    "",                    "500")
  strAgentMaxHistory  = GetParam(colGlobal,           "AgentMaxHistory",    "",                    "20000")
  strEncryptAO      = UCase(GetParam(colStrings,      "EncryptAO",          "",                    "AES"))
  strAlphabet       = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
  strAnyKey         = GetParam(colStrings,            "AnyKey",             "",                    "Press any key")
  strASProviderMSOlap = GetParam(colGlobal,           "ASProviderMSOlap",   "",                    "1")
  strAsServerMode   = UCase(GetParam(colGlobal,       "ASServerMode",       "",                    "MultiDimensional"))
  strAuditLevel     = GetParam(colGlobal,             "AuditLevel",         "",                    "2")
  strAVCmd          = GetParam(colStrings,            "AVCmd",              "",                    "POWERSHELL Add-MpPreference -ExclusionPath ")
  strBackupRetain   = GetParam(colStrings,            "SetBackupRetain",    "BackupRetain",        "23")
  strBackupDiffRetain = GetParam(colStrings,          "SetBackupDiffRetain","BackupDiffRetain",    "23")
  strBackupLogFreq  = GetParam(colStrings,            "SetBackupLogFreq",   "BackupLogFreq",       "60")
  strBackupLogRetain  = GetParam(colStrings,          "SetBackupLogRetain", "BackupLogRetain",     "24")
  strBackupStart    = GetParam(colStrings,            "SetBackupStart",     "BackupStart",         "21:00:00")
  strBPEFile        = GetParam(colGlobal,             "BPEFile",            "",                    "100 GB")
  strBuiltinDom     = GetBuildfileValue("BuiltinDom")
  strCatalogInstance  = ""
  strCatalogServer  = Ucase(GetParam(colGlobal,       "CatalogServer",      "",                    ""))
  strCLSIdDTExec    = GetParam(colStrings,            "CLSIdDTExec",        "",                    "%")
  strCLSIdNetCon    = GetParam(colStrings,            "CLSIdNetCon",        "",                    "%")
  strCLSIdRunBroker = GetParam(colStrings,            "CLSIdRunBroker",     "",                    "%")
  strCLSIdSQL       = GetParam(colStrings,            "CLSIdSQL",           "",                    "%")
  strCLSIdSQLSetup  = GetParam(colStrings,            "CLSIdSQLSetup",      "",                    "%")
  strCLSIdSSIS      = GetParam(colStrings,            "CLSIdSSIS",          "",                    "%")
  strCLSIdVS        = GetParam(colStrings,            "CLSIdVS",            "",                    "%")
  strClusterDTCFound  = ""
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterTCP     = Ucase(GetParam(colGlobal,       "ClusterTCP",         "",                    "IPv4"))
  strClusAASuffix   = UCase(GetParam(colStrings,      "ClusAASuffix",       "",                    "AA"))
  strClusAOSuffix   = UCase(GetParam(colStrings,      "ClusAOSuffix",       "",                    "AO"))
  strClusASSuffix   = Ucase(GetParam(colStrings,      "ClusASSuffix",       "",                    "AS"))
  strClusDBSuffix   = Ucase(GetParam(colStrings,      "ClusDBSuffix",       "",                    "DB"))
  strClusDTCSuffix  = UCase(GetParam(colStrings,      "ClusDTCSuffix",      "",                    "TC"))
  strClusFSSuffix   = UCase(GetParam(colStrings,      "ClusFSSuffix",       "",                    "FS"))
  strClusISSuffix   = UCase(GetParam(colStrings,      "ClusISSuffix",       "",                    "IS"))
  strClusIMSuffix   = UCase(GetParam(colStrings,      "ClusIMSuffix",       "",                    "IM"))
  strClusMRSuffix   = UCase(GetParam(colStrings,      "ClusMRSuffix",       "",                    "MR"))
  strClusPESuffix   = UCase(GetParam(colStrings,      "ClusPESuffix",       "",                    "PE"))
  strClusPMSuffix   = UCase(GetParam(colStrings,      "ClusPMSuffix",       "",                    "PM"))
  strClusRSSuffix   = UCase(GetParam(colStrings,      "ClusRSSuffix",       "",                    "RS"))
  strClusWinSuffix  = UCase(GetParam(colStrings,      "ClusWinSuffix",      "",                    ""))
  strClusStorage    = "Unknown"
  strCollationAS    = GetParam(colGlobal,             "ASCollation",        "",                    "Latin1_General_CI_AS")
  strCollationSQL   = GetParam(colGlobal,             "SQLCollation",       "",                    "Latin1_General_CI_AS")
  strCompatFlags    = "HKCU\Software\Microsoft\Windows NT\CurrentVersion\AppCompatFlags\"
  strConfirmIPDependencyChange = GetParam(colGlobal,  "ConfirmIPDependencyChange",             "", "0")
  strCSVFound       = ""
  strDBA_DB         = GetParam(colGlobal,             "DBA_DB",             "",                    "DBA_Data")
  strDBAEmail       = GetParam(colGlobal,             "DBAEmail",           "",                    "")
  strDBMailProfile  = GetParam(colGlobal,             "DBMailProfile",      "",                    "Public DB Mail")
  strDBOwnerAccount = GetParam(colGlobal,             "DBOwnerAccount",     "",                    "DBOwner")
  strDirDBA         = GetParam(colGlobal,             "DirDBA",             "",                    "DBA Files")
  strDirProgSys     = objFSO.GetAbsolutePathName(objShell.ExpandEnvironmentStrings("%PROGRAMFILES%"))
  strDirSQL         = GetParam(colGlobal,             "DirSQL",             "",                    "MSSQL")
  strDirSys         = objFSO.GetAbsolutePathName(objShell.ExpandEnvironmentStrings("%WINDIR%"))
  strDirSysData     = objFSO.GetAbsolutePathName(objShell.ExpandEnvironmentStrings("%PROGRAMDATA%"))
  strDirTempWin     = GetParam(colGlobal,             "DirTempWin",         "",                    "Temp")
  strDirTempUser    = GetParam(colGlobal,             "DirTempUser",        "",                    "Temp")
  strDisableNetworkProtocols = GetParam(colGlobal,    "DisableNetworkProtocols",               "", "0")
  strDistDatabase   = GetParam(colGlobal,             "DistributorDatabase","",                    "Distribution")
  strDistPassword   = GetParam(Null,                  "DistributorPassword","",                    strsaPwd)
  strDomain         = objShell.ExpandEnvironmentStrings("%USERDOMAIN%")
  strDQPassword     = GetParam(Null,                  "DQPassword",         "",                    strsaPwd)
  strDriveList      = ""
  strDTCMultiInstance = Ucase(GetParam(colFlags,      "DTCMultiInstance",   "",                    "Yes"))
  strEdition        = GetBuildfileValue("AuditEdition")
  strEdType         = ""
  strEnableRANU     = GetParam(Null,                  "EnableRANU",         "",                    "1")
  strErrorReporting = GetParam(colGlobal,             "ErrorReporting",     "",                    "0")
  strFailoverClusterRollOwnership = UCase(GetParam(Null, "FailoverClusterRollOwnership","",        ""))
  strFarmAccount    = GetParam(Null,                  "FarmAccount",        "",                    "")
  strFarmPassword   = GetParam(Null,                  "FarmPassword",       "",                    "")
  strFarmAdminIPort = GetParam(Null,                  "FarmAdminIPort",     "",                    "")
  strFBCmd          = objShell.ExpandEnvironmentStrings("%SQLFBCMD%")
  strFBParm         = objShell.ExpandEnvironmentStrings("%SQLFBPARM%")
  strFeatures       = GetParam(colBuild,              "Features",           "",                    "") 
  strFilePerm       = GetBuildfileValue("FilePerm")
  strFineBuildStatus  = GetBuildfileValue("FineBuildStatus")
  strFSLevel        = GetParam(colGlobal,             "FileStreamLevel",    "",                    "2")
  strFSShareName    = UCase(GetParam(Null,            "FileStreamShareName","",                    "FS" & strInstance))
  strFTUpgradeOption  = GetParam(Null,                "FTUpgradeOption",    "",                    "")
  strPath           = "SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\DomainProfile"
  objWMIReg.GetDwordValue strHKLM, strPath, "EnableFirewall", strFirewallStatus
  strGroupAdmin     = GetBuildfileValue("GroupAdmin")
  strGroupUsers     = GetBuildfileValue("GroupUsers")
  strHistoryRetain  = GetParam(colStrings,            "HistoryRetain",      "",                    "30")
  strInstADHelper   = GetParam(colStrings,            "InstADHelper",       "",                    "")
  strInstIS         = GetParam(colStrings,            "InstIS",             "",                    "%")
  strInstISMaster   = GetParam(colStrings,            "InstISMaster",       "",                    "%")
  strInstISWorker   = GetParam(colStrings,            "InstISWorker",       "",                    "%")
  strInstMR         = GetParam(colStrings,            "InstMR",             "",                    "")
  strInstRegAS      = GetBuildfileValue("InstRegAS")
  strInstRegSQL     = GetBuildfileValue("InstRegSQL")
  strInstRegRS      = GetBuildfileValue("InstRegRS")
  strIsMasterPort   = GetParam(colGlobal,             "TCPPortISMaster",    "",                    "8391")
  strIsMasterThumbprint = GetParam(Null,              "ISMasterSvcThumbprint",  "",                "")
  strIsWorkerMaster = GetParam(Null,                  "ISWorkerSvcMaster",  "",                    "")
  strIsWorkerCert   = GetParam(Null,                  "ISWorkerSvcCert",    "",                    "")
  strJobCategory    = GetParam(colGlobal,             "JobCategory",        "",                    "Database Maintenance")
  strStartJobPassword = GetParam(Null,                "StartJobPassword",   "",                    strsaPwd)
  strLabBackup      = GetParam(colStrings,            "LabBackup",          "",                    "Backup")
  strLabBackupAS    = GetParam(colStrings,            "LabBackupAS",        "",                    "AS Backup")
  strLabBPE         = GetParam(colStrings,            "LabBPE",             "",                    "BPE")
  strLabData        = GetParam(colStrings,            "LabData",            "",                    "SQL Data")
  strLabDataAS      = GetParam(colStrings,            "LabDataAS",          "",                    "AS Data")
  strLabDataFS      = GetParam(colStrings,            "LabDataFS",          "",                    "FS Data")
  strLabDataFT      = GetParam(colStrings,            "LabDataFT",          "",                    "FT Data")
  strLabDBA         = GetParam(colStrings,            "LabDBA",             "",                    "DBA Misc")
  strLabDTC         = GetParam(colStrings,            "LabDTC",             "",                    "MSDTC")
  strLabLog         = GetParam(colStrings,            "LabLog",             "",                    "SQL Log")
  strLabLogAS       = GetParam(colStrings,            "LabLogAS",           "",                    "AS Log")
  strLabLogTemp     = GetParam(colStrings,            "LabLogTemp",         "",                    "Temp Log")
  strLabPrefix      = UCase(GetParam(Null,            "LabPrefix",          "",                    ""))
  strLabSysDB       = GetParam(colStrings,            "LabSysDB",           "",                    "SQL SysDB")
  strLabProg        = GetParam(colStrings,            "LabProg",            "",                    "Programs")
  strLabSystem      = GetParam(colStrings,            "LabSystem",          "",                    "System")
  strLabTemp        = GetParam(colStrings,            "LabTemp",            "",                    "SQL Temp")
  strLabTempAS      = GetParam(colStrings,            "LabTempAS",          "",                    "AS Temp")
  strLabTempWin     = GetParam(colStrings,            "LabTempWin",         "",                    "Temp")
  strLanguage       = UCase(GetParam(colStrings,      "Language",           "",                    "ENU"))
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strMailServer     = GetParam(colGlobal,             "MailServer",         "",                    "")
  strMailServerType = ""
  strMainInstance   = GetBuildfileValue("MainInstance")
  strManagementDW   = GetParam(colGlobal,             "ManagementDW",       "",                    "ManagementDW")
  strManagementServer = Ucase(GetParam(colGlobal,     "ManagementServer",   "",                    ""))
  strPath           = Mid(strHKLMFB, 6)
  objWMIReg.GetStringValue strHKLM,strPath,"ManagementServer",strManagementServerRes
  strSQLMaxDop      = GetParam(Null,                  "SQLMaxDop",          "MaxDop",              "")
  strMDSDB          = GetParam(colGlobal,             "MDSDB",              "",                    "MDSData")
  strMDSPort        = GetParam(colGlobal,             "MDSPort",            "",                    "5091")
  strMDSSite        = GetParam(colGlobal,             "MDSSite",            "",                    "MDS")
  strMode           = Ucase(GetParam(Null,            "Mode",               "",                    "QUIET"))
  strMountRoot      = GetParam(colStrings,            "MountRoot",          "",                    "MountPoints")
  strNativeOS       = GetParam(colStrings,            "NativeOS",           "",                    "%")
  strNetNameSource  = UCase(GetParam(colStrings,      "NetNameSource",      "",                    "CLUSTER"))
  strNetworkGUID    = "4D36E972-E325-11CE-BFC1-08002BE10318"
  strNPEnabled      = GetParam(colGlobal,             "NPEnabled",          "",                    "1")
  strNTAuth         = GetBuildfileValue("NTAuth")
  strNTAuthNetwork  = GetBuildfileValue("NTAuthNetwork")
  strNTAuthAccount  = GetParam(colStrings,            "NTAuthAccount",      "",                    "")
  strNTService      = GetParam(colStrings,            "NTService",          "",                    "NT SERVICE")
  strNumErrorLogs   = GetParam(colGlobal,             "NumErrorLogs",       "",                    "31")
  strNumLogins      = GetParam(colGlobal,             "NumLogins",          "",                    "20")
  strNumTF          = GetParam(colStrings,            "NumTF",              "",                    "20")
  strPathPS         = strDirProgSys & "\WindowsPowerShell\Modules"
  strSetLowMemLimit = GetParam(colBuild,              "SetLowMemoryLimit",                    "",                                     "65")
  strSetHardMemLimit  = GetParam(colBuild,            "SetHardMemoryLimit",                   "",                                     "0")
  strSetTotalMemLimit = GetParam(colBuild,            "SetTotalMemoryLimit",                  "",                                     "")
  strSetVertiMemLimit = GetParam(colBuild,            "SetVertiPaqMemoryLimit",               "",                                     "60")
  strSetCLREnabled  = GetParam(colBuild,              "SetCLREnabled",                        "spConfigureCLREnabled",                "1")
  strSetCostThreshold = GetParam(colBuild,            "SetCostThreshold",                     "spConfigureCostThreshold",             "30")
  strSQLMaxMemory   = GetParam(colBuild,              "SQLMaxMemory",                         "spConfigureMaxServerMemory",           "")
  strSQLMinMemory   = GetParam(colBuild,              "SQLMinMemory",                         "",                                     "0")
  strSetMemOptHybridBP      = UCase(GetParam(colBuild,"SetMemoryOptimizedHybridBufferpool",   "",                                     ""))
  strSetMemOptTempdb        = UCase(GetParam(colBuild,"SetMemoryOptimizedTempdbMetadata",     "",                                     ""))
  strSetOptimizeForAdHocWorkloads = GetParam(colBuild,"SetOptimizeForAdHocWorkloads",         "spConfigureOptimizeForAdHocWorkloads", "1")
  strSetRemoteAdminConnections    = GetParam(colBuild,"SetRemoteAdminConnections",            "spConfigureRemoteAdminConnections",    "1")
  strSetRemoteProcTrans     = GetParam(colBuild,      "SetRemoteProcTrans",                   "spConfigureRemoteProcTrans",           "0")
  strSetROLAPDimensionProcessingEffort = UCase(GetParam(colBuild,"SetROLAPDimensionProcessingEffort", "",                             "100000000"))
  strSetxpCmdshell  = GetParam(colBuild,              "SetxpCmdshell",                        "spConfigurexpCmdshell",                "0")
  strOptions        = GetParam(colBuild,              "Options",            "",                    "") 
  strOSRegPath      = GetBuildfileValue("OSRegPath")
  strOSName         = objShell.RegRead("HKLM\" & strOSRegPath & "ProductName")
  strOSVersion      = objShell.RegRead("HKLM\" & strOSRegPath & "CurrentVersion")
  strPassphrase     = GetParam(Null,                  "Passphrase",         "",                    "")
  strOUPath         = GetParam(Null,                  "OUPath",             "",                    "") 
  strPathSQLDefault = "..\SQLMedia"
  strPathSQLMediaOrig = GetParam(colStrings,          "PathSQLMedia",       "",                    "")
  strPathSys        = GetParam(colStrings,            "PathSys",            "",                    strDirSys & "\system32\")
  strPBPortrange    = GetParam(colGlobal,             "PBPortRange",        "",                    "16450-16460")
  strPBScaleout     = UCase(GetParam(colGlobal,       "PBScaleout",         "",                    "True"))
  strPID            = Ucase(GetParam(Null,            "PID",                "PIDKEY",              ""))
  strPowerBIPID     = Ucase(GetParam(colStrings,      "PowerBIPID",         "",                    ""))
  strPreferredOwner = UCase(GetParam(Null,            "PreferredOwner",     "",                    ""))
  strProcArc        = Ucase(objShell.ExpandEnvironmentStrings("%PROCESSOR_ARCHITECTURE%"))
  strProcessId      = GetBuildfileValue("ProcessId")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProfileName    = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
  strProfileName    = Mid(strProfileName, InStrRev(strProfileName, "\") + 1)
  intProcNum        = objShell.ExpandEnvironmentStrings("%NUMBER_OF_PROCESSORS%")
  strProgCacls      = UCase(GetParam(colFiles,        "ProgCacls",          "",                    "cacls"))
  strProgNTRights   = UCase(GetParam(colFiles,        "ProgNtrights",       "",                    "ntrights"))
  strProgSetSPN     = UCase(GetParam(colFiles,        "ProgSetSPN",         "",                    "setspn"))
  strProgReg        = UCase(GetParam(colFiles,        "ProgReg",            "",                    "reg"))
  strReportOnly     = GetBuildfileValue("ReportOnly")
  strResponseNo     = UCase(GetParam(colStrings,      "ResponseNo",         "",                    "N"))
  strResponseYes    = UCase(GetParam(colStrings,      "ResponseYes",        "",                    "Y"))
  strRole           = GetParam(colGlobal,             "Role",               "",                    "")
  strRoleDBANonSA   = GetParam(colStrings,            "RoleDBANonSA",       "",                    "DBA_NonAdmin")
  strRSAlias        = GetParam(Null,                  "RSAlias",            "",                    "")
  strRSName         = GetParam(Null,                  "RSName",             "",                    "Report Server")
  strRsFxVersion    = GetParam(colStrings,            "RsFxVersion",        "",                    "%")
  strRSDBAccount    = GetParam(Null,                  "RSUpgradeDatabaseAccount",              "", "")
  strRSDBPassword   = GetParam(Null,                  "RSUpgradePassword",  "",                    "")
  strRSSQLLocal     = GetParam(colGlobal,             "RSSQLLocal",         "",                    "1")
  strRSVersionNum   = GetParam(colStrings,            "RSVersionNum",       "",                    "%")
  strRunCount       = GetBuildfileValue("RunCount")
  strsaName         = GetParam(colGlobal,             "saName",             "",                    "sa")
  strSecMain        = GetParam(colGlobal,             "SecMain",            "",                    """Administrators"":F ""Users"":R")
  strSecTemp        = GetParam(colGlobal,             "SecTemp",            "",                    """NETWORK SERVICE"":F ""SYSTEM"":F ""Users"":F")
  strSecurityMode   = GetParam(colGlobal,             "SecurityMode",       "",                    "")
  strServer         = GetBuildfileValue("AuditServer")
  strServParm       = UCase(GetParam(Null,            "Server",             "",                    strServer))
  strSetHeaderLength = GetParam(colStrings,           "SetHeaderLength",    "RsHeaderLength",      "65534")
  strSetWorkingSetMaximum = GetParam(colStrings,      "SetWorkingSetMaximum","RsWorkingSetMaximum","")
  strSetupABE       = Ucase(GetParam(colFlags,        "SetupABE",           "InstABE",             "Yes"))
  strSetupAlwaysOn  = Ucase(GetParam(colFlags,        "SetupAlwaysOn",      "",                    "No"))
  strSetupAOAlias   = Ucase(GetParam(colFlags,        "SetupAOAlias",       "",                    ""))
  strSetupAODB      = Ucase(GetParam(colFlags,        "SetupAODB",          "",                    "No"))
  strSetupAOProcs   = Ucase(GetParam(colFlags,        "SetupAOProcs",       "",                    ""))
  strSetupAnalytics = Ucase(GetParam(colFlags,        "SetupAnalytics",     "",                    ""))
  strSetupAPCluster = Ucase(GetParam(colFlags,        "SetupAPCluster",     "",                    "No"))
  strSetupAutoConfig  = Ucase(GetParam(colFlags,      "SetupAutoConfig",    "",                    "Yes"))
  strSetupBIDS      = UCase(GetParam(colFlags,        "SetupBIDS",          "InstBIDS",            "Yes"))
  strSetupBIDSHelper  = Ucase(GetParam(colFlags,      "SetupBIDSHelper",    "InstBIDSHelper",      "Yes"))
  strSetupBPAnalyzer  = Ucase(GetParam(colFlags,      "SetupBPAnalyzer",    "InstBPAnalyzer",      "Yes"))
  strSetupBOL       = UCase(GetParam(colFlags,        "SetupBOL",           "InstBOL",             "Yes"))
  strSetupBPE       = Ucase(GetParam(colFlags,        "SetupBPE",           "",                    "No"))
  strSetupCacheManager  = Ucase(GetParam(colFlags,    "SetupCacheManager",  "InstCacheManager",    "Yes"))
  strSetupClusterShares = Ucase(GetParam(colFlags,    "SetupClusterShares", "",                    "No"))
  strSetupCMD       = UCase(GetParam(colFlags,        "SetupCMD",           "",                    "Yes"))
  strSetupCmdshell  = Ucase(GetParam(colFlags,        "SetupCmdshell",      "ConfigCmdshell",      "No"))
  strSetupCompliance  = Ucase(GetParam(colFlags,      "SetupCompliance",    "",                    "No"))
  strSetupDB2OLE    = Ucase(GetParam(colFlags,        "SetupDB2OLE",        "InstDB2OLE",          "Yes"))
  strSetupDBAManagement = Ucase(GetParam(colFlags,    "SetupDBAManagement", "ConfigDBAManagement", "Yes"))
  strSetupDBMail    = UCase(GetParam(colFlags,        "SetupDBMail",        "ConfigDBMail",        "Yes"))
  strSetupDBOpts    = Ucase(GetParam(colFlags,        "SetupDBOpts",        "ConfigDBOpts",        "Yes"))
  strSetupDCom      = Ucase(GetParam(colFlags,        "SetupDCom",          "ConfigDCom",          "Yes"))
  strSetupDimensionSCD = Ucase(GetParam(colFlags,     "SetupDimensionSCD",  "InstDimensionSCD",    "Yes"))
  strSetupDisableSA = Ucase(GetParam(colFlags,        "SetupDisableSA",     "ConfigDisableSA",     "Yes"))
  strSetupDistributor = Ucase(GetParam(colFlags,      "SetupDistributor",   "",                    "No"))
  strSetupDTSDesigner = Ucase(GetParam(colFlags,      "SetupDTSDesigner",   "InstDTSDesigner",     "Yes"))
  strSetupDTSBackup = Ucase(GetParam(colFlags,        "SetupDTSBackup",     "InstDTSBackup",       "Yes"))
  strSetupNoDriveIndex  = UCase(GetParam(colFlags,    "SetupNoDriveIndex",  "",                    "Yes"))
  strSetupNoSSL3    = UCase(GetParam(colFlags,        "SetupNoSSL3",        "",                    ""))
  strSetupNoTCPNetBios  = UCase(GetParam(colFlags,    "SetupNoTCPNetBios",  "",                    "Yes"))
  strSetupNoTCPOffload  = UCase(GetParam(colFlags,    "SetupNoTCPOffload",  "",                    "Yes"))
  strSetupNoWinGlobal   = UCase(GetParam(colFlags,    "SetupNoWinGlobal",   "",                    "Yes"))
  strSetupDQ        = Ucase(GetParam(colFlags,        "SetupDQ",            "",                    ""))
  strSetupDQC       = Ucase(GetParam(colFlags,        "SetupDQC",           "",                    "Yes"))
  strSetupDRUClt    = Ucase(GetParam(colFlags,        "SetupDRUClt",        "",                    "No"))
  strSetupDRUCtlr   = Ucase(GetParam(colFlags,        "SetupDRUCtlr",       "",                    ""))
  strSetupDTCCID    = Ucase(GetParam(colFlags,        "SetupDTCCID",        "",                    "Yes"))
  strSetupDTCCluster  = Ucase(GetParam(colFlags,      "SetupDTCCluster",    "",                    ""))
  strSetupDTCNetAccess = UCase(GetParam(colFlags,     "SetupDTCNetAccess",  "",                    "No"))
  strSetupFirewall  = Ucase(GetParam(colFlags,        "SetupFirewall",      "",                    "Yes"))
  strSetupFT        = Ucase(GetParam(colFlags,        "SetupFT",            "ConfigFT",            "Yes"))
  strSetupGenMaint  = UCase(GetParam(colFlags,        "SetupGenMaint",      "ConfigGenMaint",      "Yes"))
  strSetupGovernor  = UCase(GetParam(colFlags,        "SetupGovernor",      "",                    ""))
  strSetupIIS       = Ucase(GetParam(colFlags,        "SetupIIS",           "InstIIS",             "No"))
  strSetupIntViewer = Ucase(GetParam(colFlags,        "SetupIntViewer",     "InstIntViewer",       "Yes"))
  strSetupISMaster  = Ucase(GetParam(colFlags,        "SetupISMaster",      "",                    ""))
  strSetupISMasterCluster = Ucase(GetParam(colFlags,  "SetupISMasterCluster", "",                  ""))
  strSetupISWorker  = Ucase(GetParam(colFlags,        "SetupISWorker",      "",                    ""))
  strSetupJavaDBC   = Ucase(GetParam(colFlags,        "SetupJavaDBC",       "InstJavaDBC",         "Yes"))
  strSetupStartJob  = Ucase(GetParam(colFlags,        "SetupStartJob",      "",                    "Yes"))
  strSetupJRE       = Ucase(GetParam(colFlags,        "SetupJRE",           "",                    "No"))
  strSetupKB925336  = UCase(GetParam(colFlags,        "SetupKB925336",      "",                    ""))
  strSetupKB932232  = UCase(GetParam(colFlags,        "SetupKB932232",      "",                    "")) 
  strSetupKB933789  = UCase(GetParam(colFlags,        "SetupKB933789",      "",                    ""))
  strSetupKB937444  = UCase(GetParam(colFlags,        "SetupKB937444",      "",                    ""))
  strSetupKB954961  = UCase(GetParam(colFlags,        "SetupKB954961",      "",                    "")) 
  strSetupKB956250  = UCase(GetParam(colFlags,        "SetupKB956250",      "",                    ""))
  strSetupKB2549864 = UCase(GetParam(colFlags,        "SetupKB2549864",     "",                    ""))
  strSetupKB2781514 = UCase(GetParam(colFlags,        "SetupKB2781514",     "",                    ""))
  strSetupKB2854082 = UCase(GetParam(colFlags,        "SetupKB2854082",     "",                    ""))
  strSetupKB2862966 = UCase(GetParam(colFlags,        "SetupKB2862966",     "",                    ""))
  strSetupKB2919355 = UCase(GetParam(colFlags,        "SetupKB2919355",     "",                    ""))
  strSetupKB2919442 = UCase(GetParam(colFlags,        "SetupKB2919442",     "",                    ""))
  strSetupKB3090973 = UCase(GetParam(colFlags,        "SetupKB3090973",     "",                    ""))
  strSetupKB4019990 = UCase(GetParam(colFlags,        "SetupKB4019990",     "",                    ""))
  strSetupManagementDW = Ucase(GetParam(colFlags,     "SetupManagementDW",  "ConfigManagementDW",  ""))
  strSetupMBCA      = Ucase(GetParam(colFlags,        "SetupMBCA",          "",                    ""))
  strSetupMDS       = Ucase(GetParam(colFlags,        "SetupMDS",           "InstMDS",             ""))
  strSetupMDSC      = Ucase(GetParam(colFlags,        "SetupMDSC",          "",                    ""))
  strSetupMDXStudio = Ucase(GetParam(colFlags,        "SetupMDXStudio",     "InstMDXStudio",       "Yes"))
  strSetupMenus     = UCase(GetParam(colFlags,        "SetupMenus",         "ConfigMenus",         "Yes"))
  strSetupMyDocs    = UCase(GetParam(colFlags,        "SetupMyDocs",        "",                    "Yes"))
  strSetupMSI45     = Ucase(GetParam(colFlags,        "SetupMSI45",         "InstMSI45",           "Yes"))
  strSetupMSMPI     = Ucase(GetParam(colFlags,        "SetupMSMPI",         "",                    ""))
  strSetupNet3      = Ucase(GetParam(colFlags,        "SetupNet3",          "InstNet3",            ""))
  strSetupNet4      = Ucase(GetParam(colFlags,        "SetupNet4",          "InstNet4",            "Yes"))
  strSetupNet4x     = Ucase(GetParam(colFlags,        "SetupNet4x",         "SetupNet45",          "Yes"))
  strSetupNetBind   = UCase(GetParam(colFlags,        "SetupNetBind",       "",                    ""))
  strSetupNetName   = UCase(GetParam(colFlags,        "SetupNetName",       "",                    ""))
  strSetupNetTrust  = UCase(GetParam(colFlags,        "SetupNetTrust",      "",                    "Yes"))
  strSetupNetwork   = Ucase(GetParam(colFlags,        "SetupNetwork",       "ConfigNetwork",       "Yes"))
  strSetupNoDefrag  = UCase(GetParam(colFlags,        "SetupNoDefrag",      "",                    ""))
  strSetupNonSAAccounts = Ucase(GetParam(colFlags,    "SetupNonSAAccounts", "ConfigNonSAAccounts", "Yes"))
  strSetupOLAP      = Ucase(GetParam(colFlags,        "SetupOLAP",          "ConfigOLAP",          "Yes"))
  strSetupOLAPAPI   = Ucase(GetParam(colFlags,        "SetupOLAPAPI",       "ConfigOLAPAPI",       "Yes"))
  strSetupOldAccounts = UCase(GetParam(colFlags,      "SetupOldAccounts",   "ConfigOldAccounts",   "Yes"))
  strSetupParam     = Ucase(GetParam(colFlags,        "SetupParam",         "ConfigParam",         "Yes"))
  strSetupPBM       = Ucase(GetParam(colFlags,        "SetupPBM",           "ConfigPBM",           "Yes"))
  strSetupPDFReader = Ucase(GetParam(colFlags,        "SetupPDFReader",     "InstPDFReader",       "Yes"))
  strSetupPerfDash  = Ucase(GetParam(colFlags,        "SetupPerfDash",      "InstPerfDash",        "Yes"))
  strSetupPlanExpAddin = Ucase(GetParam(colFlags,     "SetupPlanExpAddin",  "InstPlanExpAddin",    "Yes"))
  strSetupPlanExplorer = Ucase(GetParam(colFlags,     "SetupPlanExplorer",  "InstPlanExplorer",    "Yes"))
  strSetupPolyBase  = Ucase(GetParam(colFlags,        "SetupPolyBase",      "",                    ""))
  strSetupProcExp   = Ucase(GetParam(colFlags,        "SetupProcExp",       "InstProcExp",         "Yes"))
  strSetupProcMon   = Ucase(GetParam(colFlags,        "SetupProcMon",       "InstProcMon",         "Yes"))
  strSetupPowerCfg  = Ucase(GetParam(colFlags,        "SetupPowerCfg",      "",                    ""))
  strSetupPS1       = Ucase(GetParam(colFlags,        "SetupPS1",           "",                    ""))
  strSetupPS2       = Ucase(GetParam(colFlags,        "SetupPS2",           "InstPS2",             "Yes"))
  strSetupPowerBI   = Ucase(GetParam(colFlags,        "SetupPowerBI",       "",                    ""))
  strSetupPowerBIDesktop = Ucase(GetParam(colFlags,   "SetupPowerBIDesktop","",                    ""))
  strSetupPSRemote  = Ucase(GetParam(colFlags,        "SetupPSRemote",      "",                    "Yes"))
  strSetupPython    = Ucase(GetParam(colFlags,        "SetupPython",        "",                    ""))
  strSetupRawReader = Ucase(GetParam(colFlags,        "SetupRawReader",     "InstRawReader",       "Yes"))
  strSetupReportViewer = Ucase(GetParam(colFlags,     "SetupReportViewer",  "",                    ""))
  strSetupRMLTools  = Ucase(GetParam(colFlags,        "SetupRMLTools",      "InstRMLTools",        "Yes"))
  strSetupRptTaskPad  = Ucase(GetParam(colFlags,      "SetupRptTaskPad",    "InstRptTaskPad",      "Yes"))
  strSetupRServer   = Ucase(GetParam(colFlags,        "SetupRServer",       "",                    ""))
  strSetupRSAT      = UCase(GetParam(colFlags,        "SetupRSAT",          "",                    ""))
  strSetupRSAdmin   = Ucase(GetParam(colFlags,        "SetupRSAdmin",       "ConfigRSAdmin",       "Yes"))
  strSetupRSExec    = Ucase(GetParam(colFlags,        "SetupRSExec",        "ConfigRSExec",        "Yes"))
  strSetupRSAlias   = Ucase(GetParam(colFlags,        "SetupRSAlias",       "",                    "No"))
  strSetupRSIndexes = Ucase(GetParam(colFlags,        "SetupRSIndexes",     "",                    ""))
  strSetupRSKeepAlive = Ucase(GetParam(colFlags,      "SetupRSKeepAlive",   "",                    ""))
  strSetupRSLinkGen = Ucase(GetParam(colFlags,        "SetupRSLinkGen",     "",                    "Yes"))
  strSetupRSScripter  = Ucase(GetParam(colFlags,      "SetupRSScripter",    "InstRSScripter",      "Yes"))
  strSetupSAAccounts  = Ucase(GetParam(colFlags,      "SetupSAAccounts",    "ConfigSAAccounts",    "Yes"))
  strSetupSAPassword  = Ucase(GetParam(colFlags,      "SetupSAPassword",    "ConfigSAPassword",    "Yes"))
  strSetupSamples   = Ucase(GetParam(colFlags,        "SetupSamples",       "InstSamples",         "No"))
  strSetupSemantics = Ucase(GetParam(colFlags,        "SetupSemantics",     "InstSemantics",       "Yes"))
  strSetupServices  = Ucase(GetParam(colFlags,        "SetupServices",      "ConfigServices",      "Yes"))
  strSetupServiceRights = Ucase(GetParam(colFlags,    "SetupServiceRights", "",                    "Yes"))
  strSetupShares    = Ucase(GetParam(colFlags,        "SetupShares",        "",                    "Yes"))
  strSetupSlipstream  = Ucase(GetParam(colFlags,      "SetupSlipstream",    "",                    "Yes"))
  strSetupSnapshot  = Ucase(GetParam(colFlags,        "SetupSnapshot",      "",                    "Yes"))
  strSetupSP        = UCase(GetParam(colFlags,        "SetupSP",            "InstSP",              "Yes"))
  strSetupSPCU      = UCase(GetParam(colFlags,        "SetupSPCU",          "InstSPCU",            "Yes"))
  strSetupSPCUSNAC  = UCase(GetParam(colFlags,        "SetupSPCUSNAC",      "InstSPCUSNAC",        "Yes"))
  strSetupKerberos  = Ucase(GetParam(colFlags,        "SetupKerberos",      "SetupSPN",            "Yes"))
  strSetupSQLAgent  = Ucase(GetParam(colFlags,        "SetupSQLAgent",      "ConfigSQLAgent",      "Yes"))
  strSetupSQLAS     = Ucase(GetParam(colFlags,        "SetupSQLAS",         "InstSQLAS",           ""))
  strSetupSQLASCluster = Ucase(GetParam(Null,         "SetupSQLASCluster",  "",                    ""))
  strSetupSQLBC     = UCase(GetParam(colFlags,        "setupSQLBC",         "InstSQLBC",           "Yes"))
  strSetupSQLCE     = UCase(GetParam(colFlags,        "SetupSQLCE",         "",                    ""))
  strSetupSQLDB     = Ucase(GetParam(colFlags,        "SetupSQLDB",         "InstSQLDB",           ""))
  strSetupSQLDBCluster = Ucase(GetParam(Null,         "SetupSQLDBCluster",  "",                    ""))
  strSetupSQLDBAG   = Ucase(GetParam(colFlags,        "SetupSQLDBAG",       "",                    "Yes"))
  strSetupSQLDBRepl = Ucase(GetParam(colFlags,        "SetupSQLDBRepl",     "InstSQLDBRepl",       "Yes"))
  strSetupSQLDBFS   = UCase(GetParam(colFlags,        "SetupSQLDBFS",       "",                    "Yes"))
  strSetupSQLDBFT   = Ucase(GetParam(colFlags,        "SetupSQLDBFT",       "InstSQLDBFT",         "Yes"))
  strSetupSQLDebug  = UCase(GetParam(colFlags,        "SetupSQLDebug",      "",                    ""))
  strSetupSQLInst   = Ucase(GetParam(colFlags,        "SetupSQLInst",       "ConfigSQLInst",       "Yes"))
  strSetupSQLIS     = Ucase(GetParam(colFlags,        "SetupSQLIS",         "InstSQLIS",           ""))
  strSetupSQLMail   = UCase(GetParam(colFlags,        "SetupSQLMail",       "ConfigSQLMail",       "No"))
  strSetupSQLNexus  = Ucase(GetParam(colFlags,        "SetupSQLNexus",      "InstSQLNexus",        "Yes"))
  strSetupSQLNS     = Ucase(GetParam(colFlags,        "SetupSQLNS",         "InstSQLNS",           "No"))
  strSetupSQLPowershell = UCase(GetParam(colFlags,    "SetupSQLPowershell", "",                    "Yes"))
  strSetupSQLRS     = Ucase(GetParam(colFlags,        "SetupSQLRS",         "InstSQLRS",           ""))
  strSetupSQLRSCluster = Ucase(GetParam(colFlags,     "SetupSQLRSCluster",  "",                    ""))
  strSetupSQLServer = Ucase(GetParam(colFlags,        "SetupSQLServer",     "ConfigSQLServer",     "Yes"))
  strSetupSSL       = Ucase(GetParam(colFlags,        "SetupSSL",           "",                    "No"))
  strSetupSQLTools  = Ucase(GetParam(colFlags,        "SetupSQLTools",      "InstSQLTools",        "Yes"))
  strSetupSSDTBI    = Ucase(GetParam(colFlags,        "SetupSSDTBI",        "SetupSSDT",           ""))
  strSetupSSMS      = Ucase(GetParam(colFlags,        "SetupSSMS",          "InstSSMS",            "Yes"))
  strSetupSSISCluster = Ucase(GetParam(colFlags,      "SetupSSISCluster",   "SetupSQLISCluster",   "No"))
  strSetupSSISDB    = Ucase(GetParam(colFlags,        "SetupSSISDB",        "",                    "Yes"))
  strSetupStdAccounts = Ucase(GetParam(colFlags,      "SetupStdAccounts",   "ConfigStdAccounts",   "Yes"))
  strSetupStreamInsight = Ucase(GetParam(colFlags,    "SetupStreamInsight", "InstStreamInsight",   "No"))
  strSetupStretch   = Ucase(GetParam(colFlags,        "SetupStretch",       "",                    "No"))
  strSetupSysDB     = Ucase(GetParam(colFlags,        "SetupSysDB",         "ConfigSysDB",         "Yes"))
  strSetupSysIndex  = Ucase(GetParam(colFlags,        "SetupSysIndex",      "ConfigSysIndex",      "Yes"))
  strSetupSysManagement = Ucase(GetParam(colFlags,    "SetupSysManagement", "ConfigSysManagement", "Yes"))
  strSetupSystemViews   = Ucase(GetParam(colFlags,    "SetupSystemViews",   "InstSystemViews",     "Yes"))
  strSetupTelemetry = Ucase(GetParam(colFlags,        "SetupTelemetry",     "",                    ""))
  strSetupTempDb    = Ucase(GetParam(colFlags,        "SetupTempDb",        "",                    "Yes"))
  strSetupTempWin   = Ucase(GetParam(colFlags,        "SetupTempWin",       "",                    "Yes"))
  strSetupTLS12     = Ucase(GetParam(colFlags,        "SetupTLS12",         "",                    ""))
  strSetupTrouble   = Ucase(GetParam(colFlags,        "SetupTrouble",       "InstTrouble",         "Yes"))
  strSetupVC2010    = UCase(GetParam(colFlags,        "SetupVC2010",        "",                    ""))
  strSetupVS        = UCase(GetParam(colFlags,        "SetupVS",            "",                    "Yes"))
  strSetupVS2005SP1 = Ucase(GetParam(colFlags,        "SetupVS2005SP1",     "InstVS2005SP1",       ""))
  strSetupVS2010SP1 = Ucase(GetParam(colFlags,        "SetupVS2010SP1",     "",                    ""))
  strSetupWindows   = UCase(GetParam(colFlags,        "SetupWindows",       "",                    "Yes"))
  strSetupWinAudit  = UCase(GetParam(colFlags,        "SetupWinAudit",      "",                    ""))
  strSetupXEvents   = Ucase(GetParam(colFlags,        "SetupXEvents",       "InstXEvents",         "Yes"))
  strSetupXMLNotepad  = Ucase(GetParam(colFlags,      "SetupXMLNotepad",    "InstXMLNotepad",      "Yes"))
  strSetupZoomIt    = Ucase(GetParam(colFlags,        "SetupZoomIt",        "InstZoomIt",          "Yes"))
  intSpeedTest      = GetParam(colGlobal,             "SpeedTest",          "",                    "5.0")
  strSQLAdminAccounts  = GetParam(Null,               "SQLSysadminAccounts", "",                   "")
  strSQLJavaDir     = GetParam(Null,                  "SQLJavaDir",         "",                    "")
  strSQLOperator    = GetParam(Null,                  "SQLOperator",        "",                    "SQL Alerts")
  strSQLTempdbFileCount = GetParam(colGlobal,         "SQLTempdbFileCount", "",                    "")
  strSSASAdminAccounts = GetParam(Null,               "ASSysadminAccounts", "",                    "")
  strClusterAOFound = ""
  strClusterASFound = ""
  strClusterSQLFound  = ""
  strPath           = Mid(strHKLMFB, 6)
  objWMIReg.GetStringValue strHKLM,strPath,"DTCClusterRes",strDTCClusterRes
  objWMIReg.GetStringValue strHKLM,strPath,"SetupDTCNetAccessStatus",strSetupDTCNetAccessStatus
  strSPLevel        = UCase(GetParam(colGlobal,       "SPLevel",            "",                    "RTM"))
  strSPCULevel      = UCase(GetParam(colGlobal,       "SPCULevel",          "",                    ""))
  strSQLAgentStart  = GetParam(colStrings,            "SQLAgentStart",      "",                    "SQLSERVERAGENT starting under Windows NT service control")
  strSQLList        = "SQL2005 SQL2008 SQL2008R2 SQL2012 SQL2014 SQL2016 SQL2017 SQL2019"
  strSQLLogReinit   = GetParam(colGlobal,             "SQLLogReinit",       "",                    "The error log has been reinitialized")
  strSQLProgDir     = GetParam(colStrings,            "SQLProgDir",         "",                    "Microsoft SQL Server")
  strSQLRecoveryComplete = GetParam(colStrings,       "SQLRecoveryComplete","",                    "Recovery is complete")
  strSQMReporting   = GetParam(colGlobal,             "SQMReporting",       "",                    "0")
  strSQLSharedMR    = GetParam(colStrings,            "SQLSharedMR",        "SQL_Shared_MR",       "YES")
  strSQLRSStart     = GetParam(colStrings,            "SQLRSStartComplete", "",                    "INFO: Total Physical memory:")
  strSQLVersionNet  = GetParam(colStrings,            "SQLVersionNet",      "",                    "%")
  strSQLVersionNum  = GetParam(colStrings,            "SQLVersionNum",      "",                    "%")
  strSQLVersionWMI  = GetParam(colStrings,            "SQLVersionWMI",      "",                    "%")
  strSSISDB         = Ucase(GetParam(colStrings,      "SSISDB",             "",                    "SSISDB"))
  strSSISPassword   = GetParam(Null,                  "SSISPassword",       "",                    strsaPwd)
  strSSISRetention  = GetParam(colStrings,            "SSISRetention",      "",                    "30")
  strStatusAssumed  = GetParam(colStrings,            "StatusAssumed",      "",                    "Assumed")
  strStatusBypassed = " " & GetParam(colStrings,      "StatusBypassed",     "",                    "Bypassed")
  strStatusComplete = " " & GetParam(colStrings,      "StatusComplete",     "",                    "Complete")
  strStatusFail     = " " & GetParam(colStrings,      "StatusFail",         "",                    "Install Failed")
  strStatusManual   = " " & GetParam(colStrings,      "StatusManual",       "",                    "Configure Manually")
  strStatusPreConfig  = " " & GetParam(colStrings,    "StatusPreConfig",    "",                    "Already Configured")
  strStatusProgress = " " & GetParam(colStrings,      "StatusProgress",     "",                    "In Progress")
  strStopAt         = Ucase(GetParam(Null,            "StopAt",             "",                    ""))
  strStreamInsightPID = GetParam(colStrings,          "StreamInsightPID",   "",                    "")
  strTallyCount     = GetParam(colStrings,            "TallyCount",         "",                    "1000000")
  strtempdbFile     = GetParam(colGlobal,             "SqlTempdbFileSize",  "tempdbFile",          "200 MB")
  strtempdbLogFile  = GetParam(colGlobal,             "SqlTempdbLogFileSize", "",                  "50 MB")
  strTCPEnabled     = GetParam(colGlobal,             "TCPEnabled",         "",                    "1")
  strTCPPort        = GetParam(colGlobal,             "TCPPort",            "",                    "1433")
  strTCPPortAO      = GetParam(colGlobal,             "TCPPortAO",          "",                    "5022")
  strTCPPortAS      = GetParam(colGlobal,             "TCPPortAS",          "",                    "2383")
  strTCPPortDAC     = GetParam(colGlobal,             "TCPPortDAC",         "",                    "1434")
  strTCPPortDebug   = GetParam(colGlobal,             "TCPPortDebug",       "",                    "")
  strTCPPortDTC     = GetParam(colGlobal,             "TCPPortDTC",         "",                    "13300")
  strTCPPortRS      = GetParam(colGlobal,             "TCPPortRS",          "",                    "80")
  strUCServer       = UCase(strServer)
  strUnknown        = "Unknown"
  strUpdateSource   = GetParam(Null,                  "UpdateSource",       "",                    "")
  strUseFreeSSMS    = UCase(GetParam(colGlobal,       "UseFreeSSMS",        "",                    ""))
  strUserDNSDomain  = objShell.ExpandEnvironmentStrings("%UserDNSDomain%")
  strUserAdmin      = GetBuildfileValue("UserAdmin")
  strUserName       = GetBuildfileValue("AuditUser")
  strUserProfile    = objShell.ExpandEnvironmentStrings("%USERPROFILE%")
  strUserSID        = GetBuildfileValue("UserSID")
  strUserConfiguration    = Ucase(GetParam(colFlags,  "UserConfiguration",  "",                    "Yes"))
  strUserConfigurationvbs = GetParam(colFiles,        "UserConfigurationvbs","",                   "User2Configuration.vbs")
  strUserReg        = "HKCU\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders\"
  strUserDTop       = objShell.RegRead(strUserReg & "Desktop")
  strUserProf       = objShell.RegRead(strUserReg & "Start Menu")
  strUserPreparation      = Ucase(GetParam(colFlags,  "UserPreparation",    "",                    "Yes"))
  strUserPreparationvbs   = GetParam(colFiles,        "UserPreparationvbs", "",                    "User1Preparation.vbs")
  strUseSysDB       = GetParam(Null,                  "UseSysDB",           "",                    "")
  strValidate       = GetParam(Null,                  "Validate",           "",                    "YES")
  strVersionFB      = GetBuildfileValue("VersionFB")
  strVolErrorList   = ""
  strVSVersionNum   = GetParam(colStrings,            "VSVersionNum",       "",                    "%")
  strVSVersionPath  = GetParam(colStrings,            "VSVersionPath",      "",                    "%")
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitMed        = GetBuildfileValue("WaitMed")
  strWaitShort      = GetBuildfileValue("WaitShort")

  Select Case True
    Case Instr("DATA CENTER ENTERPRISE EVALUATION DEVELOPER", strEdition) > 0
      strEditionEnt  = "YES"
    Case Else
      strEditionEnt  = ""
  End Select

  Select Case True
    Case strSQLVersion < "SQL2019"
      ' Nothing
    Case colArgs.Exists("SQL_Inst_Java")
      strSetupJRE = "YES"
  End Select

  If colArgs.Exists("X86") Then
    strWOWX86       = "TRUE"
  End If

  Select Case True
    Case strSetupAnalytics = "N/A"
      strSetupPython  = "N/A"
      strSetupRServer = "N/A"
    Case strSetupAnalytics = "NO"
      strSetupPython  = "N/A"
      strSetupRServer = "N/A"
    Case strSetupPython = "YES"
      strSetupAnalytics = "YES"
    Case strSetupRServer = "YES"
      strSetupAnalytics = "YES"
    Case strSQLVersion >= "SQL2017"
      ' Nothing
    Case strSetupAnalytics = "YES"
      strSetupRServer = "YES"
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2019"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case Else
      Call SetParam("SQLSharedMR",       strSQLSharedMR,           "NO",  "Analytics must be installed as separate instance for SQL Cluster", "")
  End Select

  If strFineBuildStatus = "" Then
    strFineBuildStatus = strStatusProgress
  End If

  strPath           = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\PackageDetect\Microsoft-Windows-Common-Foundation-Package~31bf3856ad364e35~" & LCase(strProcArc) & "~~0.0.0.0\"
  objWMIReg.GetDwordValue strHKLM,strPath,"Package_for_KB2919355~31bf3856ad364e35~" & LCase(strProcArc) & "~~6.3.1.14",strStatusKB2919355
  Select Case True
    Case IsNull(strStatusKB2919355)
      strStatusKB2919355 = ""
    Case Else
      strStatusKB2919355 = CStr(strStatusKB2919355)
  End Select

  objWMIReg.GetStringValue strHKLM,"SYSTEM\CurrentControlSet\Services\W3SVC\","ObjectName",strIISAccount
  Select Case True
    Case IsNull(strIISAccount)
      strIISAccount = ""
  End Select

  Call SetDomainDetails()

  If Not colArgs.Exists("Edition") Then
    strEdType       = strStatusAssumed
  End If

  strRebootLoop     = GetBuildfileValue("RebootLoop") 
  If strRebootLoop = "" Then
    strRebootLoop   = "0"
  End If

  Select Case True
    Case strProcArc = "X86"
      strFileArc     = "X86"
    Case Else
      strFileArc     = "X64"
  End Select

  Select Case True
    Case strSetupSSL = "YES"
      strHTTP       = "https"
    Case Else
      strHTTP       = "http"
  End Select

  strPath           = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v3.5\"
  objWMIReg.GetStringValue strHKLM,strPath,"Version",strVersionNet3
  If IsNull(strVersionNet3) Then
    strVersionNet3  = ""
  End If

  strPath           = "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\"
  objWMIReg.GetStringValue strHKLM,strPath,"Version",strVersionNet4
  If IsNull(strVersionNet4) Then
    strVersionNet4  = ""
  End If

  Select Case True
    Case (strSQLVersion = "SQL2005") And (strFileArc = "X86")
      strPathCScript  = "CSCRIPT"
      strRegasmExe    = "%COMSPEC% /D /C ""%WINDIR%\Microsoft.Net\Framework\v2.0.50727\regasm.exe"" "
    Case strSQLVersion = "SQL2005"
      strPathCScript = strDirSys & "\SysWOW64\CSCRIPT.EXE"
      strRegasmExe    = "%COMSPEC% /D /C ""%WINDIR%\Microsoft.Net\Framework\v2.0.50727\regasm.exe"" "
    Case (strFileArc = "X86") Or (strWOWX86 = "TRUE")
      strPathCScript = strDirSys & "\SysWOW64\CSCRIPT.EXE"
      strRegasmExe    = "%COMSPEC% /D /C ""%WINDIR%\Microsoft.Net\Framework\v2.0.50727\regasm.exe"" "
    Case strOSVersion >= "6.2"
      strPathCScript  = "CSCRIPT"
      strRegasmExe    = "%COMSPEC% /D /C ""%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe"" "
    Case Else
      strPathCScript  = "CSCRIPT"
      strRegasmExe    = "%COMSPEC% /D /C ""%WINDIR%\Microsoft.Net\Framework\v2.0.50727\regasm.exe"" "
  End Select

  objWMIReg.GetStringValue strHKLM,"Cluster\","SharedVolumesRootBase",strCSVRoot
  Select Case True
    Case IsNull(strCSVRoot)
      strCSVRoot    = "Unknown"
    Case Right(strCSVRoot, 1) <> "\"
      strCSVRoot    = strCSVRoot & "\"
  End Select
  strCSVRoot        = UCase(strCSVRoot)

  Select Case True
    Case strSQLVersion <= "SQL2014"
      strReportViewerVersion = GetParam(colStrings,   "ReportViewerVersion","",                    "10.0.40219.1")
    Case Else
      strReportViewerVersion = GetParam(colStrings,   "ReportViewerVersion","",                    "12.0.2402.15")
  End Select

  Select Case True
    Case strRunCount = ""
      Call SetBuildfileValue("RunCount", "1")
    Case Else
      Call SetBuildfileValue("RunCount", CStr(CInt(strRunCount) + 1))
  End Select

  Call SetBuildfileValue("SQLVersion",              strSQLVersion)
  Call SetBuildfileValue("SQLVersionNum",           strSQLVersionNum)
  Call SetBuildfileValue("WOWX86",                  strWOWX86)

  Call DebugLog("Current ProcessId: " & strProcessId)

End Sub


Sub SetDomainDetails()
  Call DebugLog("SetDomainDetails:")

  Select Case True
    Case UCase(strDomain) = strUCServer
      strADRoot     = ""
    Case Else
      Set objADRoot = GetObject("LDAP://RootDSE")
      strADRoot     = "LDAP://" & objADRoot.Get("defaultNamingContext") 
  End Select

  If strUserDNSDomain = "%UserDNSDomain%" Then
    strUserDNSDomain = ""
  End If
  strUserDNSDomain  = UCase(strUserDNSDomain)
  Call SetBuildfileValue("UserDNSDomain",           strUserDNSDomain)

  strLocalDomain    = strServer
  Select Case True
    Case strUserDNSDomain = ""
      ' Nothing
    Case Else
      strDC         = objSysInfo.GetAnyDCName
      strDC         = Left(strDC, Instr(strDC, ".") - 1)  
      If strServer = strDC Then
        strLocalDomain = strDomain
      End If
  End Select

End Sub


Function GetParam(colParam, strParam, strAltParam, strDefault) 
  Call DebugLog("GetParam: " & strParam)
  Dim strParamAccount, strParamPassword, strParamPath, strParamNonDefault, strParamVol, strValue, strValueDefault

  Select Case True
    Case (strType = "REBUILD") And (UCase(Left(strParam, 5)) = UCase("Setup"))
      strValueDefault = "NO"
    Case strDefault = "%"
      strValueDefault = ""
    Case Else
      strValueDefault = strDefault
  End Select

  strDebugMsg1      = "Find parameter value in XML configuration file"
  Select Case True
    Case (strType = "REBUILD") And (UCase(Left(strParam, 5)) = UCase("Setup"))
      strValue      = "NO"
    Case IsNull(colParam)
      strValue      = strValueDefault
    Case IsNull(colParam.getAttribute(strParam))
      strValue      = strValueDefault
    Case Else
      strValue      = colParam.getAttribute(strParam)
  End Select

  strDebugMsg1      = "Apply any parameter overide from Alternative parameter"
  Select Case True
    Case strAltParam = ""
      ' Nothing
    Case Not colArgs.Exists(strAltParam)
      ' Nothing
    Case Else
      strValue      = colArgs.Item(strAltParam)
  End Select

  strDebugMsg1      = "Apply any parameter overide from CSCRIPT arguments"
  Select Case True
    Case Not colArgs.Exists(strParam)
      ' Nothing
    Case Else
      strValue      = colArgs.Item(strParam)
  End Select

  strDebugMsg1      = "Validate parameter value"

  Select Case True ' Validate SETUP parameters
    Case Left(strParam, 5) <> "Setup"
      ' Nothing
    Case strValue = ""
      ' Nothing
    Case InStr(" YES NO N/A ", UCase(strValue)) = 0
      Call SetBuildMessage(strMsgErrorConfig, "/" & strParam & ": value must be one of 'Yes', 'No', 'N/A'")
  End Select

  strDebugMsg1      = "Build lists of parameter types"

  strParamAccount   = ""
  Select Case True
    Case Instr(strParam, "Account") > 0
      strParamAccount = strParam
    Case Instr(strParam, "Acct") > 0
      strParamAccount = strParam
    Case Instr(strParam, "Group") > 0
      strParamAccount = strParam
  End Select
  Select Case True
    Case strParamAccount = ""
      ' Nothing
    Case (strSQLVersion > "SQL2005") And (strAltParam <> "")    
      Call ParamListAdd("ListAccount", strAltParam)
    Case Else
      Call ParamListAdd("ListAccount", strParamAccount)
  End Select

  strParamPassword  = ""
  Select Case True
    Case Instr(strParam, "Password") > 0
      strParamPassword = strParam
    Case Instr(strParam, "Passphrase") > 0
      strParamPassword = strParam
    Case Instr(strParam, "Pwd") > 0
      strParamPassword = strParam
    Case Right(strParam, 3) = "PID"
      strParamPassword = strParam
    Case Instr(strParam, "Thumbprint") > 0
      strParamPassword = strParam
  End Select
  Select Case True
    Case strParamPassword = ""
      ' Nothing
    Case (strSQLVersion > "SQL2005") And (strAltParam <> "")    
      Call ParamListAdd("ListPassword", strAltParam)
    Case Else
      Call ParamListAdd("ListPassword", strParamPassword)
  End Select

  strParamPath      = ""
  Select Case True
    Case Instr(strParam, "Path") > 0
      strParamPath  = strParam
    Case Instr(strParam, "Dir") > 0
      strParamPath  = strParam
  End Select
  If strParamPath <> "" Then
    Call ParamListAdd("ListPath", strParamPath)
  End If

  strParamVol       = ""
  Select Case True
    Case Instr(strParam, "VolSize") > 0
      ' Nothing
    Case Instr(strParam, "VolSpace") > 0
      ' Nothing
    Case Instr(strParam, "Vol") > 0
      strParamVol = strParam
  End Select
  If strParamVol <> "" Then
    Call ParamListAdd("ListVol", strParamVol)
  End If

  strParamNonDefault = ""
  Select Case True
    Case Left(strParam, 5) = "Setup"
      ' Nothing
    Case Left(strParam, 4) = "Menu"
      ' Nothing
    Case strParam = strParamAccount
      ' Nothing
    Case strParam = strParamPassword
      ' Nothing
    Case strParam = strParamPath
      ' Nothing
    Case strParam = strParamVol
      ' Nothing
    Case UCase(strValue) = UCase(strValueDefault)
      ' Nothing
    Case strDefault = "%"
      ' Nothing
    Case Else
      strParamNonDefault = strParam
  End Select
  If strParamNonDefault <> "" Then
    Call ParamListAdd("ListNonDefault", strParamNonDefault)
  End If

  GetParam    = strValue

End Function


Sub ParamListAdd(strList, strParam)
  Call DebugLog("ParamListAdd: " & strList & ", " & strParam)
  Dim strListData

  strListData       = GetBuildfileValue(strList)
  Select Case True
    Case strParam = ""
      ' Nothing
    Case Instr(" " & strListData & " ", " " & strParam & " ") > 0
      ' Nothing
    Case Else
      Call SetBuildfileValue(strList, strListData & strParam & " ")
  End Select

End Sub


Sub Process()
  Call SetProcessId("0","Prepare FineBuild Configuration (FBConfigBuild.vbs)")

  Call SetupBuild()

  Call GetBuildfileData()

  Call SetBuildfileData()

  Call CheckUtils()

  Call FineBuild_Validate()

End Sub


Sub SetupBuild()
  Call SetProcessId("0A", "Setup FineBuild Environment")

  Call SetOSFlags()

  Call SetPrimaryFlags()

  Call LogEnvironment()

  Call CheckBootTime()

  Call GetDNSDetails()

  If CheckClusterHost() = "YES" Then
    Call OpenCluster()
  End If

  Call SetSetupFlags()

  Call EarlyValidate()

  Call CheckRebootStatus()

End Sub


Sub SetOSFlags()
  Call SetProcessId("0AA", "Set flags for Windows OS")
  Dim colComputer
  Dim objComputer

  objWMIReg.GetStringValue strHKLM,strOSRegPath,"InstallationType",strOSType
  Select Case True
    Case strOSType > ""
      ' Nothing
    Case Instr(LCase(strOSName), "windows server") > 0
      strOSType     = "Server"
    Case Instr(LCase(strOSName), "windows xp") > 0
      strOSType     = "Client"
    Case Instr(LCase(strOSName), "windows vista") > 0
      strOSType     = "Client"
  End Select
  strOSType         = Ucase(strOSType)

  objWMIReg.GetStringValue strHKLM,strOSRegPath,"CurrentBuild",strOSBuild
  Select Case True
    Case strOSVersion <> "6.3"
      ' Nothing
    Case Instr(LCase(strOSName), "windows 10") > 0
      strOSVersion = "6.3A"
    Case Instr(LCase(strOSName), "windows server 2016") > 0
      strOSVersion = "6.3A"
    Case Instr(LCase(strOSName), "windows server 2019") > 0
      strOSVersion = "6.3B"
  End Select

  Select Case True
    Case strOSVersion < "6.3"
      strAutoLogonCount = "1"
    Case Else
      strAutoLogonCount = "2"
  End Select

  objWMIReg.GetStringValue strHKLM,strOSRegPath,"CSDVersion",strOSLevel
  Select Case True
    Case IsNull(strOSLevel) 
      strOSLevel    = "RTM"
    Case strOSLevel = "" 
      strOSLevel    = "RTM"
  End Select

  Select Case True
    Case strOSType <> "CLIENT"
      strCmdPS      = "POWERSHELL "
    Case strOSVersion < "6"
      strCmdPS      = strPathSys & "WindowsPowershell\V1.0\POWERSHELL.exe "
    Case Else
      strCmdPS      = "POWERSHELL "
  End Select
  Call SetBuildfileValue("CmdPS", strCmdPS)

  strCmd            = "SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"
  Set colComputer   = objWMI.ExecQuery (strCmd)
  For Each objComputer in colComputer
    strServerMB     = CLng(objComputer.TotalPhysicalMemory / (1024 * 1024))
  Next

End Sub


Sub SetPrimaryFlags()
  Call SetProcessId("0AB", "Set Primary Install Flags")
  Dim arrNames, arrTypes

  Select Case True
    Case strType = "UPGRADE"
      strAction     = Ucase(GetParam(Null,            "Action",             "",                    "UPGRADE"))
    Case Else
      strAction     = Ucase(GetParam(Null,            "Action",             "",                    "INSTALL"))
  End Select

  Select Case True
    Case strAction = strActionClusInst
      strClusterAction  = strAction
    Case strAction = "ADDNODE"
      strClusterAction  = strAction
    Case Else
      strClusterAction = ""
  End Select

  strPath           = Mid(strHKLMSQL, 6) & "Instance Names\SQL"
  objWMIReg.EnumValues strHKLM, strPath, arrNames, arrTypes
  Select Case True
    Case strMainInstance <> ""
      ' Nothing
    Case IsNull(arrNames)
      strMainInstance = "YES"
    Case Ubound(arrNames) > 0
      strMainInstance = "NO"
    Case arrNames(0) = strInstance
      strMainInstance = "YES"
    Case Else
      strMainInstance = "NO"
  End Select

  strPathAddCompOrig  = GetParam(colStrings,          "PathAddComp",     "",                    "..\Additional Components")
  strPathAddComp      = GetMediaPath(strPathAddCompOrig)
  Select Case True
    Case strPathAddComp <> ""
      ' Nothing
    Case strPathAddCompOrig <> "..\Additional Components"
      ' Nothing
    Case Else
      strPathAddComp = GetMediaPath("..\..\Additional Components")
  End Select

  Call SetupStatefile()

End Sub


Sub SetupStatefile()
  Call DebugLog("SetupStatefile:")
  Dim objElement, objRoot
  Dim strFBPathLocal, strPathStatefile

  strFBPathLocal    = GetBuildfileValue("FBPathLocal")
  strPathFBStart    = GetBuildfileValue("PathFBStart")
  Call ResetMediaPath("PathFBStart", strFBPathLocal, Left(strPathFBStart, 2))
  strPathFBStart    = GetBuildfileValue("PathFBStart")

  Select Case True
    Case FormatFolder(strPathFBStart) <> FormatFolder(strPathFB)
      strPathStatefile = strPathFBStart
    Case strPathAddComp = ""
      strPathStatefile = strPathFB
    Case objFSO.FolderExists(FormatFolder(strPathAddComp))
      strPathStatefile = strPathAddComp
    Case Else
      strPathStatefile = strPathFB
  End Select

  strStatefile      = FormatFolder(strPathStatefile) & "FineBuildState.xml"
  strDebugMsg1      = "Statefile: " & strStatefile
  If Not objFSO.FileExists(strStatefile) Then
    Set objRoot     = objStatefile.createElement("FineBuild")
    objStatefile.appendChild objRoot
    Set objElement     = objStatefile.createProcessingInstruction("xml", "version=""1.0""  encoding=""utf-8""")
    objStatefile.insertBefore objElement, objStatefile.childNodes.item(0)
    Set objElement  = objStatefile.createElement("FineBuildState")
    objRoot.appendChild objElement
    objStatefile.save strStatefile
  End If

  Call SetBuildfileValue("Statefile",          strStatefile)
  
End Sub


Sub LogEnvironment()
  Call SetProcessId("0AC", "Log FineBuild Environment")

  Call SetBuildfileValue("OSName",             strOSName)
  Call SetBuildfileValue("OSLevel",            strOSLevel)
  Call SetBuildfileValue("FileArc",            strFileArc)
  Call SetBuildfileValue("Instance",           strInstance)

  Call FBLog("**************************************************")
  Call FBLog("*")
  Call FBLog("* SQL FineBuild Version         " & strVersionFB)
  Call FBLog("* Server Name                   " & strServer)
  Call FBLog("* Operating System Name         " & strOSName)
  Call FBLog("* Operating System Level        " & strOSLevel)
  Call FBLog("* Operating System Platform     " & strFileArc)
  Call FBLog("* SQL Server Version            " & strSQLVersion)
  Call FBLog("* SQL Server Instance           " & strInstance)
  Call FBLog("* SQL Server Edition            " & strEdition)
  Call FBLog("* FineBuild run on              " & GetStdDate(""))
  Call FBLog("*")
  Call FBLog("**************************************************")

End Sub


Sub CheckBootTime()
  Call SetProcessId("0AD", "Check Time since Reboot")
  Dim colOS
  Dim objOS
  Dim intDelay
  Dim strAccount, strBootUpTime, strBuildfileTime, strClusService, strClusStatus

  strBuildfileTime  = GetBuildfileValue("BuildFileTime")
  Select Case True
    Case strClusterAction = "" ' If not Cluster Install then only delay long enough for services to start
      intTimer      = 30
    Case strBuildfileTime = ""
      intTimer      = 350
    Case CLng(strBuildfileTime) < 0 ' W2016 CTP sometimes loses track of time
      intTimer      = 350
    Case CLng(strBuildfileTime) > 150
      intTimer      = 350
    Case Else
      intTimer      = 200 + CLng(strBuildfileTime)
  End Select

  Set colOS         = objWMI.InstancesOf("Win32_OperatingSystem")
  For Each objOS In colOS
    strBootUpTime   = CStr(objOS.LastBootUpTime)
  Next
  strBootUpTime     = Left(strBootUpTime, Instr(strBootUpTime, ".") -1)
  strBootUpTime     = Mid(strBootUpTime, 1, 4) & "/" & Mid(strBootUpTime, 5, 2) & "/" & Mid(strBootUpTime, 7, 2) & " " & Mid(strBootUpTime, 9, 2) & ":" & Mid(strBootUpTime, 11, 2) & ":" & Mid(strBootUpTime, 13, 2)
  intDelay          = DateDiff("s", CDate(strBootUpTime), Now())

  strPath           = "SYSTEM\CurrentControlSet\Services\ClusSvc"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisplayName",strClusService
  objWMIReg.GetDWordValue strHKLM,strPath,"Start",strClusStatus
  Select Case True
    Case IsNull(strClusService)
      ' Nothing
    Case strOSVersion >= "6.0"
      ' Nothing
    Case strClusStatus = 2
      ' Nothing
    Case Else
      intDelay      = 0
      objWMIReg.GetStringValue strHKLM,strPath,"ObjectName",strAccount
      strCmd        = "NET LOCALGROUP """ & Mid(strLocalAdmin, Instr(strLocalAdmin, "\") + 1) & """ """ & strDomain & Mid(strAccount, Instr(strAccount, "\")) & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
      strCmd        = "NET START ClusSvc"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End Select

  Call DebugLog("Boot Time: " & strBootUpTime & ", Required delay: " & Cstr(intTimer) & " sec")
  If CLng(intDelay) < CLng(intTimer) Then
    intDelay        = CLng(intTimer - intDelay)
    Call DebugLog(" Waiting " & CStr(intDelay) & " seconds before proceeding") ' Wait at least 2m 30s following reboot to allow services to start
    Wscript.Sleep CStr(intDelay * 1000)
  End If

End Sub


Sub GetDNSDetails()
  Call SetProcessId("0AE","Get DNS Details for " & strUserDNSDomain)

  Call GetDNSServer()

  strServerIP       = GetAddress(strServer, "", "")

End Sub


Sub GetDNSServer()
  Call SetProcessId("0AEA", "Get DNS Server Name")
' Code based on sample published by Christian Dunn http://www.chrisdunn.name/jm/software/scriptsandcode/227-retrieve-dns-records
  On Error Resume Next
  Dim colAdapters
  Dim objAdapter
  Dim strDNSSuffix

  strUserDNSServer  = ""
  Set colAdapters   = objWMI.ExecQuery ("SELECT DNSServerSearchOrder,DNSDomainSuffixSearchOrder FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE",,0)
  For Each objAdapter In colAdapters
    strDNSSuffix    = ""
    Select Case True
      Case Not IsArray(objAdapter.DNSDomainSuffixSearchOrder)
        ' Nothing
      Case objAdapter.DNSDomainSuffixSearchOrder(0) = ""
        ' Nothing
      Case Else
        strDNSSuffix = UCase(objAdapter.DNSDomainSuffixSearchOrder(0))
    End Select
    Select Case True
      Case IsNull(objAdapter.DNSServerSearchOrder)
        ' Nothing
      Case IsNull(objAdapter.DNSServerSearchOrder(0))
        ' Nothing
      Case objAdapter.DNSServerSearchOrder(0) = ""
        ' Nothing
      Case objAdapter.DNSServerSearchOrder(0) = "0.0.0.0"
        ' Nothing
      Case (strDNSSuffix <> "") And (strDNSSuffix <> strUserDNSDomain)
        ' Nothing
      Case Else
        For intIdx = 0 To UBound(objAdapter.DNSServerSearchOrder, 1)
          strUserDNSServer = objAdapter.DNSServerSearchOrder(intIdx)
          strDebugMsg1    = "DNS Server: " & strUserDNSServer
          Set objWMIDNS   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strUserDNSServer & "\root\MicrosoftDNS")
          If IsObject(objWMIDNS) Then
            Exit For
          End If
        Next
        Exit For
    End Select
  Next

  On Error GoTo 0
  If Not IsObject(objWMIDNS) Then
    strUserDNSServer = ""
  End If
  Call SetBuildfileValue("UserDNSServer", strUserDNSServer)

End Sub


Sub OpenCluster()
  Call SetProcessId("0AF", "Setup Connection to Cluster")

  Select Case true
    Case GetBuildfileValue("SetupClusterCmdStatus") = strStatusComplete
      ' Nothing
    Case strOSVersion < "6.2" 
      ' Nothing
    Case Else
      Call SetupClusterCmd()
  End Select

  Call ConnectCluster()
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterName    = GetBuildfileValue("ClusterName")
  Select Case True
    Case strClusterName = ""
      Call SetBuildMessage(strMsgError, "FineBuild cancelled, unable to connect to cluster for " & strAction)
    Case strClusWinSuffix <> ""
      strClusterBase = Left(strClusterName, Len(strClusterName) - Len(strClusWinSuffix))
    Case Else
      strClusterBase = strClusterName
  End Select

  Call CheckClusterNodes()
  Call SetClusterData()

End Sub


Sub SetupClusterCmd()
  Call SetProcessId("0AFA", "Install legacy Cluster Command interface")

  strCmd            = strCmdPS & " -Command Install-WindowsFeature -Name RSAT-Clustering-AutomationServer"
  Call Util_RunExec(strCmd, "", "", 0)
  strCmd            = strCmdPS & " -Command Install-WindowsFeature -Name RSAT-Clustering-CmdInterface"
  Call Util_RunExec(strCmd, "", "", 0)
  WScript.Sleep strWaitShort

  Call SetBuildfileValue("SetupClusterCmdStatus", strStatusComplete)

End Sub


Sub CheckClusterNodes()
  Call SetProcessId("0AFC", "Check Cluster Node details")
  Dim colClusNodes, colResources
  Dim objClusNode, objResource

  Call DebugLog("Ensure Cluster Network Name is online")
  Set colResources = GetClusterResources()
  For Each objResource In colResources
    Select Case True
      Case Left(objResource.Name, 18) <> "Cluster IP Address" 
        ' Nothing
      Case Else
        Call SetResourceOn(objResource.Name, "")
    End Select
  Next

  Call DebugLog("Ensure Cluster Node is online")
  Set colClusNodes  = GetClusterNodes()
  For Each objClusNode In colClusNodes
    Select Case True
      Case Ucase(objClusNode.Name) = UCase(strServer)
        strClusterNode = objClusNode.NodeId
        If objClusNode.State = 1 Then
          strCmd    = "CLUSTER """ & strClusterName & """ NODE """ & strClusterNode & """ /START"
          Call Util_RunExec(strCmd, "", strResponseYes, -1)
        End If
        If objClusNode.State = 2 Then
          strCmd    = "CLUSTER """ & strClusterName & """ NODE """ & strClusterNode & """ /RESUME"
          Call Util_RunExec(strCmd, "", strResponseYes, -1)
        End If
        If objClusNode.State <> 0 Then
          Call DebugLog("Cluster Node " & objClusNode.Name & " state: " & CStr(objClusNode.State))
          Call SetBuildMessage(strMsgInfo,  "Cluster Node " & objClusNode.Name & " is in state: " & CStr(objClusNode.State))
        End If
    End Select
  Next

  strClusGroups     = GetAccountAttr(strClusterName, strUserDNSDomain, "memberOf")

End Sub


Sub SetClusterData()
  Call SetProcessId("0AFD", "Set Cluster data for SQL Instance")

  strClusSuffix     = GetClusSuffix()

  strClusterNameAS  = GetClusterName("ClusterNameAS",  "",       strClusterBase & strClusASSuffix & strClusSuffix)
  strClusterNameIS  = strClusterBase & strClusISSuffix
  strClusterNamePE  = strClusterBase & strClusPESuffix
  strClusterNamePM  = strClusterBase & strClusPMSuffix
  strClusterNameRS  = GetClusterName("ClusterNameRS",  "",       strClusterBase & strClusRSSuffix)
  strClusterNameSQL = GetClusterName("ClusterNameSQL", "",       strClusterBase & strClusDBSuffix & strClusSuffix)
  strClusterNameDTC = GetClusterName("ClusterNameDTC", "",       strClusterBase & strClusDTCSuffix & strClusSuffix)

  strClusterGroupAO = GetClusterName("ClusterNameAO",  "AGName", strClusterBase & strClusAOSuffix & strClusSuffix)
  strClusterGroupFS = GetClusterName("ClusterNameFS",  "",       strClusterBase & strClusFSSuffix)
  strClusterGroupRS = strClusterNameRS
  strClusterGroupSQL   = strClusterNameSQL & " (" & strInstance & ")"
  strClusterNetworkSQL = "SQL Network Name (" & strInstance & ")"

  If strSQLVersion >= "SQL2016" Then
    Call SetClusIPDetails("AA",  "YES")
  End If
  If strSQLVersion >= "SQL2016" Then
    Call SetClusIPDetails("AO",  "YES")
  End If
  Call SetClusIPDetails("AS",  "YES")
  Call SetClusIPDetails("DB",  strSetupSQLDBCluster)
  Call SetClusIPDetails("DTC", "YES")
  Call SetClusIPDetails("IS",  "YES")
  Call SetClusIPDetails("RS",  "YES")

End Sub


Function GetClusterName(strName, strAltName, strDefault)
  Call DebugLog("GetClusterName: " & strName)

  GetClusterName    = GetParam(Null,                 strName,          strAltName,                    "")

  If GetClusterName = "" Then
    GetClusterName  = GetBuildfileValue(strName)
  End If

  If GetClusterName = "" Then
    GetClusterName  = strDefault
    Call ParamListAdd("ListNonDefault", strName)
  End If

End Function


Function GetClusSuffix()
  Call DebugLog("GetClusSuffix:")
  Dim colClusGroups
  Dim objClusGroup
  Dim strGroupName

  strClusSuffix     = ""
  Set colClusGroups = GetClusterGroups()
  For Each objClusGroup In colClusGroups
    strGroupName    = objClusGroup.Name
    Select Case True
      Case Left(strGroupName, Len(strClusterBase & strClusDBSuffix)) = strClusterBase & strClusDBSuffix
        strClusSuffix = FindClusSuffix(strGroupName, strClusDBSuffix)
      Case Left(strGroupName, Len(strClusterBase & strClusDTCSuffix)) = strClusterBase & strClusDTCSuffix
        strClusSuffix = FindClusSuffix(strGroupName, strClusDTCSuffix)
    End Select
  Next

  For intIdx = 1 To Len(strAlphabet)
    Select Case True
      Case arrClusInstances(intIdx) > ""
        ' Nothing
      Case strClusSuffix <> ""
        ' Nothing
      Case Else
        strClusSuffix = Mid(strAlphabet, intIdx, 1)
    End Select
  Next

  GetClusSuffix     = strClusSuffix

End Function


Function FindClusSuffix(strGroupName, strGroupSuffix)
  Call DebugLog("FindClusSuffix:")
  Dim intIdx
  Dim strTempSuffix, strTempInstance

  strTempSuffix     = Mid(strGroupName, Len(strClusterBase & strGroupSuffix) + 1, 1)
  strTempInstance   = Mid(strGroupName, Len(strClusterBase & strGroupSuffix) + 4)
  intIdx            = Instr(strAlphabet, strTempSuffix)
  arrClusInstances(intIdx) = strTempInstance
  If strTempInstance = strInstance & ")" Then
    FindClusSuffix  = strTempSuffix
  End If

End Function


Sub SetClusIPDetails(strClusType, strClusSetup)
  Call DebugLog("SetClusIPDetails: " & strClusType)
  Dim strClusIPExtra,strClusIPSuffix

  strClusIPExtra    = GetParam(Null,                  "Clus" & strClusType & "IPExtra",      "", "")
  If strClusIPExtra <> "" Then
    Call SetBuildfileValue("Clus" & strClusType & "IPExtra", strClusIPExtra)
  End If

  strClusIPSuffix   = GetParam(Null,                  "Clus" & strClusType & "IPSuffix",     "", "")
  Select Case True
    Case strClusIPSuffix <> ""
      Call SetBuildfileValue("Clus" & strClusType & "IPSuffix", strClusIPSuffix)
    Case strClusSetup <> "YES"
      ' Nothing
    Case strUserDNSServer = ""
      Call SetBuildMessage(strMsgErrorConfig, "Unable to connect to DNS Server.  /Clus" & strClusType & "IPSuffix: parameter must be supplied")
  End Select

End Sub


Sub SetSetupFlags()
  Call SetProcessId("0AG","Set Setup flags for install")

  If strClusterHost <> "YES" Then
    Call SetParam("SetupSQLASCluster",       strSetupSQLASCluster,     "N/A", "", strListCluster)
    Call SetParam("SetupSQLDBCluster",       strSetupSQLDBCluster,     "N/A", "", strListCluster)
    Call SetParam("SetupSSISCluster",        strSetupSSISCluster,      "N/A", "", strListCluster)
    Call SetParam("SetupSQLRSCluster",       strSetupSQLRSCluster,     "N/A", "", strListCluster)
    If strSQLVersion < "SQL2017" Then
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListCluster)
    End If
  End If

  Call SetupInstDB()

  Call GetDagData()

  Call SetupInstAS()

  Call SetupInstIS()

  Call SetupInstRS()

  Call SetupInstDTC()

  Call SetupInstMR()

  strActionSQLTools = GetItemAction(strType, strAction, "SQLTools", "NO")

  strDirServInst    = Replace(strServinst, "\", "$")

End Sub


Sub GetDagData()
  Call DebugLog("GetDagData:")
  Dim objSQL, objSQLData
  Dim strAGDagPrimary

  strActionDAG      = ""
  strAGDagPrimary   = GetStatefileValue(strAGDagName)
  Select Case True
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case strAGDagName = ""
      ' Nothing
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strEditionEnt <> "YES"
      ' Nothing
    Case strClusterHost <> "YES"
      ' Nothing
    Case strAGDagPrimary = ""
      strActionDAG  = "INSTALL"
    Case strAGDagPrimary = strClusterNameSQL
      strActionDAG  = "INSTALL"
    Case Else
      strActionDAG  = "ADDNODE"
  End Select

  Select Case True
    Case strActionDAG = ""
      ' Nothing
    Case strAGDagPrimary = ""
      ' Nothing
    Case strAGDagPrimary = strClusterNameSQL
      Call MoveToNode(strAGDagPrimary, GetPrimaryNode(strClusterGroupSQL))
      Call SetResourceOn(strAGDagPrimary, "GROUP")
      Set objSQL      = CreateObject("ADODB.Connection")
      Set objSQLData  = CreateObject("ADODB.Recordset")
      objSQL.Provider = "SQLOLEDB"
      objSQL.ConnectionString = "Server=" & strAGDagPrimary & ";Database=master;Trusted_Connection=Yes;"
      strDebugMsg1  = objSQL.ConnectionString
      objSQL.Open 
      strCmd        = "SELECT COUNT(*) AS AGDagNodes "
      strCmd        = strCmd & "FROM master.sys.availability_replicas AS r "
      strCmd        = strCmd & "JOIN master.sys.availability_groups AS a "
      strCmd        = strCmd & "  ON r.group_id = a.group_id AND a.name = '" & strAGDagName & "' "
      strCmd        = strCmd & "WHERE r.replica_server_name <> '' "
      strDebugMsg2  = strCmd
      Set objSQLData  = objSQL.Execute(strCmd)
      Do Until objSQLData.EOF
        strAGDagNodes = objSQLData.Fields("AGDagNodes")
        objSQLData.MoveNext
      Loop
  End Select

End Sub


Sub SetupInstAS()
  Call DebugLog("SetupInstAS:")
  Dim strCapMemory

  Select Case True
    Case strSetupSQLASCluster <> ""
      ' Nothing
    Case (strClusterAction <> "") And (strMainInstance = "YES")
      strSetupSQLASCluster = "YES"
    Case Else
      strSetupSQLASCluster = "NO"
  End Select

  Select Case True
    Case strSetupSQLASCluster = "YES"
      strSetupSQLAS = "YES"
    Case strSetupSQLAS <> ""
      ' Nothing
    Case strMainInstance = "YES"
      strSetupSQLAS = "YES"
    Case Else   
      Call SetParam("SetupSQLAS",            strSetupSQLAS,            "NO",  "SSAS not installed by default with secondary SQL Instances on server", "")
  End Select

  If strSetupSQLAS <> "YES" Then
    Call SetParam("SetupSQLASCluster",       strSetupSQLASCluster,     "N/A", "", strListSQLAS)
  End If

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      strInstASSQL  = strInstance
      strInstASCon  = strServer
      strInstAS	    = "MSSQLServerOLAPService"
      strInstNodeAS = "MSAS.MSSQLSERVER"
    Case Else
      strInstASSQL  = strInstance
      strInstASCon  = strServer & "\" & strInstASSQL
      strInstAS	    = "MSOLAP$" & strInstance
      strInstNodeAS = "MSAS."  & strInstance
  End Select

  Select Case True
    Case strSetupSQLASCluster <> "YES"
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      strInstASSQL      = strClusterNameAS
      strInstASCon      = strClusterNameAS
      strInstAS	        = "MSOLAP$" & strClusterNameAS
      strInstNodeAS     = "MSAS."  & strClusterNameAS
      strClusterGroupAS = strClusterNameAS & " (" & strInstASSQL & ")"
      strResSuffixAS    = " (" & strClusterNameAS & ")"
    Case Else
      strClusterGroupAS = strClusterNameAS & " (" & strInstASSQL & ")"
      strInstASCon      = strClusterNameAS & "\" & strInstASSQL
      strResSuffixAS    = " (" & strInstance & ")"
  End Select

  strActionSQLAS    = GetItemAction(strType, strAction, "SQLAS", strSetupSQLASCluster)
  strClusterNetworkAS  = "SQL Network Name (" & strInstASSQL & ")"

  strCapMemory      = 6000
  Select Case True
    Case Not IsNumeric("0" & strSetTotalMemLimit)
      Call SetBuildMessage(strMsgErrorConfig, "/SetTotalMemoryLimit: value must be Numeric")
    Case strSetTotalMemLimit <> ""
      ' Nothing
    Case strServerMB <= (strCapMemory / 80) * 100
      strSetTotalMemLimit = "80"
    Case Else
      strSetTotalMemLimit = Int(strCapMemory * 1000 * 1000)
      Call ParamListAdd("ListNonDefault", "SetTotalMemoryLimit")
  End Select

End Sub


Sub SetupInstDB()
  Call DebugLog("SetupInstDB:")

  Select Case True
    Case strSetupSQLDBCluster <> ""
      ' Nothing
    Case strSetupSQLDB = "NO"
      strSetupSQLDBCluster = "NO"
    Case strClusterAction <> ""
      strSetupSQLDBCluster = "YES"
    Case Else
      strSetupSQLDBCluster = "NO"
  End Select

  Select Case True
    Case strSetupSQLDBCluster = "YES"
      strSetupSQLDB  = "YES"
    Case strSetupSQLDB <> ""
      ' Nothing
    Case Else
      strSetupSQLDB  = "YES"
  End Select

  If strSetupSQLDB <> "YES" Then
    Call SetParam("SetupSQLDBCluster",       strSetupSQLDBCluster,     "N/A", "", strListSQLDB)
    Call SetParam("SetupAlwaysOn",           strSetupAlwaysOn,         "N/A", "", strListSQLDB)
  End If

  Select Case True
    Case strType = "CLIENT"
      strInstNode   = "CLIENT"
      strInstLog    = ""
    Case strInstance = "MSSQLSERVER"
      strInstAgent  = "SQLSERVERAGENT"
      strInstAnal   = "MSSQLLaunchpad"
      strInstAO     = strServer & "\" & "DEFAULT"
      strInstFT	    = "MSSQLFDLauncher"
      strInstPE     = "SQLPBENGINE"
      strInstPM     = "SQLPBDMS"
      strInstSQL    = "MSSQLSERVER"
      strInstStream = "Default"
      strServInst   = strServer
      strServName   = "SQL Server (" & strInstance & ")"
      strInstNode   = "MSSQL.MSSQLSERVER"
      strInstNodeIS = "MSIS.MSSQLSERVER"
      strInstLog    = ""
      strInstTel    = "SQLTELEMETRY"
    Case Else
      strInstAgent  = "SQLAgent$" & strInstance
      strInstAnal   = "MSSQLLaunchpad$" & strInstance
      strInstAO     = strServer & "\" & strInstance
      strInstPE     = "SQLPBENGINE"
      strInstPM     = "SQLPBDMS"
      strInstSQL    = "MSSQL$" & strInstance
      strInstStream = strInstance
      strServInst   = strServer & "\" & strInstance
      strServName   = "SQL Server (" & strInstance & ")"
      strInstNode   = "MSSQL." & strInstance
      strInstNodeIS = "MSIS.MSSQLSERVER"
      strInstLog    = strInstance & " "
      strInstFT	    = "MSSQLFDLauncher"
      If strSQLVersion > "SQL2005" Then
        strInstFT   = "MSSQLFDLauncher$" & strInstance
      End If
      strInstTel    = "SQLTELEMETRY"
  End Select

  Select Case True
    Case strClusterHost <> "YES"
      strServerAO    = strServer
    Case strInstance = "MSSQLSERVER"
      strInstAO      = strClusterNameSQL & "\" & "DEFAULT"
      strResSuffixDB = ""
      strServInst    = strClusterNameSQL
      strServerAO    = strClusterNameSQL
    Case Else
      strInstAO      = strClusterNameSQL & "\" & strInstance
      strResSuffixDB = " (" & strInstance & ")"
      strServInst    = strClusterNameSQL & "\" & strInstance
      strServerAO    = strClusterNameSQL
  End Select

  strActionSQLDB    = GetItemAction(strType, strAction, "SQLDB", strSetupSQLDBCluster)

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      strGroupAO    = ""
    Case strClusterHost = "YES"
      strGroupAO    = strClusterGroupAO
    Case Else
      strGroupAO    = strAGName
  End Select

End Sub


Sub SetupInstDTC()
  Call DebugLog("SetupInstDTC:")

  Select Case True
    Case strClusterHost <> "YES"
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "N/A", "", strListCluster)
    Case strSetupDTCCluster <> ""
      ' Nothing
    Case strClusterHost = "YES"
      strSetupDTCCluster = "YES"
    Case Else
      strSetupDTCCluster = "NO"
  End Select

  strActionDTC      = GetItemAction(strType, strAction, "DTC", strSetupDTCCluster)

  If strClusterDTCFound = "" Then
    strDTCClusterRes = ""
  End If

  Select Case True
    Case strOSVersion < "6.0"
      Call SetParam("DTCMultiInstance",      strDTCMultiInstance,      "N/A", "", strListOSVersion)
    Case strDTCMultiInstance <> "YES"
      ' Nothing
    Case Else
      strLabDTC      = strLabDTC & strClusSuffix
  End Select

End Sub


Sub SetupInstMR()
  Call DebugLog("SetupInstMR:")

  Select Case True
    Case strSQLSharedMR = "YES"
      ' Nothing
    Case strInstMR <> ""
      ' Nothing
    Case strClusterHost = "YES"
      strInstMR     = strClusterBase & strClusMRSuffix
    Case Else
      strInstMR     = "MSRService"
  End Select

End Sub


Sub SetupInstIS()
  Call DebugLog("SetupInstIS:")

  Select Case True
    Case strSetupSQLIS <> ""
      ' Nothing
    Case strMainInstance = "YES"
      strSetupSQLIS = "YES"
    Case Else
      Call SetParam("SetupSQLIS",            strSetupSQLIS,            "NO",  "SSIS not required for secondary SQL Instances on server", "")
  End Select


  Select Case True
    Case strSetupSQLASCluster = "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      ' Nothing
    Case Else
      Call SetParam("SetupSSISCluster",      strSetupSSISCluster,      "N/A", "", strListCluster)
  End Select

  Select Case True
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupSSISCluster",      strSetupSSISCluster,      "N/A", "", strListCluster)
    Case strSetupAPCluster = "YES"
      Call SetParam("SetupSSISCluster",      strSetupSSISCluster,      "YES", "SSIS Cluster mandatory for /SetupAPCluster: " & strSetupAPCluster, "")
    Case strSetupSSISCluster = ""
      strSetupSSISCluster = "NO"
  End Select

  strActionSQLIS    = GetItemAction(strType, strAction, "SQLIS", strSetupSQLDBCluster)

  Select Case True
    Case strSetupISMasterCluster = "YES"
      strDNSNameIM  = strClusterBase & strClusIMSuffix
      strDNSIPIM    = ""
    Case Else
      strDNSNameIM  = strServer
      strDNSIPIM    = strServerIP
  End Select

End Sub


Function GetItemAction(strType, strAction, strItem, strItemCluster)
  Call DebugLog("GetItemAction:" & strItem)

  Select Case True
    Case strItemCluster = "YES"
      GetItemAction = GetItemActionCluster(strItem)
    Case strType = "UPGRADE"
      GetItemAction = "UPGRADE"
    Case Else
      GetItemAction = "INSTALL"
 End Select
 Call DebugLog(" Action: " & GetItemAction)

End Function


Function GetItemActionCluster(strItem)
  Call DebugLog("GetItemActionCluster:" & strItem)
  Dim colResources
  Dim objResource
  Dim strCurrentAction, strItemAction

  strCurrentAction  = GetBuildfileValue("Action" & strItem)
  strItemAction     = strActionClusInst
  Set colResources  = GetClusterResources()
  For Each objResource In colResources
    Select Case True
      Case (strItem = "DTC") And (objResource.TypeName = "Distributed Transaction Coordinator")
        Select Case True
          Case strCurrentAction <> ""
            strItemAction      = strCurrentAction
            strClusterDTCFound = "Y"
          Case Else
            strItemAction      = "ADDNODE"
            strClusterDTCFound = "Y"
        End Select
        Call SetResourceOnline(objResource, strItemAction)
      Case (strItem = "SQLDB") And (objResource.Name = "SQL Server")
        Select Case True
          Case objResource.PrivateProperties("InstanceName") <> strInstance
            Exit For
          Case strCurrentAction <> ""
            strItemAction      = strCurrentAction
            strClusterSQLFound = "Y"
          Case Else
            strItemAction      = "ADDNODE"
            strClusterSQLFound = "Y"
        End Select
        Call SetResourceOnline(objResource, strItemAction)
      Case (strItem = "SQLAS") And (objResource.Name = "Analysis Services" & strResSuffixAS)
        Select Case True
          Case objResource.PrivateProperties("ServiceName") <> strInstAS
            Exit For
          Case strCurrentAction <> ""
            strItemAction = strCurrentAction
          Case Else
            strItemAction = "ADDNODE"
        End Select
        Call SetResourceOnline(objResource, strItemAction)
      Case (strItem = "SQLRS") And (objResource.Name = "Reporting Services")
        Select Case True
          Case objResource.PrivateProperties("ServiceName") <> strInstRS
            Exit For
          Case strCurrentAction <> ""
            strItemAction = strCurrentAction
          Case Else
            strItemAction = "ADDNODE"
        End Select
        Call SetResourceOnline(objResource, strItemAction)
      Case (strItem = "SQLIS") And (objResource.CommonProperties("Description") = "SQL Server Integration Services")
        Select Case True
          Case strCurrentAction <> ""
            strItemAction = strCurrentAction
          Case Else
            strItemAction = "ADDNODE"
        End Select
      Case (strItem = "AO") And (objResource.TypeName = "SQL Server Availability Group") And (objResource.Name = strClusterGroupAO)
        Select Case True
          Case strCurrentAction <> ""
            strItemAction = strCurrentAction
          Case Else
            strItemAction = "ADDNODE"
        End Select
    End Select
  Next
 
  strClusterAction     = strItemAction
  GetItemActionCluster = strItemAction

End Function


Sub SetResourceOnline(objResource, strItemAction)
  Call DebugLog("SetResourceOnline: " & objResource.Name)
  Dim objGroup
  Dim strGroup, strResource, strState

  Set objGroup      = objResource.Group
  strGroup          = objGroup.Name
  strResource       = objResource.Name
  strState          = objResource.State

  Select Case True
    Case IsNull(objGroup)
      Call SetBuildMessage(strMsgErrorConfig, "Owner Group for Resource """ & objResource.Name & """ can not be found")
    Case strState = 0 ' Resource Inherited
      ' Nothing
    Case strState = 2 ' Resource Operational
      ' Nothing
    Case strItemAction = "ADDNODE"
      Call MoveToNode(strGroup, GetPrimaryNode(strGroup))
      Call SetResourceOn(strResource, "")
    Case Else
      Call MoveToNode(strGroup, "")
      Call SetResourceOn(strResource, "")
  End Select

End Sub


Sub EarlyValidate()
  Call SetProcessId("0AH","Perform high-priority validation")
  Dim objExec

  Select Case True
    Case strServParm = ""
      ' Nothing
    Case strServParm = UCase(strServer)
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Requested server " & strServParm & " does not match actual server")
  End Select

  Select Case True
    Case strUserAdmin = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "SQL FineBuild must be run using Administrator priviliges")
  End Select

  Select Case True
    Case Instr(strTypeList, strType) > 0 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Unknown /Type:" & strType)
  End Select

  If strClusterHost <> "" Then
    Call CheckClusterResources()
  End If

  Call CheckSQLMemory()

End Sub


Sub CheckClusterResources()
  Call SetProcessId("0AHA", "Check state of cluster resources")

  strClusStorage    = GetStorageGroup(GetParam(colStrings, "ClusStorage", "", "Available Storage"))

  If strProcessId < "2C" Then
    Call MoveToNode(strClusStorage, "")
  End If

  Call SetResourceAllOn()

End Sub

Sub CheckSQLMemory()
  Call DebugLog("CheckSQLMemory:")
  Dim intMemWork, intOSMemory, intSQLMinMem, intSQLReqMem, intSQLMaxMem, intThreadMemory
' intSQLMinMem: Minimum possible memory needed to install components
' intSQLReqMem: Recommended minimum memory for components
' intSQLMaxMem: SQL Maximum memory limit

  intOSMemory       = 0
  intSQLMemory      = 0
  intSQLReqMem      = 0

  intMemWork        = 0
  Select Case True ' OS overhead
    Case strServerMB <= 4096 ' 4GB
      intMemWork    = Max(256 + (strServerMB / 4), 512)
    Case strServerMB < 16384 ' 16GB
      intMemWork    = Max(512 + (strServerMB / 4), 512)
    Case Else
      intMemWork    = 512 + (Min(strServerMB, 16384) / 4) + (Max(strServerMB - 16384, 0) / 8)
      intMemWork    = intMemWork + (1024 * (Int(intProcNum / 4) + 1))
  End Select
  intOSMemory       = intOSMemory + intMemWork

  Select Case True
    Case strOSVersion < "6"
      intSQLMinMem  = intSQLMinMem + 256
    Case Else
      intSQLMinMem  = intSQLMinMem + 512
  End Select

  intMemWork        = strSetTotalMemLimit ' Add AS Overhead (given in Bytes or %)
  Select Case True 
    Case strSetupSQLAS <> "YES"
      intMemWork    = 0
    Case intMemWork < "100"
      intSQLMinMem  = intSQLMinMem + 256
      intMemWork    = strServerMB * 0.1  ' Based on 'Best Guess' of 10% Server mem
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
    Case CLng(strServerMB * 0.1) >= CLng((intMemWork / 1024) / 1024) ' Best Guess mem > AS mem limit
      intMemWork    = (intMemWork / 1024) / 1024
      intSQLMinMem  = intSQLMinMem + 256
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
    Case Else
      intSQLMinMem  = intSQLMinMem + 256
      intMemWork    = strServerMB * 0.1
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
  End Select
  intOSMemory       = intOSMemory + intMemWork

  Select Case True ' Add Analytics Overhead
    Case strSetupAnalytics <> "YES"
      intMemWork    = 0
    Case Else
      intSQLMinMem  = intSQLMinMem + 64
      intMemWork    = strServerMB * 0.05 ' Based on 'Best Guess' of 5% Server mem
      intSQLReqMem  = intSQLReqMem + 64  ' No recommendation available
  End Select
  intOSMemory       = intOSMemory + intMemWork

  intMemWork        = 512 + (Max(intProcNum - 4, 0) * 16) ' Thread Count
  Select Case True
    Case strSetupSQLDB <> "YES"
      intMemWork    = 0
    Case strServerMB <= 4096
      intMemWork    = intMemWork
    Case Else
      intMemWork    = intMemWork * 2 
  End Select
  intOSMemory       = intOSMemory + intMemWork

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strEdition = "EXPRESS"
      intSQLMinMem  = intSQLMinMem + 256
      intSQLReqMem  = intSQLReqMem + 1024
    Case Else
      intSQLMinMem  = intSQLMinMem + 512
      intSQLReqMem  = intSQLReqMem + 4096 ' https://docs.microsoft.com/en-us/sql/sql-server/install/hardware-and-software-requirements-for-installing-sql-server-ver15
  End Select

  Select Case True
    Case strSetupDQ <> "YES"
      ' Nothing
    Case Else
      intSQLMinMem  = intSQLMinMem + 512
      intSQLReqMem  = intSQLReqMem + 4096 ' https://docs.microsoft.com/en-us/sql/data-quality-services/install-windows/install-data-quality-services
  End Select

  Select Case True
    Case strSetupPolyBase <> "YES"
      ' Nothing
    Case Else
      intSQLMinMem  = intSQLMinMem + 512
      intSQLReqMem  = intSQLReqMem + 16384 ' https://docs.microsoft.com/en-us/sql/relational-databases/polybase/polybase-installation
  End Select

  intMemWork        = strSetWorkingSetMaximum ' Add RS Overhead (given in KB or %)
  Select Case True 
    Case strSetupSQLRS <> "YES"
      intMemWork    = 0
    Case intMemWork < "100"
      intSQLMinMem  = intSQLMinMem + 256
      intMemWork    = strServerMB * 0.1  ' Based on 'Best Guess' of 10% Server mem
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
    Case CLng(strServerMB * 0.1) >= CLng(intMemWork / 1024) ' Best Guess mem > RS mem limit
      intSQLMinMem  = intSQLMinMem + 256
      intMemWork    = intMemWork / 1024
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
    Case Else
      intSQLMinMem  = intSQLMinMem + 256
      intMemWork    = strServerMB * 0.1
      intSQLReqMem  = intSQLReqMem + 256 ' No recommendation available
  End Select
  intOSMemory       = intOSMemory + intMemWork

  intSQLMaxMem      = Max(strServerMB - intOSMemory, 768) ' Ensure SQL gets something on low-memory systems
  Select Case True ' Limit memory for specific Editions
    Case (intSQLMaxMem> 1024) And (strEdition = "EXPRESS")
      intSQLMaxMem  = 1024
    Case (intSQLMaxMem> 65536) And (InStr(" STANDARD WEB WORKGROUP ", strEdition) > 0)
      intSQLMaxMem  = 65536
  End Select

  Select Case True
    Case strSQLMaxMemory = ""
      strSQLMaxMemory  = Int(intSQLMaxMem)
      Call ParamListAdd("ListNonDefault", "SQLMaxMemory")
    Case Else
      ' Nothing
  End Select

  If intSQLMinMem > (strServerMB * 1.01) Then
    Call SetBuildMessage(strMsgErrorConfig, "Required minimum memory " & CStr(intSQLMinMem) & " MB exceeds server memory " & strServerMB & " MB")
  End If
  If intSQLReqMem > (strServerMB * 1.01) Then
    Call SetBuildMessage(strMsgInfo, "Recommended memory " & CStr(intSQLReqMem) & " MB exceeds server memory " & strServerMB & " MB")
  End If

End Sub


Sub CheckRebootStatus()
  Call SetProcessId("0AI", "Check if Reboot is required")
  Dim arrOperations
  Dim objKey
  Dim strPathReg

  strRebootStatus   = GetBuildfileValue("RebootStatus")
  Select Case True
    Case strRebootStatus = ""
      strRebootStatus = "N/A"
    Case strRebootStatus = "Done"
      strRebootStatus = "N/A"
      If strAdminPassword <> "" Then
        strPath     = "HKLM\" & strOSRegPath & "Winlogon\AutoAdminLogon"
        Call Util_RegWrite(strPath, "0", "REG_SZ")
        strPath     = strOSRegPath & "Winlogon"
        objWMIReg.DeleteValue strHKLM, strPath, "DefaultPassword"
      End If
  End Select

  Select Case True
    Case (strType = "FIX") And (strProcessId > "3")
      ' Nothing
    Case strProcessId > "1"
      ' Nothing
    Case CheckReboot() = "Pending" 
      Call DebugLog("Reboot is pending")
  End Select

End Sub


Sub GetBuildfileData()

  Call SetProcessId("0B", "Get data needed for Buildfile")

  Call GetLanguageData()

  Call GetSQLPath()

  Call GetEditionData()

  If strClusterHost = "YES" Then
    Call GetClusterData()
  End If

  Call GetServerData()

  Call GetAccountData()

  Call GetGroupData()

  Call GetVolumeData()

  Call GetPIDData()

  Call GetMenuData()

  Call GetPathData()

  Call GetFileData()

  Call GetSetupData()

  Call GetMiscData()

End Sub


Sub GetLanguageData()
  Call SetProcessId("0BA", "Get data for OS and SQL Languages")

  strOSLanguage     = strLanguage

  Select Case True
    Case colArgs.Exists("ENU")
      strEnu         = "YES"
      strSQLLanguage = "ENU"
    Case Else
      strEnu         = "NO"
      strSQLLanguage = strLanguage
  End Select  

End Sub


Sub GetSQLPath()
  Call SetProcessId("0BB", "Get SQL Media paths")
  Dim strPathSQL

  strPathFB         = GetMediaPath(strPathFB)
  strPathFBScripts  = strPathFB & "Build Scripts\"
 
  strPathSQL        =  GetSQLMediaPath(strSQLVersion, strPathSQLMediaOrig)
  Select Case True
    Case strPathSQL = strPathSQLDefault
      strPathSQLMediaOrig = strPathSQL
    Case strPathSQLMediaOrig <> strPathSQL
      strPathSQLMediaOrig = ""
  End Select
  strPathSQLMedia   = GetMediaPath(strPathSQL)
  Select Case True
    Case strPathSQLMedia = ""
      Call SetBuildMessage(strMsgErrorConfig, "SQL Install Media cannot be found")
    Case Else
      Call DebugLog("PathSQLMedia: " & strPathSQLMedia)
  End Select

  strPathSQLSPOrig  = GetParam(colStrings,            "PathSQLSP",       "",                    "..\Service Packs")
  Select Case True
    Case GetMediaPath(strPathSQLSPOrig) <> ""
      ' Nothing
    Case GetMediaPath(Replace(strPathSQLSPOrig, " ", "")) <> ""
      strPathSQLSPOrig = Replace(strPathSQLSPOrig, " ", "")
  End Select

  strPathAutoConfigOrig = GetParam(colStrings,            "PathAutoConfig",  "",                     "AutoConfig")
  strPathAutoConfig     = strPathAutoConfigOrig
  Select Case True
    Case Mid(strPathAutoConfig, 1, 1) = "."
      strPathAutoConfig = GetMediaPath(strPathAutoConfig)
    Case Mid(strPathAutoConfig, 2, 1) = ":"
      strPathAutoConfig = GetMediaPath(strPathAutoConfig)
    Case GetMediaPath(FormatFolder(strPathFBStart) & strPathAutoConfig) <> ""
      strPathAutoConfig = GetMediaPath(FormatFolder(strPathFBStart) & strPathAutoConfig)
    Case Else
      strPathAutoConfig = GetMediaPath(strPathAddComp & strPathAutoConfig)
  End Select

  strActionAO       = GetItemActionAO()

End Sub


Function GetSQLMediaPath(strSQLVersion, strPathSQLMediaOrig)
  Call DebugLog("GetSQLMediaPath: " & strSQLVersion)
  Dim strPathSQL, strSuffix

  Select Case True
    Case strEdition = "BUSINESS INTELLIGENCE"
      strSuffix     = "_BI"
    Case strEdition = "DATA CENTER"
      strSuffix     = "_DC"
    Case strEdition = "DEVELOPER"
      strSuffix     = "_Dev"
    Case strEdition = "EXPRESS"
      strSuffix     = "_Exp"
    Case strEdition = "ENTERPRISE"
      strSuffix     = "_Ent"
    Case strEdition = "ENTERPRISE EVALUATION"
      strSuffix     = "_Eval"
    Case strEdition = "STANDARD"
      strSuffix     = "_Std"
    Case strEdition = "WEB"
      strSuffix     = "_Web"
    Case strEdition = "WORKGROUP"
      strSuffix     = "_Wkg"
    Case Else
      strSuffix     = ""
  End Select

  Select Case True
    Case objFSO.FolderExists(strPathSQLMediaOrig)
      strPathSQL    = strPathSQLMediaOrig
    Case Instr(strOSType, "CORE") > 0 
      strPathSQL    =  GetPathSQL(strSQLVersion, strSuffix & "_Core", strPathSQLMediaOrig)
  End Select
  If strPathSQL = "" Then
    strPathSQL      =  GetPathSQL(strSQLVersion, strSuffix, strPathSQLMediaOrig)
  End If

  GetSQLMediaPath   = strPathSQL

End Function


Function GetPathSQL(strSQLVersion, strSuffix, strPathSQLMediaOrig)
  Call DebugLog("GetSQLPath: " & strSuffix)
  Dim strPathSQL

  Select Case True
    Case (strPathSQLMediaOrig <> strPathSQLDefault) And (strPathSQLMediaOrig <> "")
      strPathSQL    = strPathSQLMediaOrig
    Case strSuffix = ""
      strPathSQL    = strPathDefault
    Case GetMediaPath("..\" & strSQLVersion & strSuffix & "_" & strFileArc & "_" & strLanguage) <> ""
      strPathSQL    = "..\" & strSQLVersion & strSuffix & "_" & strFileArc & "_" & strLanguage
    Case GetMediaPath("..\" & strSQLVersion & strSuffix & "_" & strFileArc) <> ""
      strPathSQL    = "..\" & strSQLVersion & strSuffix & "_" & strFileArc
    Case GetMediaPath("..\" & strSQLVersion & strSuffix & "_" & strLanguage) <> ""
      strPathSQL    = "..\" & strSQLVersion & strSuffix & "_" & strLanguage
    Case GetMediaPath("..\" & strSQLVersion & strSuffix) <> ""
      strPathSQL    = "..\" & strSQLVersion & strSuffix
    Case Instr(strSuffix, "Core") > 0
      strPathSQL    = ""
    Case Else
      strPathSQL    = strPathSQLDefault
  End Select

  GetPathSQL        = strPathSQL

End Function


Function GetItemActionAO()
  Call DebugLog("GetItemActionO:")
  Dim strAGState

  strAGState        = GetStatefileValue(strGroupAO)

  Select Case True
    Case strAGState = ""
      GetItemActionAO   = "INSTALL"
    Case strAGState = strServInst
      GetItemActionAO   = "INSTALL"
    Case Else
      GetItemActionAO   = "ADDNODE"
  End Select

End Function


Sub GetEditionData()
  Call SetProcessId("0BC", "Setup Install flags depending on Edition")

  Select Case True                   
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupSQLTools",      strSetupSQLTools,      "N/A", "", strListCore)
  End Select

  Select Case True
    Case strEdition = "EXPRESS"
      Call GetExpressEditionData
    Case strSQLVersion = "SQL2005"
      strSQLExe     = GetParam(colFiles,    "SQLFullExe",         "",                    "Servers\SETUP.EXE")
    Case Else
      strSQLExe     = GetParam(colFiles,    "SQLFullExe",         "",                    "SETUP.EXE")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012" 
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListSQLVersion)
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListSQLVersion)
      Call SetParam("SetupPowerBIDesktop",   strSetupPowerBIDesktop,   "N/A", "", strListSQLVersion)
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A", "", strListSQLVersion)
      Call SetParam("SetupRServer",          strSetupRServer,          "N/A", "", strListSQLVersion)
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A", "", strListSQLVersion)
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListSQLVersion)
    Case strSQLVersion < "SQL2017"
      Call SetParam("SetupISMaster",         strSetupISMaster,         "N/A", "", strListSQLVersion)
      Call SetParam("SetupISWorker",         strSetupISWorker,         "N/A", "", strListSQLVersion)
      Call SetParam("SetupISMasterCluster",  strSetupISMasterCluster,  "N/A", "", strListSQLVersion)
      Call SetParam("SetupPython",           strSetupPython,           "N/A", "", strListSQLVersion)
    Case strSQLVersion < "SQL2019"
      Call SetParam("SetMemoryOptimizedHybridBufferpool", strSetMemOptHybridBP,   "N/A", "", strListSQLVersion)
      Call SetParam("SetMemoryOptimizedTempdbMetadata",   strSetMemOptTempdb,     "N/A", "", strListSQLVersion)
  End Select

  Select Case True
    Case strtype = "CLIENT"
      Call SetParam("SetupSQLTools",         strSetupSQLTools,         "YES", "SQL Tools mandatory for CLIENT build", "")
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListType)
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A", "", strListType)
      Call SetParam("SetupAPCluster",        strSetupAPCluster,        "N/A", "", strListType)
      Call SetParam("SetupBPE",              strSetupBPE,              "N/A", "", strListType)
      Call SetParam("SetupDBMail",           strSetupDBMail,           "N/A", "", strListType)
      Call SetParam("SetupDBAManagement",    strSetupDBAManagement,    "N/A", "", strListType)
      Call SetParam("SetupDBOpts",           strSetupDBOpts,           "N/A", "", strListType)
      Call SetParam("SetupDCom",             strSetupDCom,             "N/A", "", strListType)
      Call SetParam("SetupDisableSA",        strSetupDisableSA,        "N/A", "", strListType)
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListType)
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "N/A", "", strListType)
      Call SetParam("SetupGenMaint",         strSetupGenMaint,         "N/A", "", strListType)
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListType)
      Call SetParam("SetupISMaster",         strSetupISMaster,         "N/A", "", strListType)
      Call SetParam("SetupISMasterCluster",  strSetupISMasterCluster,  "N/A", "", strListType)
      Call SetParam("SetupISWorker",         strSetupISWorker,         "N/A", "", strListType)
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A", "", strListType)
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListType)
      Call SetParam("SetupNetwork",          strSetupNetwork,          "N/A", "", strListType)
      Call SetParam("SetupNonSAAccounts",    strSetupNonSAAccounts,    "N/A", "", strListType)
      Call SetParam("SetupOldAccounts",      strSetupOldAccounts,      "N/A", "", strListType)
      Call SetParam("SetupParam",            strSetupParam,            "N/A", "", strListType)
      Call SetParam("SetupPBM",              strSetupPBM,              "N/A", "", strListType)
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A", "", strListType)
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListType)
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListType)
      Call SetParam("SetupRSAdmin",          strSetupRSAdmin,          "N/A", "", strListType)
      Call SetParam("SetupRSAlias",          strSetupRSAlias,          "N/A", "", strListType)
      Call SetParam("SetupRSExec",           strSetupRSExec,           "N/A", "", strListType)
      Call SetParam("SetupRSIndexes",        strSetupRSIndexes,        "N/A", "", strListType)
      Call SetParam("SetupRSKeepAlive",      strSetupRSKeepAlive,      "N/A", "", strListType)
      Call SetParam("SetupSAAccounts",       strSetupSAAccounts,       "N/A", "", strListType)
      Call SetParam("SetupServices",         strSetupServices,         "N/A", "", strListType)
      Call SetParam("SetupServiceRights",    strSetupServiceRights,    "N/A", "", strListType)
      Call SetParam("SetupKerberos",         strSetupKerberos,         "N/A", "", strListType)
      Call SetParam("SetupSQLAS",            strSetupSQLAS,            "N/A", "", strListType)
      Call SetParam("SetupSQLASCluster",     strSetupSQLASCluster,     "N/A", "", strListType)
      Call SetParam("SetupSQLDB",            strSetupSQLDB,            "N/A", "", strListType)
      Call SetParam("SetupSQLDBAG",          strSetupSQLDBAG,          "N/A", "", strListType)
      Call SetParam("SetupSQLDBCluster",     strSetupSQLDBCluster,     "N/A", "", strListType)
      Call SetParam("SetupSQLDBRepl",        strSetupSQLDBRepl,        "N/A", "", strListType)
      Call SetParam("SetupSQLDBFS",          strSetupSQLDBFS,          "N/A", "", strListType)
      Call SetParam("SetupSQLDBFT",          strSetupSQLDBFT,          "N/A", "", strListType)
      Call SetParam("SetupSQLIS",            strSetupSQLIS,            "N/A", "", strListType)
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "N/A", "", strListType)
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListType)
      Call SetParam("SetupSQLMail",          strSetupSQLMail,          "N/A", "", strListType)
      Call SetParam("SetupSSISDB",           strSetupSSISDB,           "N/A", "", strListType)
      Call SetParam("SetupStdAccounts",      strSetupStdAccounts,      "N/A", "", strListType)
      Call SetParam("SetupSysDB",            strSetupSysDB,            "N/A", "", strListType)
      Call SetParam("SetupSysIndex",         strSetupSysIndex,         "N/A", "", strListType)
      Call SetParam("SetupSysManagement",    strSetupSysManagement,    "N/A", "", strListType)
    Case strEdition = "STANDARD"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListEdition)
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListEdition)
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListEdition)
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListEdition)
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListEdition)
      If (strSQLVersion = "SQL2016") And (strSPLevel < "SP1") Then
        Call SetParam("SetupAnalytics",      strSetupAnalytics,        "N/A", "", strListEdition)
        Call SetParam("SetupRServer",        strSetupRServer,          "N/A", "", strListEdition)
        Call SetParam("SetupPolyBase",       strSetupPolyBase,         "N/A", "", strListEdition)
        Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,"N/A", "", strListEdition)
      End If
      Call SetParam("SetMemoryOptimizedTempdbMetadata",   strSetMemOptTempdb,     "N/A", "", strListEdition)
    Case strEdition = "WEB"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListEdition)
      Call SetParam("SetupAPCluster",        strSetupAPCluster,        "N/A", "", strListEdition)
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListEdition)
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListEdition)
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListEdition)
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListEdition)
      Call SetParam("SetupDRUCtlr",          strSetupDRUCtlr,          "N/A", "", strListEdition)
      Call SetParam("SetupSQLAS",            strSetupSQLAS,            "N/A", "", strListEdition)
      Call SetParam("SetupSQLASCluster",     strSetupSQLASCluster,     "N/A", "", strListEdition)
      Call SetParam("SetupSQLIS",            strSetupSQLIS,            "N/A", "", strListEdition)
      Call SetParam("SetupSQLNS",            strSetupSQLNS,            "N/A", "", strListEdition)
      Call SetParam("SetupStreamInsight",    strSetupStreamInsight,    "N/A", "", strListEdition)
      If (strSQLVersion = "SQL2016") And (strSPLevel < "SP1") Then
        Call SetParam("SetupAnalytics",      strSetupAnalytics,        "N/A", "", strListEdition)
        Call SetParam("SetupRServer",        strSetupRServer,          "N/A", "", strListEdition)
        Call SetParam("SetupPolyBase",       strSetupPolyBase,         "N/A", "", strListEdition)
      End If
      Call SetParam("SetupISMasterCluster",  strSetupISMasterCluster,  "N/A", "", strListEdition)
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListEdition)
      Call SetParam("SetMemoryOptimizedHybridBufferpool", strSetMemOptHybridBP,   "N/A", "", strListEdition)
      Call SetParam("SetMemoryOptimizedTempdbMetadata",   strSetMemOptTempdb,     "N/A", "", strListEdition)
    Case strEdition = "WORKGROUP"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListEdition)
      Call SetParam("SetupAPCluster",        strSetupAPCluster,        "N/A", "", strListEdition)
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListEdition)
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListEdition)
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListEdition)
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListEdition)
      Call SetParam("SetupDRUCtlr",          strSetupDRUCtlr,          "N/A", "", strListEdition)
      Call SetParam("SetupSQLAS",            strSetupSQLAS,            "N/A", "", strListEdition)
      Call SetParam("SetupSQLASCluster",     strSetupSQLASCluster,     "N/A", "", strListEdition)
      Call SetParam("SetupSQLDBCluster",     strSetupSQLDBCluster,     "N/A", "", strListEdition)
      Call SetParam("SetupSQLIS",            strSetupSQLIS,            "N/A", "", strListEdition)
      Call SetParam("SetupSQLNS",            strSetupSQLNS,            "N/A", "", strListEdition)
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListEdition)
      Call SetParam("SetupStreamInsight",    strSetupStreamInsight,    "N/A", "", strListEdition)
      Call SetParam("SetupStretch",          strSetupStretch,          "N/A", "", strListEdition)
      If (strSQLVersion = "SQL2016") And (strSPLevel < "SP1") Then
        Call SetParam("SetupAnalytics",      strSetupAnalytics,        "N/A", "", strListEdition)
        Call SetParam("SetupRServer",        strSetupRServer,          "N/A", "", strListEdition)
        Call SetParam("SetupPolyBase",       strSetupPolyBase,         "N/A", "", strListEdition)
      End If
      Call SetParam("SetupISMasterCluster",  strSetupISMasterCluster,  "N/A", "", strListEdition)
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListEdition)
      Call SetParam("SetMemoryOptimizedHybridBufferpool", strSetMemOptHybridBP,   "N/A", "", strListEdition)
      Call SetParam("SetMemoryOptimizedTempdbMetadata",   strSetMemOptTempdb,     "N/A", "", strListEdition)
  End Select

End Sub


Sub GetExpressEditionData
  Call SetProcessId("0BCA", "Setup details for SQL Express Edition")

  Call GetExpressExe

  strTCPPortRS      = "8080"

  Call SetParam("SetupAlwaysOn",             strSetupAlwaysOn,         "N/A", "", strListEdition)
  Call SetParam("SetupAPCluster",            strSetupAPCluster,        "N/A", "", strListEdition)
  Call SetParam("SetupNoDriveIndex",         strSetupNoDriveIndex,     "N/A", "", strListEdition)
  Call SetParam("SetupSQLAgent",             strSetupSQLAgent,         "N/A", "", strListEdition)
  Call SetParam("SetupSQLAS",                strSetupSQLAS,            "N/A", "", strListEdition)
  Call SetParam("SetupSQLDBAG",              strSetupSQLDBAG,          "N/A", "", strListEdition)
  Call SetParam("SetupSQLIS",                strSetupSQLIS,            "N/A", "", strListEdition)
  Call SetParam("SetupSQLNS",                strSetupSQLNS,            "N/A", "", strListEdition)
  Call SetParam("SetupDistributor",          strSetupDistributor,      "N/A", "", strListEdition)
  Call SetParam("SetupDQ",                   strSetupDQ,               "N/A", "", strListEdition)
  Call SetParam("SetupDQC",                  strSetupDQC,              "N/A", "", strListEdition)
  Call SetParam("SetupGovernor",             strSetupGovernor,         "N/A", "", strListEdition)
  Call SetParam("SetupMDS",                  strSetupMDS,              "N/A", "", strListEdition)
  Call SetParam("SetupDRUCtlr",              strSetupDRUCtlr,          "N/A", "", strListEdition)
  Call SetParam("SetupDRUClt",               strSetupDRUClt,           "N/A", "", strListEdition)
  Call SetParam("SetupManagementDW",         strSetupManagementDW,     "N/A", "", strListEdition)
  Call SetParam("SetupSlipstream",           strSetupSlipstream,       "N/A", "", strListEdition)
  Call SetParam("SetupSSDTBI",               strSetupSSDTBI,           "N/A", "", strListEdition)
  Call SetParam("SetupStreamInsight",        strSetupStreamInsight,    "N/A", "", strListEdition)
  Call SetParam("SetMemoryOptimizedHybridBufferpool", strSetMemOptHybridBP,   "N/A", "", strListEdition)
  Call SetParam("SetMemoryOptimizedTempdbMetadata",   strSetMemOptTempdb,     "N/A", "", strListEdition)

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupSQLBC",            strSetupSQLBC,            "N/A", "", strListEdition)
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListEdition)
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "N/A", "", strListEdition)
    Case (strSQLVersion = "SQL2016") And (strSPLevel < "SP1")
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A", "", strListEdition)
      Call SetParam("SetupRServer",          strSetupRServer,          "N/A", "", strListEdition)
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A", "", strListEdition)
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListEdition)
  End Select

  Select Case True
    Case strExpVersion = "Basic" 
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListEdition)
      Call SetParam("SetupSQLDBFT",          strSetupSQLDBFT,          "N/A", "", strListEdition)
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "N/A", "", strListEdition)
      Call SetParam("SetupSSMS",             strSetupSSMS,             "N/A", "", strListEdition)
    Case strExpVersion = "With Tools" 
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListEdition)
      Call SetParam("SetupSQLDBFT",          strSetupSQLDBFT,          "N/A", "", strListEdition)
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "N/A", "", strListEdition)
  End Select

End Sub


Sub GetExpressExe()
  Call SetProcessId("0BCAA", "Get File Name for SQL Express")
  Dim colMediaFiles
  Dim strFileName

  strExpVersion     = ""
  strSQLExe         = GetParam(colFiles,              "SQLExpExe",          "",                    "")
  strDebugMsg1      = "Source: " & strPathSQLMedia
  Set objFolder     = objFSO.GetFolder(strPathSQLMedia)
  Set colMediaFiles = objFolder.Files
  For Each objFile In colMediaFiles
    strFileName     = objFile.name
    If strSQLExe <> "" Then
      strFileName   = strSQLExe
    End If
    Select Case True
      Case CheckExpressExe("Advanced",   strFileName)
        strSQLExe   = strFileName
      Case CheckExpressExe("With Tools", strFileName)
        strSQLExe   = strFileName
      Case CheckExpressExe("Basic",      strFileName)
        strSQLExe   = strFileName
    End Select
  Next

  strSQLExe         = Replace(strSQLExe, "ENU", strSQLLanguage, 1, -1, 1)
  Call DebugLog("File Name for SQL Express: " & strSQLExe)
  If strSQLExe = "" Then
    Call SetBuildMessage(strMsgErrorConfig, "No EXPRESS Edition install file found in " & strPathSQLMedia)
  End If

End Sub


Function CheckExpressExe(strVersion, strFileName)
  Call DebugLog("CheckExpressExe: " & strFileName & " for " & strVersion)
  Dim strExpExe
  Dim intExpFound

  intExpFound       = False
  Select Case True
    Case (strVersion = "Advanced") And (strSQLVersion = "SQL2005")
      strExpExe     = "SQLEXPR_ADV"
    Case strVersion = "Advanced"
      strExpExe     = "SQLEXPRADV"
    Case (strVersion = "With Tools") And (strSQLVersion = "SQL2005")
      strExpExe     = "UNKNOWN"
    Case strVersion = "With Tools"
      strExpExe     = "SQLEXPRWT"
    Case strVersion = "Basic"  
      strExpExe     = "SQLEXPR"
  End Select

  Select case True
    Case strExpExe = "UNKNOWN"
      ' Nothing
    Case UCase(Left(strFileName, Len(strExpExe))) <> UCase(strExpExe)
      ' Nothing
    Case UCase(Right(strFileName, 4)) <> ".EXE"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      intExpFound = True
    Case (Instr(UCase(strFileName), "X86") > 0) And (strWOWX86 = "TRUE")
      intExpFound = True
    Case Instr(UCase(strFileName), UCase(strFileArc)) > 0
      intExpFound = True
  End Select

  If intExpFound Then
    strExpVersion   = strVersion
    strEdType       = strVersion
  End If
  CheckExpressExe   = intExpFound

End Function


Sub GetClusterData()
  Call SetProcessId("0BD", "Get details of Windows cluster if it exists")

  Call SetClusterVars()
  Call CheckClusterGroups()
  Call CheckClusterSubnet()
  Call CheckClusterNetwork()
  Call CheckNetworkAdapters()

End Sub


Sub SetClusterVars()
  Call SetProcessId("0BDA", "Set Cluster variables")

  Select Case True
    Case strOSVersion >= "6.0"
      strClusterPath = strDirSys & "\Cluster\Reports"
      Call SetFileData("ClusterReport",    strClusterPath,      "", "Validation Report*.*ht*")
    Case Else
      strClusterPath = strPathSys & "LogFiles\Cluster"
      Call SetFileData("ClusterReport",    strClusterPath,      "", "ClCfgSrv.log")
  End Select
  strClusterReport  = GetBuildfileValue("ClusterReport")
  strClusterPath    = strClusterPath & "\" & strClusterReport

  If strSetupNetBind = "" Then
    Call SetParam("SetupNetBind",       strSetupNetBind,       "YES", "NetBind processing recommended for Cluster", "")
  End If
  If strSetupNetName = "" Then
    Call SetParam("SetupNetName",       strSetupNetName,       "YES", "NetName processing recommended for Cluster", "")
  End If

End Sub


Sub CheckClusterGroups()
  Call SetProcessId("0BDC", "Check Cluster Group details")
  Dim colClusGroups
  Dim objClusGroup
  Dim strGroupName

  Set colClusGroups = GetClusterGroups()
  For Each objClusGroup In colClusGroups
    strGroupName    = objClusGroup.Name
    Select Case True
      Case strGroupName = "Cluster Group"
        ' Nothing
      Case Else
        Call SetClusterFound(strGroupName)
    End Select
  Next

End Sub


Sub SetClusterFound(strGroupName)
  Call SetProcessId("0BDCA", "Set Cluster Found Flag for: " & strGroupName)

  Select Case True
    Case strGroupName = strClusterGroupAO
      strClusterAOFound = "Y"
    Case strGroupName = strClusterGroupAS
      strClusterASFound = "Y"
    Case strGroupName = strClusterGroupSQL
      strClusterSQLFound = "Y"
  End Select

  If GetGroupAction(strGroupName) = strActionClusInst Then
    Call MoveToNode(strGroupName, "")
  End If

End Sub


Sub CheckClusterSubNet()
  Call SetProcessId("0BDD", "Get Cluster Subnet Details")
  Dim arrDependencies, arrAddresses
  Dim colDependent
  Dim strNameResource, strPathAddresses, strPathDependent, strPathDependencies, strPathName

' ClusSubnet: S=Single subnet, M=Multiple subnets
  strClusSubnet     = "S"
  If strOSVersion < "6.2" Then
    Exit Sub
  End If

  strPathName       = "HKLM\Cluster\ClusterNameResource"
  strNameResource   = objShell.RegRead(strPathName)
  strDebugMsg1      = "Name Resource: " & strNameResource

  strPathDependencies  = "HKLM\Cluster\Dependencies\"
  objWMIReg.EnumKey strHKLM, Mid(strPathDependencies, 6), arrDependencies
  For Each colDependent In arrDependencies
    strPathDependent   = strPathDependencies & colDependent
    strDebugMsg2       = "Dependent: " & colDependent
    If objShell.RegRead(strPathDependent & "\Dependent") = strNameResource Then
      strPathAddresses = strPathDependent & "\"
      objWMIReg.GetMultiStringValue strHKLM, Mid(strPathAddresses, 6), "Provider List", arrAddresses
      If UBound(arrAddresses) > 0 Then
        strClusSubnet  = "M"
      End If
    End If
  Next

End Sub


Sub CheckClusterNetwork()
  Call SetProcessId("0BDE", "Get details for Cluster Network")
  Dim colClusNetworks, colCommonProps
  Dim objClusNetwork
  Dim intClient, intCluster
  Dim strClusNetworkRole, strClusRole

  Call GetClusterIP()

  intClient           = 0
  intCluster          = 0
  strClusNetworkRole  = ""
  Set colClusNetworks = GetClusterNetworks()

  For Each objClusNetwork In colClusNetworks
    Set colCommonProps   = objClusNetwork.CommonProperties
    strClusRole          = colCommonProps.Item("Role").Value
    Select Case True
      Case (strClusIPV4Network = objClusNetwork.Name) And (strClusIPVersion = "IPv4")
        strClusNetworkRole = strClusRole
        If strClusNetworkRole >= "2" Then
          intClient   = 1
        End If
      Case (strClusIPV6Network = objClusNetwork.Name) And (strClusIPVersion = "IPv6")
        strClusNetworkRole = strClusRole
        If strClusNetworkRole >= "2" Then
          intClient   = 1
        End If
      Case strClusRole = "1"
        intCluster    = 1
      Case strClusRole = "2" 
        intClient     = 1
      Case strClusRole = "3"
        intClient     = 1
        intCluster    = 1
    End Select
  Next

  If strClusIPVersion = "" Then
    Call SetBuildMessage(strMsgErrorConfig, "No Cluster IP Address found")
  End If
  If strClusNetworkRole < "2" Then
    Call SetBuildMessage(strMsgErrorConfig, "Primary Cluster Network must not be 'Cluster Only'")
  End If
  If intCluster = 0 Then
    Call SetBuildMessage(strMsgWarning, "No 'Cluster Only' Network found.  Best practice recommends that a network is dedicated to Cluster traffic only")
  End If
  If intClient = 0 Then
    Call SetBuildMessage(strMsgErrorConfig, "No Client Cluster Network found")
  End If

End Sub


Sub GetClusterIP()
  Call SetProcessId("0BDEA", "Get Cluster IP Details")
  Dim colClusNetworks, colResources, colNetInterfaces, colNetCProps, colNetIProps, colResourceProps
  Dim objResource, objNetInterface, objNetwork
  Dim strClusIPV4NetRole, strClusIPV6NetRole

  strClusIPVersion     = ""
  strClusIPV4Address   = ""
  strClusIPV4Network   = ""
  strClusIPV4NetRole   = ""
  strClusIPV6Address   = ""
  strClusIPV6Network   = ""
  strClusIPV6NetRole   = ""
  Set colClusNetworks  = GetClusterNetworks()
  Set colResources     = GetClusterResources()
  Set colNetInterfaces = GetClusterInterfaces()

  For Each objResource In colResources
    Select Case True
      Case Left(objResource.Name, 18) <> "Cluster IP Address" 
        ' Nothing
      Case objResource.CommonProperties.Item("Type").Value = "IP Address"
        Set colResourceProps = objResource.PrivateProperties
        For Each objNetInterface In colNetInterfaces
          strDebugMsg1       = "Interface: " & objNetInterface.Name
          Set colNetIProps   = objNetInterface.CommonROProperties
          Set colNetCProps   = Nothing
          For Each objNetwork In colClusNetworks
            If objNetwork.Name = colNetIProps.Item("Network").Value Then
              Set colNetCProps = objNetwork.CommonProperties
            End If
          Next
          Select Case True
            Case colResourceProps.Item("Network").Value <> colNetIProps.Item("Network").Value
              ' Nothing
            Case UCase(colNetIProps.Item("Node").Value) <> UCase(strServer)
              ' Nothing
            Case strClusIPV4NetRole >= CStr(colNetCProps.Item("Role").Value)
              ' Nothing
            Case Else
              strClusIPV4Network = colResourceProps.Item("Network").Value
              strClusIPV4Address = colResourceProps.Item("Address").Value
              strClusIPV4Mask    = colResourceProps.Item("SubnetMask").Value
              strClusIPV4NetRole = CStr(colNetCProps.Item("Role").Value)
          End Select
        Next
      Case strSQLVersion <= "SQL2005"
        ' Nothing
      Case objResource.CommonProperties.Item("Type").Value = "IPv6 Address"
        Set colResourceProps = objResource.PrivateProperties
        For Each objNetInterface In colNetInterfaces
          strDebugMsg1       = "Interface: " & objNetInterface.Name
          Set colNetIProps   = objNetInterface.CommonROProperties
          Set colNetCProps   = Nothing
          For Each objNetwork In colClusNetworks
            If objNetwork.Name = colNetIProps.Item("Network").Value Then
              Set colNetCProps = objNetwork.CommonProperties
            End If
          Next
          Select Case True
            Case colResourceProps.Item("Network").Value <> colNetIProps.Item("Network").Value
              ' Nothing
            Case UCase(colNetIProps.Item("Node").Value) <> UCase(strServer)
              ' Nothing
            Case strClusIPV6NetRole >= CStr(colNetCProps.Item("Role").Value)
              ' Nothing
            Case Else
              strClusIPV6Network = colResourceProps.Item("Network").Value
              strClusIPV6Address = colResourceProps.Item("Address").Value
              strClusIPV6Mask    = colResourceProps.Item("PrefixLength").Value
              strClusIPV6NetRole = CStr(colNetCProps.Item("Role").Value)
          End Select
        Next
    End Select
  Next
call debuglog("ClusterTCP: " & strClusterTCP)
call debuglog("ClusIPV4Network: " & strClusIPV4Network)
call debuglog("ClusIPV6Network: " & strClusIPV6Network)
  Select Case True
    Case (strClusterTCP = "IPV4") And (strClusIPV4Network <> "")
      strClusIPVersion = "IPv4"
      strClusIPAddress = strClusIPV4Address
    Case (strClusterTCP = "IPV4") And (strClusIPV6Network <> "")
      strClusIPVersion = "IPv6"
      strClusIPAddress = strClusIPV6Address
    Case (strClusterTCP = "IPV6") And (strClusIPV6Network <> "")
      strClusIPVersion = "IPv6"
      strClusIPAddress = strClusIPV6Address
    Case (strClusterTCP = "IPV6") And (strClusIPV4Network <> "")
      strClusIPVersion = "IPv4"
      strClusIPAddress = strClusIPV4Address
  End Select
call debuglog("ClusIPAddress: " & strClusIPAddress)
End Sub


Sub CheckNetworkAdapters()
  Call SetProcessId("0BDG", "Check Network Adapter Status")
  Dim colAdapters
  Dim objAdapter

  Set colAdapters   = objWMI.ExecQuery ("SELECT NetConnectionId from Win32_NetworkAdapter WHERE NetConnectionStatus >= 0 AND NetConnectionStatus <> 2")
  For Each objAdapter In colAdapters
    Call SetBuildMessage(strMsgErrorConfig, "Cannot install SQL Cluster because Network Adapter is not useable: " & objAdapter.NetConnectionId)
  Next

End Sub


Sub GetServerData()
  Call SetProcessId("0BE", "Get data for Role Servers")

  Call ParseRoleServer(strCatalogServer, strCatalogServerName, strCatalogInstance, strRSAlias)

  Select Case True
    Case strManagementServer = "" 
      strMSSupplied    = ""
    Case strSetupSQLDB <> "YES"
      strMSSupplied    = ""
    Case UCase(strManagementServer) = "YES"
      strManagementServer = strServer & "\" & strInstance
      strMSSupplied    = "Y"
    Case Else
      strMSSupplied    = "Y"
  End Select

  Select Case True
    Case IsNull(strManagementServerRes) 
      ' Nothing
    Case strManagementServerRes = ""
      ' Nothing
    Case Else
      strMSSupplied       = "Y"
      strManagementServer = strManagementServerRes
  End Select

  Call ParseRoleServer(strManagementServer, strManagementServerName, strManagementInstance, "")

  Select Case True
    Case strMSSupplied = ""
      SetManagementServerOptions("NO")
    Case (strManagementServerName = strUCServer)       And (strManagementInstance = strInstance)
      SetManagementServerOptions("YES")
    Case (strManagementServerName = strClusterNameSQL) And (strManagementInstance = strInstance)
      SetManagementServerOptions("YES")
    Case Else
      SetManagementServerOptions("NO")
  End Select

  strServerGroups   = GetAccountAttr(strServer, strUserDNSDomain, "memberOf")

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      Call SetParam("SetupAOAlias",          strSetupAOAlias,          "N/A" ,"/SetupAOAlias:Yes not allowed with /SetupAlwaysOn:No", "")
    Case strSetupAOAlias <> ""
      ' Nothing
    Case Else
      strSetupAOAlias = "YES"
  End Select
  strAOAliasOwner   = GetAddress(strGroupAO, "Alias", "")

  Select Case True
    Case strSetupAOAlias <> "YES"
      ' Nothing
    Case strGroupAO = strServer
      Call SetParam("SetupAOAlias",          strSetupAOAlias,          "NO"  ,"/AGName:" & strAGName & " matches server name", "")
  End Select

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      Call SetParam("SetupAOProcs",          strSetupAOProcs,          "N/A" ,"/SetupAOProcs:Yes not allowed with /SetupAlwaysOn:No", "")
    Case strSetupAOProcs <> ""
      ' Nothing
    Case Else
      strSetupAOProcs = "YES"
  End Select

  Call SetupInstRS()

End Sub


Sub ParseRoleServer(strRoleServer, strRoleServerName, strRoleInstance, strAlias)
  Call SetProcessId("0BEA", "ParseRoleServer: " & strRoleServer)
  Dim strPort

  strPort           = strTCPPort
  intIdx            = Instr(strRoleServer, "\")
  Select Case True
    Case intIdx > 0
      strRoleServerName = Left(strRoleServer, intIdx - 1)
      strRoleInstance   = Mid(strRoleServer, intIdx + 1)
    Case strRoleServer <> ""
      strRoleServerName = strRoleServer
      strRoleInstance   = strInstance
    Case (strClusterHost = "YES") And (strSetupAlwaysOn = "YES")
      strRoleServerName = strGroupAO
      strRoleInstance   = ""
    Case strSetupSQLDBCluster = "YES"
      strRoleServerName = strClusterNameSQL
      strRoleInstance   = strInstance
    Case Else
      strRoleServerName = strUCServer
      strRoleInstance   = strInstance
  End Select

  Select Case True
    Case Not (strRoleServerName = strClusterName Or strRoleServerName = strClusterBase) 
      ' Nothing
    Case strAlias <> ""
      strRoleServerName = strAlias
      strRoleInstance   = ""
    Case (strClusterHost = "YES") And (strSetupAlwaysOn = "YES")
      strRoleServerName = strGroupAO
      strRoleInstance   = ""
    Case strSetupSQLDBCluster = "YES"
      strRoleServerName = strClusterNameSQL
    Case Else
      strRoleServerName = strUCServer
  End Select
  intIdx            = Instr(strRoleInstance, ":")
  If intIdx > 0 Then
    strPort         = Mid(strRoleInstance, intIdx + 1)
    strRoleInstance = Left(strRoleInstance, (intIdx - 1))
  End If

  Select Case True
    Case strRoleInstance <> ""
      strRoleServer = strRoleServerName & "\" & strRoleInstance
    Case Else
      strRoleServer = strRoleServerName
  End Select
  If (strPort > "") And (strPort <> "1433") Then
    strRoleServer   = strRoleServer & ":" & strPort
  End If

  strRoleServerName = GetAddress(strRoleServerName, "", "Y")

End Sub


Sub SetManagementServerOptions(strMSOption)
  Call SetProcessId("0BEB", "Set Management Server Options: " & strMSOption)
  Dim strMessage

  Select Case True
    Case strMSOption = "YES"
      strMessage    = "installed by default with Management Server"
    Case Else
      strMessage    = "not installed by default except with Management Server"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDRUCtlr",          strSetupDRUCtlr,          "N/A",    "", strListSQLVersion)
    Case strSetupDRUCtlr = ""
      Call SetParam("SetupDRUCtlr",          strSetupDRUCtlr,          strMSOption,"Distributed Replay Controller " & strMessage, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A",    "", strListSQLVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A",    "", strListSQLDB)
    Case strSetupDQ = ""
      Call SetParam("SetupDQ",               strSetupDQ,               strMSOption,"Data Quality Services " & strMessage, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2017"
      Call SetParam("SetupISMaster",         strSetupISMaster,         "N/A",    "", strListSQLVersion)
    Case strSetupISMaster = ""
      Call SetParam("SetupISMaster",         strSetupISMaster,         strMSOption,"SSIS Scaleout Master " & strMessage, "")
  End Select

  Select Case True
    Case strSetupISMaster <> "YES"
      ' Nothing
    Case strIsWorkerMaster <> ""
      ' Nothing
    Case Else
      strIsWorkerMaster = strManagementServerName
  End Select    

  Select Case True
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A",    "", strListCore)
    Case strSQLVersion < "SQL2008R2"
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A",    "", strListSQLVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A",    "", strListSQLDB)
    Case strSetupMDS = ""
      Call SetParam("SetupMDS",              strSetupMDS,              strMSOption,"Master Data Services " & strMessage, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008"
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A",    "", strListSQLVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A",    "", strListSQLDB)
    Case strSetupManagementDW = "" 
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     strMSOption,"Management Data Warehouse " & strMessage, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A",    "", strListSQLVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A",    "", strListSQLDB)
    Case strSetupAnalytics = ""
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        strMSOption,"Advanced Analytics " & strMessage, "")
  End Select

  Select Case True
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A",    "", strListCore)
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A",    "", strListSQLVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "N/A",    "", strListSQLDB)
    Case strSetupPolyBase = "YES"
      ' Nothing
    Case Else
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         strMSOption,"PolyBase " & strMessage, "")
  End Select

  Select Case True
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strSetupPython = ""
      Call SetParam("SetupPython",           strSetupPython,           strMSOption,"Python " & strMessage, "")
  End Select

  Select Case True
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strSetupRServer = ""
      Call SetParam("SetupRServer",          strSetupRServer,          strMSOption,"R Server " & strMessage, "")
  End Select

End Sub


Sub GetAccountData()
  Call SetProcessId("0BF", "Get details of Windows accounts")

  Call GetLocalAccounts()

  Call GetServiceAccounts()

End Sub


Sub GetLocalAccounts()
  Call SetProcessId("0BFA", "Get Local Account details")

  Set objAccount    = objWMI.Get("Win32_SID.SID='S-1-5-32-559'") ' Performance Log Users
  strGroupPerfLogUsers = objAccount.AccountName

  Set objAccount    = objWMI.Get("Win32_SID.SID='S-1-5-32-558'") ' Performance Monitor Users
  strGroupPerfMonUsers = objAccount.AccountName

  Set objAccount    = objWMI.Get("Win32_SID.SID='S-1-5-32-555'") ' Remote Desktop Users
  strGroupRDUsers   = objAccount.AccountName

  Set objAccount    = objWMI.Get("Win32_SID.SID='S-1-1-0'")      ' Everyone
  strNTAuthEveryone = objAccount.AccountName
  Call SetBuildfileValue("NTAuthEveryone",          strNTAuthEveryone)

  strSIDDistComUsers = GetBuildfileValue("SIDDistComUsers")
  If strSIDDistComUsers = "" Then
    strSIDDistComUsers = "S-1-5-32-562"
  End If
  Set objAccount       = objWMI.Get("Win32_SID.SID='" & strSIDDistComUsers & "'") ' Distributed Com Users
  strGroupDistComUsers = objAccount.AccountName

  strSIDIISIUsers   = GetBuildfileValue("SIDIISIUsers")
  If strSIDIISIUsers = "" Then
    strSIDIISIUsers = "S-1-5-32-568"
  End If
  Set objAccount    = objWMI.Get("Win32_SID.SID='" & strSIDIISIUsers & "'") ' IIS Users
  strGroupIISIUsers = objAccount.AccountName

  If strNTAuthAccount = "" Then
    strNTAuthAccount = strNTAuth & "\" & strNTAuthNetwork
  End If 

  Select Case True
    Case strNTAuthAccount = strNTAuth & "\" & strNTAuthNetwork' Network Service
      strNTAuthOSName = strNTAuth & "\" & strNTAuthNetwork
    Case strNTAuthAccount = strNTAuth & "\" & objWMI.Get("Win32_SID.SID='S-1-5-18'").AccountName ' Local System
      strNTAuthOSName = "LocalSystem"
    Case strNTAuthAccount = strNTAuth & "\" & objWMI.Get("Win32_SID.SID='S-1-5-19'").AccountName ' Local Service
      strNTAuthOSName = strNTAuth & "\" & objWMI.Get("Win32_SID.SID='S-1-5-19'").AccountName
    Case Else
      strNTAuthOSName = strNTAuthAccount
  End Select 
  strNTAuthOSName   = UCase(strNTAuthOSName)

  strSecMain        = Replace(strSecMain, "Administrators", strGroupAdmin)
  strSecMain        = Replace(strSecMain, "Users",          strGroupUsers)
  strSecTemp        = Replace(strSecTemp, "Users",          strGroupUsers)

End Sub


Sub GetServiceAccounts()
  Call SetProcessId("0BFB", "Get Service Account details")
  Dim objAccountParm
  Dim strInst

  strDomainSID      = GetAccountAttr(strUserName, strUserDNSDomain, "objectSID")
  If strDomainSID > "" Then
    strDomainSID    = Mid(strDomainSid, Instr(strDomainSid, "-") + 1)
    strDomainSID    = Mid(strDomainSid, Instr(strDomainSid, "-") + 1)
    strDomainSID    = Mid(strDomainSid, Instr(strDomainSid, "-") + 1)
    strDomainSID    = Mid(strDomainSid, Instr(strDomainSid, "-") + 1)
    strDomainSID    = Left(strDomainSid, InstrRev(strDomainSid, "-") - 1)
    strDomainComputers = objWMI.Get("Win32_SID.SID='S-1-5-21-" & strDomainSID & "-515'").AccountName
    strDomainUsers     = objWMI.Get("Win32_SID.SID='S-1-5-21-" & strDomainSID & "-513'").AccountName
  End If
  Call SetBuildfileValue("DomainComputers", strDomainComputers)
  Call SetBuildfileValue("DomainUsers",     strDomainUsers)

  Call SetXMLParm(objAccountParm, "AccountParm",     "SqlSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "SqlAccount")
  Call SetXMLParm(objAccountParm, "GroupParm",       "SQLDomainGroup")
  Call SetXMLParm(objAccountParm, "GroupParmAlt",    "SQLClusterGroup")
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstSQL)
  Call SetXMLParm(objAccountParm, "Instance",        strInstSQL)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "SqlSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "SqlPassword")
  Call GetAccount("SQLDB",        "YES",             strSqlAccount,        strSqlPassword,         objAccountParm)

  intIdx            = Instr(strSqlAccount, "\")
  If intIdx > 0 Then
    strSqlAcDomain  = Left(strSqlAccount, IntIdx - 1)
  Else
    strSqlAcDomain  = ""
  End If

  Call SetXMLParm(objAccountParm, "AccountParm",     "AgtSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "AgtAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "GroupParm",       "AgtDomainGroup")
  Call SetXMLParm(objAccountParm, "GroupParmAlt",    "AgtClusterGroup")
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstAgent)
  Call SetXMLParm(objAccountParm, "Instance",        strInstAgent)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "AgtSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "AgtPassword")
  Call GetAccount("SQLDBAG",      strSetupSQLDBAG,   strAgtAccount,        strAgtPassword,         objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "AsSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "AsAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "GroupParm",       "ASDomainGroup")
  Call SetXMLParm(objAccountParm, "GroupParmAlt",    "ASClusterGroup")
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstAS)
  Call SetXMLParm(objAccountParm, "Instance",        strInstAS)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "AsSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "AsPassword")
  Call GetAccount("SQLAS",        strSetupSQLAS,     strAsAccount,         strAsPassword,          objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "BrowserSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "SqlBrowserAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strNTAuthAccount)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & "SQLBrowser")
  Call SetXMLParm(objAccountParm, "Instance",        "SQLBrowser")
  Call SetXMLParm(objAccountParm, "PasswordParm",    "BrowserSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "SqlBrowserPassword")
  Call GetAccount("SQLBrowser",   "YES",             strSqlBrowserAccount, strSqlBrowserPassword,  objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "CmdshellAccount")
  Call SetXMLParm(objAccountParm, "PasswordParm",    "CmdshellPassword")
  Call SetXMLParm(objAccountParm, "NoAccount",       "IGNORE")
  Call GetAccount("CmdShell",     strSetupCmdShell,  strCmdshellAccount,   strCmdshellPassword,    objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "CltSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTAuthAccount)
  Call SetXMLParm(objAccountParm, "Instance",        "SQL Server Distributed Replay Client")
  Call SetXMLParm(objAccountParm, "PasswordParm",    "CltScvPassword")
  Call GetAccount("DRUClt",       strSetupDRUClt,    strCltAccount,        strCltPassword,         objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "CtlrSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & "DRUCtlr")
  Call SetXMLParm(objAccountParm, "PasswordParm",    "CtlrSvcPassword")
  Call GetAccount("DRUCtlr",      strSetupDRUCtlr,   strCtlrAccount,    strCtlrPassword,     objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "ExtSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strNTService & "\" & strInstAnal)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstAnal)
  Call SetXMLParm(objAccountParm, "Instance",        strInstAnal)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "ExtSvcPassword")
  Call GetAccount("Analytics",    strSetupAnalytics, strExtSvcAccount,     strExtSvcPassword,      objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "FTSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "FtAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strNTAuthAccount)
  Call SetXMLParm(objAccountParm, "GroupParm",       "FTSDomainGroup")
  Call SetXMLParm(objAccountParm, "GroupParmAlt",    "FTSClusterGroup")
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstFT)
  Call SetXMLParm(objAccountParm, "Instance",        strInstFT)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "FTSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "FtPassword")
  Call GetAccount("SQLDBFT",      strSetupSQLDBFT,   strFTAccount,         strFTPassword,          objAccountParm)
 
  Call SetXMLParm(objAccountParm, "AccountParm",     "IsSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "IsAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Select Case True
    Case strSQLVersion <= "SQL2014"
      Call SetXMLParm(objAccountParm, "NTServiceAC", strNTAuthAccount)
    Case Else
      Call SetXMLParm(objAccountParm, "NTServiceAC", strNTService & "\" & strInstIS)
  End Select
  Call SetXMLParm(objAccountParm, "Instance",        strInstIS)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "IsSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "IsPassword")
  Call GetAccount("SQLIS",        strSetupSQLIS,     strIsAccount,         strIsPassword,          objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "ISMasterSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strIsAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strIsPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstISMaster)
  Call SetXMLParm(objAccountParm, "Instance",        strInstIsMaster)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "ISMasterSvcPassword")
  Call GetAccount("ISMaster",     strSetupISMaster,  strIsMasterAccount,   strIsMasterPassword,    objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "ISWorkerSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strIsAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strIsPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstISWorker)
  Call SetXMLParm(objAccountParm, "Instance",        strInstIsWorker)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "ISWorkerSvcPassword")
  Call GetAccount("ISWorker",     strSetupISWorker,  strIsWorkerAccount,   strIsWorkerPassword,    objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "MDSAccount")
  Call SetXMLParm(objAccountParm, "NoAccount",       "IGNORE") ' "ERROR") ' Temporary until next phase of MDS configuration added to FineBuild
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTAuthAccount)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "MDSPassword")
  Call GetAccount("MDS",          strSetupMDS,       strMDSAccount,        strMDSPassword,         objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "MDWAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strAgtAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strAgtPassword)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "MDWPassword")
  Call GetAccount("MDW",          strSetupManagementDW, strMDWAccount,     strMDWPassword,         objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "PBDMSSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & "PBDMSSvc")
  Call SetXMLParm(objAccountParm, "Instance",        strInstPM)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "PBDMSSvcPassword")
  Call GetAccount("PolyBase",     strSetupPolyBase,  strPBDMSSvcAccount,   strPBDMSSvcPassword,    objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "PBEngSvcAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strPBDMSSvcAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strPBDMSSvcPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & "PBEngSvc")
  Call SetXMLParm(objAccountParm, "Instance",        strInstPE)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "PBEngSvcPassword")
  Call GetAccount("PolyBase",     strSetupPolyBase,  strPBEngSvcAccount,   strPBEngSvcPassword,    objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "RsSvcAccount")
  Call SetXMLParm(objAccountParm, "AccountParmAlt",  "RsAccount")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strDefaultAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strDefaultPassword)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstRS)
  Call SetXMLParm(objAccountParm, "Instance",        strInstRS)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "RsSvcPassword")
  Call SetXMLParm(objAccountParm, "PasswordParmAlt", "RsPassword")
  Select Case True
    Case strOSVersion <> "6.2"
      Call GetAccount("SQLRS",    strSetupSQLRS,     strRsAccount,         strRsPassword,          objAccountParm)
    Case strRsPassword <> ""
      Call GetAccount("SQLRS",    strSetupSQLRS,     strRsAccount,         strRsPassword,          objAccountParm)
    Case strSetupPowerBI <> "YES"
      Call GetAccount("SQLRS",    strSetupSQLRS,     strRsAccount,         strRsPassword,          objAccountParm)
    Case Else
      strRsAccount   = strNTService & "\" & strInstRS
      strRsPassword  = ""
      objAccountParm = ""
  End Select

  Call SetXMLParm(objAccountParm, "AccountParm",     "RsExecAccount")
  Call SetXMLParm(objAccountParm, "NoAccount",       "NOSETUP")
  Call SetXMLParm(objAccountParm, "PasswordParm",    "RsExecPassword")
  Call GetAccount("RSExec",       strSetupRSExec,    strRSExecAccount,     strRSExecPassword,      objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "RsShareAccount")
  Call SetXMLParm(objAccountParm, "NoAccount",       "NOSETUP")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strRSExecAccount)
  Call SetXMLParm(objAccountParm, "DefaultPassword", strRSExecPassword)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "RsSharePassword")
  Call GetAccount("RSShare",      strSetupRSShare,   strRsShareAccount,    strRsSharePassword,     objAccountParm)

  Call SetXMLParm(objAccountParm, "AccountParm",     "TelSvcAcct")
  Call SetXMLParm(objAccountParm, "DefaultAC",       strNTService & "\" & strInstTel)
  Call SetXMLParm(objAccountParm, "NTServiceAC",     strNTService & "\" & strInstTel)
  Call SetXMLParm(objAccountParm, "Instance",        strInstTel)
  Call SetXMLParm(objAccountParm, "PasswordParm",    "TelSvcPassword")
  Call GetAccount("ISWorker",     "YES",             strTelSvcAcct,        strTelSvcPassword,      objAccountParm)

End Sub


Sub GetAccount(strSetupName, strSetup, strAccount, strPassword, objAccountParm)
  Call DebugLog("GetAccount: " & strSetupName)
  Dim strAccountParm, strAccountParmAlt, strAccountReqd, strDefaultAC, strDefaultPswd, strDomain, strDomainParm
  Dim strGroupParm, strGroupParmAlt, strNTServiceAC, strInstance, strNoAccount, strPasswordParm, strPasswordParmAlt

' XML Parameters
' Value             Default                 Description
' AccountParm                               Name of Parameter for Account
' AccountParmAlt                            Alternative Name of Parameter for Account
' GroupParm                                 Name of Parameter for Group containing Account
' GroupParmAlt                              Alternative Name of Parameter for Group containing Account
' DefaultAC                                 Value for Default Account
' DefaultPassword                           Value for Default Password
' DomainParm                                Name of Buildfile item to hold Domain if it is to be split from Account
' Instance                                  Instance Name to be interrogated
' NoAccount         "ERROR"                 Action to take if Account not found
' NTServiceAC                               Value for NT Service
' PasswordParm                              Name of Parameter for Password
' PasswordParmAlt                           Alternative Name of Parameter for Password
'
  strAccountParm    = GetXMLParm(objAccountParm, "AccountParm",     "")
  strAccountParmAlt = GetXMLParm(objAccountParm, "AccountParmAlt",  "")
  strGroupParm      = GetXMLParm(objAccountParm, "GroupParm",       "")
  strGroupParmAlt   = GetXMLParm(objAccountParm, "GroupParmAlt",    "")
  strDefaultAC      = GetXMLParm(objAccountParm, "DefaultAC",       "")
  strDefaultPswd    = GetXMLParm(objAccountParm, "DefaultPassword", "")
  strDomainParm     = GetXMLParm(objAccountParm, "DomainParm",      "")
  strInstance       = GetXMLParm(objAccountParm, "Instance",        "")
  strNoAccount      = GetXMLParm(objAccountParm, "NoAccount",       "ERROR")
  strNTServiceAC    = GetXMLParm(objAccountParm, "NTServiceAC",     "")
  strPasswordParm   = GetXMLParm(objAccountParm, "PasswordParm",    "")
  strPasswordParmAlt= GetXMLParm(objAccountParm, "PasswordParmAlt", "")

  strAccount        = ""
  strPassword       = ""
  Select Case True	' Get Account details from Parameter or Service Discovery
    Case colArgs.Exists(strAccountParm) OR colArgs.Exists(strAccountParmAlt)
      strAccount    = GetParam(Null,                  strAccountParm,       strAccountParmAlt,     strDefaultAC)
      strPassword   = GetParam(Null,                  strPasswordParm,      strPasswordParmAlt,    "")
    Case strInstance <> ""
      strPath       = "SYSTEM\CurrentControlSet\Services\" & strInstance & "\"
      objWMIReg.GetStringValue strHKLM,strPath,"ObjectName",strAccount
      If IsNull(strAccount) Then
        strAccount  = ""
      End If
  End Select

  Select Case True	' Initial check on Account
    Case strSetupName = "SQLDB" 
      Call CheckSqlAccount(strAccountParm, strAccountParmAlt, strAccount)
    Case strAccount > ""
      ' Nothing
    Case Else
      strAccount    = strDefaultAC
      strPassword   = strDefaultPswd
  End Select

  If strAccount = strNTService Then
    strAccount      = strNTServiceAC
    strPassword     = ""
  End If

  Select Case True	' Force Setup for CONFIG or DISCOVER process
    Case strAccount = ""
      ' Nothing
    Case Instr("CONFIG DISCOVER", strType) > 0
      strSetup      = "YES"
  End Select

  Select Case True	' Check account details
    Case strSetup <> "YES"
      ' Nothing
    Case strAccount <> ""
      strAccount    = CheckAccount(strAccountParm, strAccountParmAlt, strAccount, "Y")
    Case strNoAccount = "IGNORE"
      ' Nothing
    Case strNoAccount = "NOSETUP"
      strSetup      = "NO"
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/" & strAccountParm & ": must be specified")
  End Select

  intIdx            = InStr(strAccount, "\")
  Select Case True
    Case strDomainParm = ""
      ' Nothing
    Case intIdx = 0
      ' Nothing
    Case Else
      strDebugmsg1  = "Account: " & strAccount
      strDomain     = Left(strAccount, intIdx - 1)
      strAccount    = Mid(strAccount, intIdx + 1)
      Call SetBuildfileValue(strDomainParm,      strDomain)
  End Select

  Select Case True
    Case strAccountParmAlt <> ""
      Call SetBuildfileValue(strAccountParmAlt,  strAccount)
    Case Else
      Call SetBuildfileValue(strAccountParm,     strAccount)
  End Select
  Select Case True
    Case strPasswordParmAlt <> ""
      Call SetBuildfileValue(strPasswordParmAlt, strPassword)
    Case Else
      Call SetBuildfileValue(strPasswordParm,    strPassword)
  End Select

  Select Case True
    Case strGroupParm = ""
      ' Nothing
    Case strClusterHost <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2005"
      Call SetGroupParm(strAccount, strGroupParm, strGroupParmAlt)
    Case strOSVersion < "6"
      Call SetGroupParm(strAccount, strGroupParm, strGroupParmAlt)
  End Select

  objAccountParm    = ""

End Sub


Sub CheckSqlAccount(strAccountParm, strAccountParmAlt, strSqlAccount)
  Call DebugLog("CheckSqlAccount:")

  Select Case True
    Case (strSQLVersion <= "SQL2008R2") And (strSQLAccount = "")
      strSqlAccount      = strNTAuthAccount
      strSqlPassword     = ""
      strDefaultAccount  = strSqlAccount
      strDefaultPassword = strSQLPassword
    Case (strOSVersion < "6.1") And (strSQLAccount = "")
      strSqlAccount      = strNTAuthAccount
      strSqlPassword     = ""
      strDefaultAccount  = strSqlAccount
      strDefaultPassword = strSQLPassword
    Case strOSVersion < "6.1" 
      strDefaultAccount  = strSqlAccount
      strDefaultPassword = strSQLPassword
    Case strSqlAccount = ""
      strSqlAccount      = strNTService & "\" & strInstSQL  ' Local Virtual Account
      strSqlPassword     = ""
      strDefaultAccount  = strNTService
      strDefaultPassword = strSQLPassword
    Case Left(strSQLAccount, Len(strNTService) + 1) = strNTService & "\" ' Local Virtual Account
      strSqlPassword     = ""
      strDefaultAccount  = strNTService
      strDefaultPassword = strSQLPassword
    Case Else
      strDefaultAccount  = strSqlAccount
      strDefaultPassword = strSqlPassword
  End Select

  strSqlAccount     = CheckAccount(strAccountParm, strAccountParmAlt, strSqlAccount, "N")

  If GetBuildfileValue("SqlAccountType") = "M" Then        ' Domain Managed Account
    strSqlPassword     = ""
    strDefaultAccount  = strSqlAccount
    strDefaultPassword = strSQLPassword
  End If

End Sub


Function CheckAccount(strAccountParm, strAltParm, strUserAccount, strVerify)
  Call DebugLog("CheckAccount: " & strAccountParm & " Account: " & strUserAccount)
  Dim intIdx
  Dim strAccount, strAccountDom, strAccountVar
' AccountType: S=Local Service,L=Local User,M=Domain Managed,G=Domain User

  strAccount        = strUserAccount
  intIdx            = Instr(strAccount, "\")
  Select Case True
    Case intIdx = 0 
      strAccountDom = ""
    Case Left(strAccount, intIdx) = ".\"
      strAccountDom = strServer
      strAccount    = strAccountDom & Mid(strAccount, 2)
    Case Else
      strAccountDom = Left(strAccount, intIdx - 1)
  End Select

  Select Case True
    Case strAltParm = ""
      strAccountVar = strAccountParm
      Call SetBuildfileValue(strAccountVar & "Name", strAccountParm)
    Case strSQLVersion <= "SQL2005"
      strAccountVar = strAltParm
      Call SetBuildfileValue(strAccountVar & "Name", strAltParm)
    Case Else
      strAccountVar = strAltParm
      Call SetBuildfileValue(strAccountVar & "Name", strAccountParm)
  End Select

  Select Case True
    Case strAccount = ""
      Call SetBuildfileValue(strAccountVar & "Type", "")
    Case GetAccountAttr(strAccount, strUserDNSDomain, "msDS-GroupMSAMembership") <> ""
      Call CheckMSAAccount(strAccountVar, strAccount)
      strListMSA    = GetBuildfileValue("ListMSA")
      If Instr(" " & strListMSA, " " & strAccountVar & " ") = 0 Then
        Call SetBuildfileValue("ListMSA", strListMSA & strAccountVar & " ")
      End If
      Call SetBuildfileValue(strAccountVar & "Group", Trim(Mid(GetAccountAttr(strAccount, strUserDNSDomain, "msDS-GroupMSAMembership"), 2)))
      Call SetBuildfileValue(strAccountVar & "Type", "M")
    Case StrComp(strAccountDom, strNTAuth, vbTextCompare) = 0
      Call SetBuildfileValue(strAccountVar & "Type", "S")
    Case StrComp(strAccountDom, strNTService, vbTextCompare) = 0
      Call SetBuildfileValue(strAccountVar & "Type", "S")
    Case StrComp(strAccountDom, strServer, vbTextCompare) = 0
      If GetAccountAttr(strAccount, strAccountDom, "name") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "Account " & strUserAccount & " can not be found")
      End If
      Call SetBuildfileValue(strAccountVar & "Type", "L")
    Case Else
      If GetAccountAttr(strAccount, strUserDNSDomain, "name") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "Account " & strUserAccount & " can not be found")
      End If
      Call SetBuildfileValue(strAccountVar & "Type", "D")
  End Select

  Select Case True
    Case strVerify <> "Y"
      ' Nothing
    Case GetBuildfileValue(strAccountVar & "Type") <> "M"
      ' Nothing
    Case strOSVersion <= "6.1"
      Call SetBuildMessage(strMsgErrorConfig, strAccountVar & " - Domain Managed Account can not be used on " & strOSName)
    Case strSQLVersion <= "SQL2008R2"
      Call SetBuildMessage(strMsgErrorConfig, strAccountVar & " - Domain Managed Account can not be used on " & strSQLVersion)
  End Select

  Call DebugLog(" Account found: " & strAccount)
  CheckAccount      = strAccount

End Function


Sub CheckMSAAccount(strAccountParm, strAccount)
  Call DebugLog("CheckMSAAccount: " & strAccount)
  Dim intIdx
  Dim strAccountName

  intIdx            = Instr(strAccount, "\") + 1
  strAccountName    = Mid(strAccount, intIdx)

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetBuildMessage(strMsgErrorConfig, "Domain Managed Accounts cannot be used with " & strSQLVersion)
    Case Len(strAccountName) > 15
      Call SetBuildMessage(strMsgErrorConfig, "/" & strAccountParm & " must be 15 characters or less")
    Case strSQLVersion >= "SQL2016"
      ' Nothing
    Case Right(strAccount, 1) <> "$" 
      Call SetBuildMessage(strMsgErrorConfig, "/" & strAccountParm & " MSA Account must end with $")
  End Select
 
End Sub


Sub SetGroupParm(strAccount, strGroupParm, strGroupParmAlt)
  Call DebugLog("GetGroupParm: " & strAccount)
  Dim strGroupName, strGroupList

  strGroupName      = GetParam(Null,                  strGroupParm,         strGroupParmAlt  ,     "")
  strGroupList      = GetAccountAttr(strAccount, strUserDNSDomain, "memberOf")
  Select Case True
    Case strGroupName <> ""
      ' Nothing
    Case strGroupList = ""
      ' Nothing
    Case Else
      strGroupList  = Trim(Replace(strGroupList, strDomainUsers, "")) & " "
      strGroupName  = strDomain & "\" & Left(strGroupList, Instr(strGroupList, " ") - 1)
  End Select

  Select Case True
    Case strGroupName <> ""
      ' Nothing
    Case strSQLVersion = "SQL2005"
      Call SetBuildMessage(strMsgErrorConfig, "/" & strGroupParmAlt & ": must be specified")
    Case strGroupParm = "FTSDomainGroup"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/" & strGroupParm & ": must be specified")
  End Select

  Call SetBuildfileValue(strGroupParm, strGroupName)

End Sub


Sub GetGroupData()
  Call SetProcessId("0BG", "GetGroupData:")

  Call GetDBAGroups()

  strListMSA        = Trim(GetBuildfileValue("ListMSA"))
  If strListMSA <> "" Then
    Call GetMSAGroups(strListMSA)
  End If

End Sub


Sub GetDBAGroups()
  Call SetProcessId("0BGA", "Get DBA Groups")

  strGroupDBA       = UCase(GetParam(colGlobal,       "GroupDBA",           "",                    ""))
  strGroupDBA       = CheckGroup(strGroupDBA) 
  Select Case True
    Case strGroupDBA <> ""
      strSecDBA     = strSecMain & " """ & FormatAccount(strGroupDBA) & """:F "
    Case Else
      Call SetBuildMessage(strMsgWarning, "/GroupDBA value can not be found so using local Administators")
      strGroupDBA   = CheckGroup(strLocalAdmin)
      strSecDBA     = Replace(strSecMain, " ", "  ") ' Add extra space into SecDBA so it is different to SecMain
  End Select

  Call Util_RunExec("WHOAMI /GROUPS /FO CSV | FIND /C """ & strGroupDBA & """", "", "", -1)
  strIsInstallDBA   = CStr(intErrSave)
  If strSQLVersion >= "SQL2012" Then
    strMembersDBA   = GetGroupMembers(strGroupDBA, 1)
  End If

  Select Case True
    Case strUserDNSDomain = ""
      strUserAccount = strServer & "\" & strUserName
    Case Else
      strUserAccount = strDomain & "\" & strUserName
  End Select

  strGroupDBANonSA  = UCase(GetParam(colGlobal,       "GroupDBANonSA",      "",                    ""))
  strGroupDBANonSA  = CheckGroup(strGroupDBANonSA) 

End Sub


Function CheckGroup(strGroupParm)
  Call DebugLog("CheckGroup: " & strGroupParm)
  Dim intFound
  Dim strGroup, strGroupDom, strGroupName

  strGroup          = ""
  intFound          = False
  intIdx            = Instr(strGroupParm, "\")
  Select Case True
    Case intIdx = 0
      strGroupDom   = strDomain
      strGroupName  = strGroupParm
      intFound      = CheckGroupExists(strGroupDom, strGroupName)
    Case Left(strGroupParm, intIdx) = ".\"
      strGroupDom   = strServer
      strGroupName  = Mid(strGroupParm,  intIdx + 1)
      intFound      = CheckGroupExists(strGroupDom, strGroupName)
    Case Else
      strGroupDom   = Left(strGroupParm, intIdx - 1)
      strGroupName  = Mid(strGroupParm,  intIdx + 1)
      intFound      = CheckGroupExists(strGroupDom, strGroupName)
  End Select

  If intFound = False Then
    strGroupDom     = strServer
    intFound        = CheckGroupExists(strGroupDom, strGroupName)
  End If

  Select Case True
    Case intFound = True
      strGroup      = strGroupDom & "\" & strGroupName
    Case Else
      strGroup      = ""
  End Select

  CheckGroup = strGroup 

End Function


Function CheckGroupExists(strGroupDom, strGroupName)
  Call DebugLog("CheckGroupExists: " & strGroupDom & " " & strGroupName)
  Dim intFound

  intFound          = -1
  Select Case True
    Case strGroupName = ""
      ' Nothing
    Case strGroupDom = "BUILTIN"
      Call DebugLog("Check BUILTIN for " & strServer)
      Call Util_RunExec("NET LOCALGROUP """ & strGroupName & """ ", "", "", -1)
      intFound      = intErrSave
      strGroupDom   = strBuiltinDom
    Case strGroupDom = strServer
      Call DebugLog("Check Server " & strServer)
      Call Util_RunExec("NET LOCALGROUP """ & strGroupName & """ ", "", "", -1)
      intFound      = intErrSave
    Case strUserDNSDomain <> ""
      Call DebugLog("Check Domain " & strUserDNSDomain)
      Call Util_RunExec("NET GROUP """ & strGroupName & """ /Domain", "", "", -1)
      intFound      = intErrSave
    Case Else
      Call DebugLog("Check Workgroup")
      Call Util_RunExec("NET LOCALGROUP """ & strGroupName & """ ", "", "", -1)
      intFound      = intErrSave
      strGroupDom   = strServer
  End Select

  Select Case True
    Case intFound = 0
      CheckGroupExists = True
    Case Else
      CheckGroupExists = False
  End Select

End Function


Function GetGroupMembers(strGroup, intLevel)
  Call DebugLog("GetGroupMembers: " & strGroup)
  Dim strGroupDom, strGroupName, strMembers

  intIdx            = Instr(strGroup, "\")
  strGroupDom       = Left(strGroup, intIdx - 1)
  strGroupName      = Mid(strGroup,  intIdx + 1)
  If strGroupDom = "BUILTIN" Then
    strGroupDom     = strServer
  End If

  Select Case True
    Case UCase(strGroupDom) = UCase(strServer)
      strMembers    = GetLocalMembers(strGroupDom, strGroupName, intLevel)
    Case Else
      strMembers    = GetADMembers(strGroupDom, strGroupName, intLevel)
  End Select

  Call DebugLog("Members of Group " & strGroup & ": " & strMembers)
  GetGroupMembers   = strMembers

End Function


Function GetLocalMembers(strGroupDom, strGroupName, intLevel)
  Call DebugLog("GetLocalMembers: " & strGroupDom & "\" & strGroupName)
  Dim objGroup, objMember
  Dim strMembers, strMemberName

  strMembers        = ""
  Set objGroup      = GetObject("WinNT://" & strGroupDom & "/" & strGroupName)
  Select Case True
    Case objGroup.Class = "User"
      strMembers = strGroup
    Case objGroup.Class = "Group"
      For Each objMember In objGroup.Members
        strMemberName = Mid(objMember.Parent, InstrRev(objMember.Parent, "/") + 1) & "\" & objMember.Name
        Select Case True
          Case (objMember.Class = "Group") And (intLevel <= 5)
            strMembers = strMembers & " " & GetGroupMembers(strMemberName, intLevel + 1)
          Case objMember.Class <> "User" 
            ' Nothing
          Case CheckUser(strGroupDom, strMemberName) = True
            strMembers = strMembers & """" & strMemberName & """ "
        End Select
      Next
  End Select

  Set objMember     = Nothing
  Set objGroup      = Nothing
  GetLocalMembers   = RTrim(strMembers)

End Function


Function GetADMembers(strGroupDom, strGroupName, intLevel)
  Call DebugLog("GetADMembers: " & strGroupDom & "\" & strGroupName)
' Based on: https://gallery.technet.microsoft.com/scriptcenter/b160d928-fb9e-4c49-a194-f2e5a3e806ae
  Dim objAdsPath, objMember, objADRecords
  Dim strMemberDom, strMembers

  strCmd            = "SELECT AdsPath FROM '" & strADRoot & "' WHERE sAMAccountName='" & strGroupName & "'"
  Set objADRecords  = WScript.CreateObject("ADODB.Recordset")
  objADRecords.Open strCmd, objADOConn, 3
  Set objAdsPath    = GetObject(objADRecords("AdsPath"))
  For Each objMember In objAdsPath.Members
    strMemberDom    = UCase(objMember.AdsPath)
    strMemberDom    = Mid(strMemberDom, Instr(strMemberDom, "DC=") + 3)
    intIdx          = Instr(strMemberDom, ",")
    If intIdx > 0 Then
      strMemberDom  = Left(strMemberDom, IntIdx - 1)
    End If
    Select Case True
      Case (objMember.Class = "Group") And (intLevel <= 5) 
        strMembers  = strMembers & " " & GetGroupMembers(strMemberDom & "\" & objMember.sAMAccountName, intLevel + 1)
      Case objMember.Class <> "User" 
        ' Nothing
      Case CheckUser(strGroupDom, objMember.sAMAccountName) = True
        strMembers  = strMembers & " " & strMemberDom & "\" & objMember.sAMAccountName
    End Select
  Next

  Set objADRecords  = Nothing
  Set objAdsPath    = Nothing
  Set objMember     = Nothing
  GetADMembers      = RTrim(strMembers)

End Function


Function CheckUser(strGroupDom, strUser)
  Call DebugLog("CheckUser: " & strGroupDom & "\" &  strUser)
  On Error Resume Next

  Set objUser       = GetObject("WinNT://" & strGroupDom & "/" & strUser)
  Select Case True
    Case Err.Number = 0
      CheckADUser   = True
    Case Else
      CheckADUser   = False
  End Select

  On Error GoTo 0

End Function


Sub GetMSAGroups(strListMSA)
  Call SetProcessId("0BGB", "GetMSAGroups: " & strListMSA)
  ReDim arrAccounts(2,0), arrGroupMSA(1,0)
  Dim arrItems
  Dim intAccount, intGroupMSA, intGroupUse, intIdx, intItems, intUAccounts, intUGroupMSA, intUItems

  Call DebugLog("Build arrays of Accounts and Groups")
  arrItems          = Split(strListMSA, " ")
  intUItems         = UBound(arrItems)
  For intItems = 0 To intUItems
    Call SetupGroupArray(arrAccounts, arrGroupMSA, arrItems(intItems))
  Next 

  If strServerGroups <> "" Then
    Call SetupGroupArray(arrAccounts, arrGroupMSA, strServer)
  End If
  If strClusGroups <> "" Then
    Call SetupGroupArray(arrAccounts, arrGroupMSA, strClusterName)
  End If

  Call DebugLog("Identify common Group")
  intGroupUse       = 0
  strGroupMSA       = ""
  intUGroupMSA      = UBound(arrGroupMSA, 2)
  For intGroupMSA = 0 To intUGroupMSA
    If arrGroupMSA(1, intGroupMSA) > intGroupUse Then
      intGroupUse   = arrGroupMSA(1, intGroupMSA)
      strGroupMSA   = arrGroupMSA(0, intGroupMSA)
      intIdx        = InStr(strGroupMSA, "\")
      If intIdx > 0 Then
        strGroupMSA = Mid(strGroupMSA, intIdx + 1)
      End If
    End If
  Next

  Call DebugLog("Identify Accounts and Servers without common Group")
  Select Case True
    Case intGroupUse < (intUGroupMSA / 2)
      Call SetBuildMessage(strMsgErrorConfig, "Cannot find common Windows Group for MSA Accounts and Servers")
    Case Else
      intUAccounts  = UBound(arrAccounts, 2)
      For intAccount = 0 To intUAccounts
        If Instr(" " & arrAccounts(2, intAccount) & " ", strGroupMSA & " ") = 0 Then ' Set reminder that Item must be added to MSA Group
          Call SetBuildfileValue(arrAccounts(0, intAccount) & "MSAGroup", strGroupMSA)
        End If
      Next
  End Select

End Sub


Sub SetupGroupArray(arrAccounts, arrGroupMSA, strItem)
  Call DebugLog("SetupGroupArray: " & strItem)
  Dim arrGroups
  Dim intAccount, intGroup, intGroupMSA, intUAccounts, intUGroups, intUGroupMSA
  Dim strAccount, strAccountName, strGroups, strGroup

  strAccount        = GetBuildfileValue(strItem)
  strGroups         = GetBuildfileValue(strItem & "Group")

  intUAccounts  = UBound(arrAccounts, 2)
  For intAccount = 0 To intUAccounts
    Select Case True
      Case strAccount = ""
        Exit For
      Case intAccount = intUAccounts ' Build array of accounts
        arrAccounts(0, intAccount) = strItem
        arrAccounts(1, intAccount) = strAccount
        arrAccounts(2, intAccount) = strGroups
        arrGroups                  = Split(strGroups)
        intUGroups                 = UBound(arrGroups, 1)
        intUAccounts = intUAccounts + 1
        ReDim Preserve arrAccounts(2, intUAccounts)
        For intGroup = 0 To intUGroups
          strGroup     = arrGroups(intGroup)
          intUGroupMSA = UBound(arrGroupMSA, 2)
          For intGroupMSA = 0 To intUGroupMSA ' Build array of Groups and usage count
            Select Case True
              Case strGroup = ""
                Exit For
              Case strGroup = strDomainComputers
                ' Nothing
              Case intGroupMSA = intUGroupMSA
                arrGroupMSA(0, intGroupMSA) = strGroup
                arrGroupMSA(1, intGroupMSA) = 1
                intUGroupMSA = intUGroupMSA + 1
                ReDim Preserve arrGroupMSA(1, intUGroupMSA)
              Case strGroup = arrGroupMSA(0, intGroupMSA)
                arrGroupMSA(1, intGroupMSA) = arrGroupMSA(1, intGroupMSA) + 1
                Exit For
            End Select
          Next
        Next
      Case strAccount = arrAccounts(1, intAccount)
        Exit For
    End Select
  Next

End Sub


Sub GetVolumeData()
  Call SetProcessId("0BH", "Get data for Storage Volumes")

  Call SetLocalVolCodes()

  If strClusterHost <> "" Then
    Call SetClusterVolCodes()
  End If

  Call SetNetworkVolCodes()

  Call GetInstallVolumes()

End Sub


Sub SetLocalVolCodes()
  Call SetProcessId("0BHA", "Set Local Volume codes")
  Dim colVol
  Dim objVol
  Dim strVolume, strVolSource, strVolType
' VolSource: C=CSV, D=Disk, M=Mount Point, N=Mapped Network Drive, S=Share
' VolType:   C=Clustered, L=Local, X=Either

  strCmd            = "SELECT * FROM Win32_LogicalDisk"
  Set colVol        = objWMI.ExecQuery(strCmd)
  For Each objVol In colVol
    strVolume       = UCase(Left(objVol.DeviceId, 1))
    strVolSource    = GetBuildfileValue("Vol" & strVolume & "Source")
    strVolType      = GetBuildfileValue("Vol" & strVolume & "Type")
    If strVolType = "" Then
      strVolType    = "L"
    End If
    Select Case True
      Case strVolSource = "M"
        ' Nothing
      Case objVol.DriveType = 3 ' Local Disk
        strDriveList = strDriveList & strVolume
        Call SetBuildfileValue("Vol" & strVolume & "Source",  "D")
        Call SetBuildfileValue("Vol" & strVolume & "Type",    strVolType)
      Case objVol.DriveType = 4 ' Network Share
        strDriveList = strDriveList & strVolume
        Call SetBuildfileValue("Vol" & strVolume & "Source",  "N")
        Call SetBuildfileValue("Vol" & strVolume & "Type",    strVolType)
        Call SetBuildfileValue("Vol" & strVolume & "Path",    objVol.ProviderName)
      Case objVol.DriveType = 6 ' RAM Disk
        strDriveList = strDriveList & strVolume
        Call SetBuildfileValue("Vol" & strVolume & "Source",  "D")
        Call SetBuildfileValue("Vol" & strVolume & "Type",    strVolType)
    End Select
    Call CheckDriveSize (objVol, objVol.DeviceId)
    Call CheckDriveSpace(objVol, objVol.DeviceId)
  Next

End Sub


Sub SetClusterVolCodes()
  Call SetProcessId("0BHB", "Set Cluster Volume codes")
  Dim colClusPartitions, colResources
  Dim objClusDisk, objClusPartition, objResource
  Dim strResName, strVolName, strVolSource
' VolSource: C=CSV, D=Disk, M=Mount Point, N=Mapped Network Drive, S=Share
' VolType:   C=Clustered, L=Local, X=Either

  Set colResources  = GetClusterResources()
  For Each objResource In colResources
    If objResource.TypeName = "Physical Disk" Then
      Call DebugLog("Processing volume: " & objResource.Name)
      Set objClusDisk       = objResource.Disk
      Set colClusPartitions = objClusDisk.Partitions
      For Each objClusPartition In colClusPartitions
        Select Case True
          Case objClusPartition.FileSystem = "CSVFS"
            strResName   = objResource.Name
            strVolName   = GetCSVFolder(strResName)
            strVolSource = "C"
          Case Left(objClusPartition.DeviceName, 11) = "\\?\Volume{"
            strResName   = Mid(objClusPartition.DeviceName, InStr(objClusPartition.DeviceName, "{") + 1)
            strResName   = Left(strResName, Instr(strResName, "}") - 1)
            strVolName   = strResName
            strVolSource = "M"
          Case Else
            strResName   = objResource.Name
            strVolName   = Left(objClusPartition.DeviceName, 1)
            strVolSource = "D"
        End Select
        strDebugMsg1 = "Vol Name:   " & strVolName
        strDebugMsg2 = "Vol Source: " & strVolSource
        Call SetBuildfileValue("Vol_" & UCase(strVolName) & "Res",  strResName)
        Call SetBuildfileValue("Vol" & UCase(strVolName) & "Source",  strVolSource)
        Call SetBuildfileValue("Vol" & UCase(strVolName) & "Type",    "C")
        Call SetResourceOn(strResName, "")
      Next
    End If
  Next

End Sub


Function GetCSVFolder(strResName)
  Call DebugLog("GetCSVFolder: " & strResName)
  Dim arrResources, arrVolumeNames, arrVolumeParts
  Dim colResource
  Dim strName, strPathResource, strPathResources, strResFolder, strResType

  strResFolder      = Replace(strResName, " ", "")
  strPathResources  = "HKLM\Cluster\Resources\"
  objWMIReg.EnumKey strHKLM, Mid(strPathResources, 6), arrResources
  For Each colResource In arrResources
    strPathResource = strPathResources & colResource
    strResType      = objShell.RegRead(strPathResource & "\Type")
    strName         = objShell.RegRead(strPathResource & "\Name")
    Select Case True
      Case strResType <> "Physical Disk"
        ' Nothing
      Case strName <> strResName
        ' Nothing
      Case Else
        objWMIReg.GetMultiStringValue strHKLM, Mid(strPathResource, 6), "VolumeNames", arrVolumeNames
        arrVolumeParts = Split(arrVolumeNames(0), " ")
        strResFolder   = arrVolumeParts(3)
    End Select
  Next 

  GetCSVFolder      = strResFolder

End Function


Sub CheckDriveSize(objVol, strVol)
  Call DebugLog("CheckDriveSize: " & strVol)
  Dim intSizeVol
  Dim strSizeReq, strSizeUnit

  Call GetSpaceData("Size", strVol, intSizeVol, strSizeUnit, strSizeReq, objVol.Size)
  Select Case True
    Case strSizeReq = ""
     ' Nothing
    Case strSizeUnit <> ""
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case (strProcessId >= "2") And (strProcessId < "7")
      ' Nothing
    Case IsNumeric(intSizeReq) = False
      ' Nothing
    Case Int(intSizeVol * 0.99) < Int(intSizeReq)
      Call SetBuildMessage(strMsgErrorConfig, "Volume " & strVol & " size is too small. " & strSizeReq & strSizeUnit & " required, " & Cstr(Int(intSizeVol)) & strSizeUnit & " found.")
    Case Else
      Call SetBuildMessage(strMsgInfo,  "Volume " & strVol & " OK, Size required " & strSizeReq & strSizeUnit & ", Size found " & Cstr(Int(intSizeVol)) & strSizeUnit)
  End Select

End Sub


Sub CheckDriveSpace(objVol, strVol)
  Call DebugLog("CheckDriveSpace: " & strVol)
  Dim intSpaceVol
  Dim strSpaceReq, strSpaceUnit

  Call GetSpaceData("Space", strVol, intSpaceVol, strSpaceUnit, strSpaceReq, objVol.FreeSpace)
  Select Case True
    Case strSpaceReq = ""
     ' Nothing
    Case strSpaceUnit <> ""
      ' Nothing
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case (strProcessId >= "2") And (strProcessId < "7")
      ' Nothing
    Case IsNumeric(intSpaceReq) = False
      ' Nothing
    Case Int(intSpaceVol * 0.99) < Int(intSpaceReq)
      Call SetBuildMessage(strMsgErrorConfig, "Volume " & strVol & " does not have enough free space. " & strSpaceReq & strSpaceUnits & " required, " & Cstr(Int(intSpaceVol)) & strSpaceUnit & " found.")
    Case Else
      Call SetBuildMessage(strMsgInfo,  "Volume " & strVol & " OK, space required " & strSpaceReq & strSpaceUnit & ", space found " & Cstr(Int(intSpaceVol)) & strSpaceUnit)
  End Select

End Sub


Sub GetSpaceData(strUnitType, strUnitVol, intUnitValue, strUnitSize, strUnitReq, intUnitBase)
  Call DebugLog("GetSpaceData: " & strUnitType)
  Dim strUnitParm

  strUnitSize       = ""
  strUnitParm       = Ucase(GetParam(Null, "Vol" & strUnitType & Left(strUnitVol, 1), "Drv" & strUnitType & Left(strUnitVol, 1), ""))
  If strUnitParm > "" Then
    strUnitSize     = UCase(Right(strUnitParm, 2))
    Select Case True
      Case Instr("KB MB GB TB", strUnitSize) > 0 
        strUnitReq  = RTrim(Left(strUnitParm, Len(strUnitParm) - 2))
      Case Else
        strUnitSize = "MB"
        strUnitReq  = RTrim(strUnitParm)
    End Select

    intUnitValue    = intUnitBase + 1
    Select Case True
      Case strUnitSize = "KB"
        intUnitValue = intUnitValue / 1024
      Case strUnitSize = "MB"
        intUnitValue = intUnitValue / 1024 / 1024
      Case strUnitSize = "GB"
        intUnitValue = intUnitValue / 1024 / 1024 / 1024
      Case strUnitSize = "TB"
        intUnitValue = intUnitValue / 1024 / 1024 / 1024 / 1024
    End Select
  End If

End Sub


Sub SetNetworkVolCodes()
  Call SetProcessId("0BHC", "Set Network Volume codes")
  Dim arrDrives
  Dim objDrive
  Dim strDrive, strRemotePath

  strPath           = strUserSID & "\Network"
  objWMIReg.EnumKey strHKU, strPath, arrDrives
  Select Case True
    Case IsNull(arrDrives)
      ' Nothing
    Case Else
      For Each objDrive In arrDrives
        strDriveList = strDriveList & objDrive
        objWMIReg.GetStringValue strHKU,strPath & "\" & objDrive,"RemotePath",strRemotePath
        strDrive     = UCase(objDrive)
        Call SetBuildfileValue("Vol" & strDrive & "Source",  "N")
        Call SetBuildfileValue("Vol" & strDrive & "Type",    "L")
        Call SetBuildfileValue("Vol" & strDrive & "Path",    strRemotePath)
      Next      
  End Select

End Sub


Sub GetInstallVolumes()
  Call SetProcessId("0BHD", "Get volumes needed for SQL install")
' VolType:   C=Clustered, L=Local, X=Either

  strVolSys         = GetVolumes("VolSys",       "L", "N", objShell.ExpandEnvironmentStrings("%SYSTEMDRIVE%"), 1, "")

  Select Case True
    Case strClusterAction <> ""
      strVolProg    = GetVolumes("VolProg",      "L", "N", strVolSys,  1, "")
    Case Else  
      strVolProg    = GetVolumes("VolProg",      "L", "Y", strVolSys,  1, "")
  End Select

  strVolRoot        = GetVolumes("VolRoot",      "",  "Y", strVolProg, 1, strClusterAction)
  strMountRoot      = strVolRoot & ":\" & strMountRoot & "\"
  strVolDBA         = GetVolumes("VolDBA",       "X", "Y", strVolProg, 1, "")
  strVolDRU         = GetVolumes("VolDRU",       "X", "Y", strVolDBA,  1, "")
  strVolMDW         = GetVolumes("VolMDW",       "X", "Y", strVolDBA,  1, "")
  strVolTempWin     = GetVolumes("VolTempWin",   "L", "Y", strVolProg, 1, "")

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call GetSQLDBVolumes()
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      Call GetSQLASVolumes()
  End Select

  Call GetDTCVolumes()

End Sub


Sub GetSQLDBVolumes()
  Call SetProcessId("0BHDA", "Get SQL DB Volumes")

  strVolBackup      = GetVolumes("VolBackup",    "X", "Y", strVolRoot, 0, strActionSQLDB)
  strVolData        = GetVolumes("VolData",      "",  "Y", strVolRoot, 0, strActionSQLDB)
  Select Case True
    Case strSQLVersion < "SQL2008"
      ' Nothing
    Case strFSLevel = "0"
      ' Nothing
    Case strSQLVersion >= "SQL2017"
      strVolDataFS  = GetVolumes("VolDataFS",    "X", "Y", strVolData, 1, strActionSQLDB)
    Case Else
      strVolDataFS  = GetVolumes("VolDataFS",    "",  "Y", strVolData, 1, strActionSQLDB)
  End Select
  strVolDataFT      = GetVolumes("VolDataFT",    "",  "Y", strVolData, 1, strActionSQLDB)
  strVolLog         = GetVolumes("VolLog",       "",  "Y", strVolProg, 0, strActionSQLDB)
  strVolSysDB       = GetVolumes("VolSysDB",     "",  "Y", strVolData, 1, strActionSQLDB)
  Select Case True
    Case (strSetupSQLDBCluster = "YES") And (strSQLVersion >= "SQL2012")
      strVolTemp    = GetVolumes("VolTemp",      "X", "Y", strVolData, 0, strActionSQLDB)
      strVolLogTemp = GetVolumes("VolLogTemp",   "X", "Y", strVolLog,  1, strActionSQLDB)
    Case Else
      strVolTemp    = GetVolumes("VolTemp",      "",  "Y", strVolData, 0, strActionSQLDB)
      strVolLogTemp = GetVolumes("VolLogTemp",   "",  "Y", strVolLog,  1, strActionSQLDB)
  End Select
  If strSQLVersion >= "SQL2014" Then
    strVolBPE       = GetVolumes("VolBPE",       "L", "Y", strVolTemp, 1, strActionSQLDB)
  End If

End Sub


Sub GetSQLASVolumes()
  Call SetProcessId("0BHDB", "Get SQL AS Volumes")

  strVolRootAS      = GetVolumes("VolRootAS",    "",  "Y", "",           1, strActionSQLAS)
  Select Case True
    Case strVolRootAS <> ""
      strVolBackupAS = GetVolumes("VolBackupAS", "X", "Y", strVolRootAS, 0, strActionSQLAS)
      strVolDataAS  = GetVolumes("VolDataAS",    "",  "Y", strVolRootAS, 0, strActionSQLAS)
      strVolLogAS   = GetVolumes("VolLogAS",     "",  "Y", strVolRootAS, 0, strActionSQLAS)
      strVolTempAS  = GetVolumes("VolTempAS",    "",  "Y", strVolDataAS, 1, strActionSQLAS)
    Case Else
      strVolBackupAS = GetVolumes("VolBackupAS", "X", "Y", strVolBackup, 0, strActionSQLAS)
      strVolDataAS  = GetVolumes("VolDataAS",    "",  "Y", strVolData,   0, strActionSQLAS)
      strVolLogAS   = GetVolumes("VolLogAS",     "",  "Y", strVolLog,    0, strActionSQLAS)
      strVolTempAS  = GetVolumes("VolTempAS",    "",  "Y", strVolTemp,   1, strActionSQLAS)
  End Select

End Sub


Function GetVolumes(strVolParam, strVolReq, strGetParam, strVolDefault, intVolNum, strItemAction)
  Call DebugLog("GetVolumes: " & strVolParam)
  Dim arrItems
  Dim intIdx, intUBound
  Dim strItem, strReq, strVolList
' Req: C=Must be Cluster volume, L=Must be Local volume, X=Either type of volume
' Source: C=CSV, D=Disk, M=Mount Point, N=Mapped Network Drive, S=Share
' Type:   C=Clustered, L=Local, X=Either

  Select Case True
    Case strVolReq <> ""
      strReq        = strVolReq
    Case strItemAction = strActionClusInst
      strReq        = "C"
    Case strItemAction = "ADDNODE"
      strReq        = "C"
    Case Else  
      strReq        = "L"
  End Select

  Select Case True
    Case strGetParam = "N"
      strVolList    = strVolDefault
    Case strType = "FULLPROG"
      strVolList    = strVolDefault
    Case Else
      strVolList    = GetParam(colBuild, strVolParam, "Drv" & Mid(strVolParam, 4), strVolDefault)
  End Select
  arrItems          = Split(strVolList, ",")
  intUBound         = UBound(arrItems)
  Call SetBuildfileValue(strVolParam & "Req", strReq)
  strVolFoundList   = ""
  For intIdx = 0 To intUBound
    strItem         = Trim(arrItems(intIdx))
    Select Case True
      Case (intVolNum > 0) And (intIdx + 1 > intVolNum)
        arrItems(intIdx) = ""
      Case strVolParam = "VolSys"
        If intVolNum > 0 Then
          strItem   = Left(strItem, intVolNum)
          arrItems(intIdx) = strItem
        End If
        Call CheckDriveLetter(strVolParam, strItem, intVolNum, "")
      Case CheckCSV(strVolParam, strItem) <> 0
        Call SetBuildfileValue(strVolParam & "Source",  "C")
        Call SetBuildfileValue(strVolParam & "Type",    "C")
        strCSVFound = "Y"
      Case CheckMountPoint(strVolParam, strItem) <> 0
        Call SetBuildfileValue(strVolParam & "Source",  "M")
        Call SetBuildfileValue(strVolParam & "Type",    "L")
      Case CheckShare(strVolParam, strItem) <> 0
        Call SetBuildfileValue(strVolParam & "Source",  "S")
        Call SetBuildfileValue(strVolParam & "Type",    "L")
      Case Else
        Call CheckDriveLetter(strVolParam, strItem, intVolNum, strItemAction)
        arrItems(intIdx) = strItem
    End Select
  Next

  strVolList        = Join(arrItems, ",")
  If Right(strVolList, 1) = "," Then
    strVolList      = Left(strVolList, Len(strVolList) - 1)
  End If
  Call SetBuildfileValue(strVolParam, strVolList)
  GetVolumes        = strVolList

End Function


Function CheckCSV(strVolParam, strVolume)
  Call DebugLog("CheckCSV: " & strVolParam & " for " & strVolume)
  Dim intCSVFound
  Dim strCSVFolder, strCSVName

  intCSVFound       = 0
  strCSVFolder      = UCase(strVolume)
  If Right(strCSVFolder, 1) = "\" Then
    strCSVFolder    = Left(strCSVFolder, Len(strCSVFolder) - 1)
  End If
  If Left(strCSVFolder, Len(strCSVRoot)) = strCSVRoot Then
    strCSVFolder    = Mid(strCSVFolder, Len(strCSVRoot) + 1)
    If InStr(strCSVFolder, "\") > 0 Then
      strCSVFolder  = Left(strCSVFolder, InStr(strCSVFolder, "\") - 1)
    End If
  End If

  Select Case True
    Case strOSVersion < "6.1"
      ' Nothing
    Case Instr(strCSVFolder, "\") > 0
      ' Nothing
    Case GetBuildfileValue("Vol" & strCSVFolder & "Source") = "C" 
      intCSVFound   = 1
  End Select

  If intCSVFound = 1 Then
    strCSVName      = GetBuildfileValue("Vol_" & strCSVFolder & "Name")
    If strCSVName = "" Then
      Call SetBuildfileValue("Vol_" & strCSVFolder & "Name", strVolParam)
    End If
  End If

  CheckCSV          = intCSVFound

End Function


Function CheckMountPoint(strVolParam, strVolume)
  Call DebugLog("CheckMountPoint: " & strVolParam & " for " & strVolume)
  Dim colMountPoints
  Dim objMountPoint
  Dim intMPFound
  Dim strMPDir, strMPVol

  intMPFound        = 0
  Select Case True
    Case Len(strVolume) = 1
      ' Nothing
    Case Left(strVolume, 2) = "\\"
      ' Nothing
    Case Instr(Ucase(strOSName), " XP") <> 0 
      ' Nothing
    Case Else
      Set colMountPoints = objWMI.ExecQuery("SELECT * FROM Win32_MountPoint")
      For Each objMountPoint In colMountPoints
        strMPDir    = objMountPoint.Directory
        strMPDir    = Replace(strMPDir, "\\", "\")
        strMPVol    = Mid(objMountPoint.Volume, InStr(objMountPoint.Volume, "{") + 1)
        strMPVol    = Left(strMPVol, InStr(strMPVol, "}") - 1)
        strDebugMsg1  = strMPVol
        strDebugMsg2  = strVolume
        If StrComp(strMPVol, strVolume, 1) = 0 Then
          intMPFound  = 1
        End If
      Next
  End Select

  CheckMountPoint   = intMPFound

End Function


Function CheckShare(strVolParam, strVolume)
  Call DebugLog("CheckShare: " & strVolParam & " for " & strVolume)

  Select Case True
    Case Left(strVolume, 2) = "\\"
      CheckShare    = BuildShareList(strVolume)
    Case Else
      CheckShare    = 0
  End Select

End Function


Function BuildShareList(strVolume)
  Call DebugLog("BuildShareList: " & strVolume)
  Dim strShare, strShareList, strReadAll, strRemoteServer, strRemoteShare
  Dim intFound, intDelim

  intDelim          = Instr(3, strVolume, "\")
  If intDelim = 0 Then
    Call SetBuildMessage(strMsgErrorConfig, "Unable to find Share for " & strVolume)
  End If

  strRemoteServer   = Left(strVolume, intDelim - 1)
  strRemoteShare    = Mid(strVolume, intDelim + 1)
  If Instr(strRemoteShare, "\") > 0 Then
    strRemoteShare  = Left(strRemoteShare, Instr(strRemoteShare, "\") - 1)
  End If
  strRemoteRoot     = strRemoteServer & "\" & strRemoteShare

  intFound          = 0
  strShare          = ""
  strShareList      = GetBuildfileValue("ShareList")
  intFound          = GetShare(strRemoteServer, strRemoteShare, strShare, strRemoteRoot)
  If intFound = 0 Then
    WScript.Sleep strWaitShort
    intFound        = GetShare(strRemoteServer, strRemoteShare, strShare, strRemoteRoot)
  End If

  Select Case True
    Case intFound = 0
      Call SetBuildMessage(strMsgErrorConfig, "Unable to find Share on " & strRemoteRoot & " for " & strVolume)
    Case strShare <> strRemoteRoot
      ' Nothing
    Case Instr(strShareList, strRemoteRoot) > 0
      ' Nothing
    Case Else
      strShareList  = LTrim(strShareList & "," & strShare)
  End Select

  Call SetBuildfileValue("ShareList", strShareList)
  BuildShareList    = intFound

End Function


Function GetShare(strRemoteServer, strRemoteShare, strShare, strRemoteRoot)
  Call DebugLog("GetShare: " & strRemoteServer & ", " & strRemoteShare)
  Dim arrReadAll
  Dim objExec
  Dim strReadAll, strServerWork, strShareWork, strWorkline
  Dim intFound, intLines

  intFound          = 0
  strDebugMsg1      = "Remote Server:" & strRemoteServer
  strDebugMsg2      = "Remote Share:" & strRemoteShare

  strCmd            = "NET VIEW " & strRemoteServer
  Call DebugLog(strCmd)
  Set objExec       = objShell.Exec(strCmd)
  strReadAll        = Replace(objExec.StdOut.ReadAll, vbLf, "")
  arrReadAll        = Filter(Split(strReadAll, vbCr), " ")
  intLines          = UBound(arrReadAll)
  Call DebugLog("NET VIEW output:" & Cstr(intLines) & ">" & Join(arrReadAll, "< >") & "<")

  If intLines > 2 Then
    For intIdx = 2 To intLines - 1
      strWorkLine   = arrReadAll(intIdx)
      strServerWork = Left(strWorkLine, Instr(strWorkLine, " ") - 1)
      strShareWork  = RTrim(Left(strWorkLine, Len(strRemoteShare) + 1))
      Select Case True
        Case UCase(strShareWork) = UCase(strRemoteShare)
          intFound  = 1
          strShare  = strRemoteRoot
        Case UCase("\\" & strServerWork) = UCase(strRemoteServer)
          intFound  = 1
      End Select
    Next
  End If

  GetShare          = intFound

End Function


Sub CheckDriveLetter(strVolParam, strVolList, intVolNum, strItemAction)
  Call DebugLog("CheckDriveLetter: " & strVolList)
  Dim intIdx, intShare
  Dim strVolume, strVolPath, strVolSource, strVolType

  strVolume         = UCase(Left(strVolList, 1))
  strVolSource      = GetBuildfileValue("Vol" & strVolume & "Source")
  strVolType        = GetBuildfileValue("Vol" & strVolume & "Type")
  Select Case True
    Case StrComp(strCSVRoot, Left(strVolList, Len(strCSVRoot)), 1) = 0
      Call SetBuildMessage(strMsgErrorConfig, "/" & strVolParam & ": can not be found: " & strVolList)
    Case strVolSource = "N"
      strVolPath    = GetBuildfileValue("Vol" & strVolume & "Path")
      strVolList    = strVolPath & Mid(strVolList, 3)
    Case strItemAction = "ADDNODE"
      ' Nothing
    Case Else
      Call CheckVolExists(strVolParam, strVolList, intVolNum, strItemAction)
  End Select

  Call SetBuildfileValue(strVolParam & "Source",  strVolSource)
  Call SetBuildfileValue(strVolParam & "Type",    strVolType)

End Sub


Sub CheckVolExists(strVolParam, strVolList, intVolNum, strItemAction)
  Call DebugLog("CheckVolExists: " & strVolList)
  Dim strVolume, strVolReq

  strVolReq         = GetBuildFileValue(strVolParam & "Req")
  For intIdx = 1 To Len(strVolList)
    strVolume       = UCase(Mid(strVolList, intIdx, 1))
    strPath         = strVolume & ":\"
    Select Case True
      Case InStr(strVolErrorList, strVolume) > 0
        ' Nothing
      Case objFSO.FolderExists(strPath)
        strVolFoundList = strVolFoundList & strVolume
      Case (strVolReq = "C") And (strItemAction = "ADDNODE")
        strVolFoundList = strVolFoundList & strVolume
      Case (strVolReq = "X") And (strItemAction = "ADDNODE")
        strVolFoundList = strVolFoundList & strVolume
      Case GetBuildfileValue("Vol" & strVolume & "Source") = "D"
        strVolFoundList = strVolFoundList & strVolume
      Case Else
        strVolErrorList = strVolErrorList & strVolume
        Call SetBuildMessage(strMsgErrorConfig, "Drive for /" & strVolParam & ": can not be found: " & strPath)
    End Select
  Next

  If intVolNum > 0 Then
    strVolFoundList = Left(strVolFoundList, intVolNum)
  End If
  strVolList        = strVolFoundList

End Sub


Sub GetPIDData()
  Call SetProcessId("0BI", "Get SQL Server PID")

  Select Case True
    Case strPID <> ""
      ' Nothing
    Case strSQLVersion <= "SQL2005"
      ' Nothing
    Case strEdition = "EXPRESS"
      ' Nothing
    Case Else
      strPID        = GetPID(strPathSQLMedia)
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2008R2"
      ' Nothing
    Case strEdition = "EXPRESS"
      ' Nothing
    Case strSetupStreamInsight <> "YES"
      ' Nothing
    Case strStreamInsightPID <> ""
      ' Nothing
    Case strSQLVersion <= "SQL2014"
      strStreamInsightPID = strPID
    Case Else
      strStreamInsightPID = GetPID(GetMediaPath(GetSQLMediaPath("SQL2014", "")))
  End Select

End Sub


Function GetPID(strPathSQLMedia)
  Call DebugLog("GetPID: " & strPathSQLMedia)
  Dim strFile, strFileText, strPID

  strFile           = FormatFolder(strPathSQLMedia & "x64") & "DefaultSetup.ini"
  strPID            = ""

  If objFSO.FileExists(strFile) Then
    Set objFile     = objFSO.OpenTextFile(strFile, 1, False, GetFileType(strFile))
    strFileText     = objFile.ReadAll()
    objFile.Close
    intIdx          = Instr(strFileText, "PID=")
    If intIdx > 0 Then
      strPid        = Mid(strFileText, intIdx + 5, 29)
    End If
  End If

  Select Case True
    Case strPID <> ""
      Call DebugLog("PID Found")
    Case Else
      Call DebugLog("Using Evaluation PID")
      strPID        = "00000-00000-00000-00000-00000"
  End Select

  GetPID            = strPID

End Function


Function GetFileType(strFile)
  Call DebugLog("GetFileType: " & strFile)
' Adapted from https://groups.google.com/forum/#!topic/microsoft.public.scripting.vbscript/Yo-T-EMAAKU
  Dim objFile
  Dim strC1, strC2

  Set objFile       = objFSO.OpenTextFile(strFile)
  strC1             = objFile.Read(1)
  strC2             = objFile.Read(1)
  objFile.Close

  Select Case True
    Case Asc(strC1) <> 255
      GetFileType   = False ' ASCII
    Case Asc(strC2) <> 254
      GetFileType   = False ' ASCII
    Case Else
      GetFileType   = True  ' Unicode
  End Select

End Function


Sub GetMenuData()
  Call SetProcessId("0BJ", "Get data for SQL Menus")

  Call GetSQLVersionMenus()

  Call SetBuildfileValue("MenuAdminTools",  GetMenu("MenuAdminTools",      "Administrative Tools"))
  Call SetBuildfileValue("MenuBOL",         GetMenu("MenuBOL",             "SQL Server Books Online"))
  Call SetBuildfileValue("MenuConfigTools", GetMenu("MenuConfigTools",     "Configuration Tools"))
  Call SetBuildfileValue("MenuPerfTools",   GetMenu("MenuPerfTools",       "Performance Tools"))
  Call SetBuildfileValue("MenuPrograms",    GetMenu("MenuPrograms",        "Programs"))
  Call SetBuildfileValue("MenuSQLAS",       GetMenu("MenuSQLAS",           "Analysis Services"))
  Call SetBuildfileValue("MenuSQLIS",       GetMenu("MenuSQLIS",           "Integration Services"))
  Call SetBuildfileValue("MenuSQLNS",       GetMenu("MenuSQLNS",           "Notification Services Command Prompt"))
  Call SetBuildfileValue("MenuSQLRS",       GetMenu("MenuSQLRS",           "Reporting Services"))

  Select Case True
    Case strOSVersion < "6.2"
      Call SetBuildfileValue("MenuAccessories", GetMenu("MenuAccessories", "Accessories"))
      Call SetBuildfileValue("MenuSystem",      GetMenu("MenuSystem",      "Accessories"))
    Case Else
      Call SetBuildfileValue("MenuAccessories", GetMenu("MenuAccessories", "Windows Accessories"))
      Call SetBuildfileValue("MenuSystem",      GetMenu("MenuSystem",      "Windows System"))
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetBuildfileValue("MenuSQLDocs", GetMenu("MenuSQLDocs",         "Documentation and Tutorials"))
    Case Else
      Call SetBuildfileValue("MenuSQLDocs", GetMenu("MenuSQLDocs",         "Documentation & Community"))
  End Select

  If strEdition = "EXPRESS" Then
    Call SetBuildfileValue("MenuSSMS", GetMenu("MenuSSMSExp",              "SQL Server Management Studio Express"))
  Else
    Call SetBuildfileValue("MenuSSMS", GetMenu("MenuSSMS",                 "SQL Server Management Studio"))
  End If

End Sub


Sub GetSQLVersionMenus()
  Call DebugLog("GetSQLVersionMenus:")
  Dim arrSQL
  Dim intUBound
  Dim strDefaultMenu, strVersion, strVersionMenu

  arrSQL            = Split(strSQLList, " ", -1)
  intUBound         = UBound(arrSQL)
  For intIdx = 0 To intUBound
    strVersion      = arrSQL(intIdx)
    strDefaultMenu  = Mid(strVersion, 4)
    If Len(strDefaultMenu) > 4 Then
      strDefaultMenu = Left(strDefaultMenu, 4) & " " & Mid(strDefaultMenu, 5)
    End If
    strVersionMenu  = GetMenu("Menu" & strVersion, "Microsoft SQL Server " & strDefaultMenu)
    Call SetBuildfileValue("Menu" & strVersion, strVersionMenu)
    Call SetBuildfileValue("Menu" & strVersion & "Flag",  CheckMenuExists("Menu" & strVersion, strVersionMenu))
    If strSQLVersion = strVersion Then
      Call SetBuildfileValue("MenuSQL", strVersionMenu)
    End If
  Next

End Sub


Function GetMenu(strMenu, strDefault)
  Call DebugLog("GetMenu: " & strMenu)
  Dim strMenuText

  strMenuText       = GetBuildfileValue(strMenu)
  
  Select Case True
    Case strMenuText <> ""
      ' Nothing
    Case Else
      strMenuText   = GetParam(colStrings,            strMenu,              "",                    strDefault)
  End Select

  GetMenu           = strMenuText

End Function


Function CheckMenuExists(strMenu, strMenuText)
  Call DebugLog("CheckMenuExists: " & strMenu)
  Dim strFlag

  strFlag           = GetBuildfileValue(strMenu & "Flag")
  strPath           = strAllUserProf & "\" & GetBuildfileValue("MenuPrograms") & "\" & strMenuText

  Select Case True
    Case strFlag <> ""
      ' Nothing
    Case objFSO.FileExists(strPath) 
      strFlag       = "Y"
    Case Else
      strFlag       = "N"
  End Select

  CheckMenuExists   = strFlag

End Function


Sub GetPathData()
  Call SetProcessId("0BK", "Get data for Folder paths")
  Dim strRegPath, strSSISVol, strWorkPathl, strWorkVol

  Select Case True
    Case strVolProg = strVolSys
      strDirProg    = strVolProg & Mid(strDirProgSys, 2) & "\" & strSQLProgDir
    Case Else
      strPath       = GetParam(colStrings,            "DirProg",            "",                    Mid(strDirProgSys, 4))
      strDirProg    = strVolProg & ":\" & strPath & "\" & strSQLProgDir
  End Select
  Call SetFolderPath("DirProg",           "VolProg",     "",           "",         strDirProg,                   "")

  strDirSQLBootstrap = strDirProg & "\" & strSQLVersionNum & "\Setup Bootstrap\Log\"

  Select Case True
    Case strSQLVersion = "SQL2008"
      If (strClusterAction <> "") And (strSetupSlipstream <> "DONE") Then
        Call SetParam("SetupSlipstream",     strSetupSlipstream,       "YES", "Slipstream Media must be configured for Cluster installs", "")
      End If 
    Case strSQLVersion = "SQL2008R2"
      Select Case True
        Case (strClusterAction <> "") And (strSetupSlipstream <> "DONE") 
          Call SetParam("SetupSlipstream",   strSetupSlipstream,       "YES", "Slipstream Media must be configured for Cluster installs", "")
        Case Instr(strOSType, "CORE") > 0  
          Call SetParam("SetupSlipstream",   strSetupSlipstream,       "N/A", "", strListCore)
      End Select
    Case strSPLevel < "RTM"
      Call SetParam("SetupSlipstream",       strSetupSlipstream,       "YES", "Slipstream Media must be configured for Pre-RTM installs", "")
    Case strSQLVersion >= "SQL2012"
      Call SetParam("SetupSlipstream",       strSetupSlipstream,       "N/A", "", strListSQLVersion)
  End Select
  
  Select Case True
    Case strSetupSlipstream = "DONE"
      strPathSQLSP  = ""
      strPCUSource  = GetMediaPath(GetParam(colStrings,   "PCUSource",      "",                    strPathSQLMedia & "PCU"))
      If strPCUSource = "" Then
        strPCUSource  = GetMediaPath(GetParam(colStrings, "PCUSource",      "",                    strPathSQLMedia & strSPLevel))
      End If
      strCUSource   = GetMediaPath(GetParam(colStrings,   "CUSource",       "",                    strPathSQLMedia & "CU"))
      If strCUSource = "" Then
        strCUSource  = GetMediaPath(GetParam(colStrings,  "CUSource",       "",                    strPathSQLMedia & strSPCULevel))
      End If
    Case Else
      strPathSQLSP  = GetMediaPath(strPathSQLSPOrig)
      strPCUSource  = ""
      strCUSource   = ""
  End Select

  If strUpdateSource = "" Then
    strUpdateSource = strPathSQLSP
  End If

  strPathSSIS       = strDirProgSys & "\" & strSQLProgDir & "\" & strSQLVersionNum & "\DTS\Binn\Microsoft.SqlServer.IntegrationServices.Server.dll"

  Select Case True
    Case strProcArc = "X86"
      Select Case True
        Case strClusterAction = ""
          strPathSSMS    = strDirProg & "\" & strSQLVersionNum & "\Tools\"
        Case Else
          strPathSSMS    = strDirProgSys & "\" & strSQLProgDir & "\" & strSQLVersionNum & "\Tools\"
      End Select
      strDirProgX86 = strDirProg
      strDirProgSysX86 = strDirProgSys
      strPathSSMSx86 = strPathSSMS
      strPathBOL     = strPathSSMS
      Select Case True
        Case strSQLVersion = "SQL2005"
          strSQLMediaArc = strPathSQLMedia
        Case Else
          strSQLMediaArc = strPathSQLMedia & strFileArc & "\"
      End Select
    Case strProcArc = "AMD64"
      strDirProgSysX86 = objFSO.GetAbsolutePathName(objShell.ExpandEnvironmentStrings("%ProgramFiles(x86)%"))
      strDirProgX86 = strVolProg & Mid(strDirProgSysX86, 2) & "\" & strSQLProgDir
      Select Case True
        Case strClusterAction = ""
          strPathSSMS    = strDirProg & "\" & strSQLVersionNum & "\Tools\"
          strPathSSMSx86 = strDirProgX86 & "\" & strSQLVersionNum & "\Tools\"
        Case Else
          strPathSSMS    = strDirProgSys & "\" & strSQLProgDir & "\" & strSQLVersionNum & "\Tools\"
          strPathSSMSx86 = strDirProgSysX86 & "\" & strSQLVersionNum & "\Tools\"
      End Select
      Select Case True
        Case strSQLVersion <= "SQL2005"
          If strVolProg = strVolsys Then
            strPathSSMSx86 = strDirProgX86 & "\" & strSQLVersionNum & "\Tools\"
          Else 
            strPathSSMSx86 = strDirProg & " (x86)\" & strSQLVersionNum & "\Tools\"
          End If
          strPathBOL     = strPathSSMSx86
          strSQLMediaArc = strPathSQLMedia
        Case strSQLVersion = "SQL2008"
          strPathBOL     = strPathSSMSx86
          strSQLMediaArc = strPathSQLMedia & strFileArc & "\"
        Case strSQLVersion = "SQL2008R2"
          strPathBOL     = strPathSSMSx86
          strSQLMediaArc = strPathSQLMedia & strFileArc & "\"
        Case strSQLVersion >= "SQL2012"
          strPathBOL     = strPathSSMS
          strSQLMediaArc = strPathSQLMedia & strFileArc & "\"
      End Select
  End Select
  Call SetFolderPath("DirProgX86",        "VolProg",     "",           "",         strDirProgX86,                "")

  Select Case True
    Case strInstRegAS <> ""
      ' Nothing
    Case strType = "CLIENT"
      strInstRegAS = strSQLVersionNum & "\Tools"
    Case Else
      strPath       = Mid(strHKLMSQL, 6) & "Instance Names\OLAP\"
      objWMIReg.GetStringValue strHKLM, strPath, strInstASSQL, strInstRegAS
      If IsNull(strInstRegAS) Then
        strInstRegAS  = ""
      End If
  End Select

  Select Case True
    Case strInstRegSQL <> ""
      ' Nothing
    Case strType = "CLIENT"
      strInstRegSQL = strSQLVersionNum & "\Tools"
    Case Else
      strPath       = Mid(strHKLMSQL, 6) & "Instance Names\SQL\"
      objWMIReg.GetStringValue strHKLM, strPath, strInstance, strInstRegSQL
      If IsNull(strInstRegSQL) Then
        strInstRegSQL = ""
      End If
  End Select

  If Right(strDirDBA, 1) = "\" Then
    strDirDBA       = Left(strDirDBA, Len(strDirDBA) - 1)
  End If
  strPathVS         = strDirProgSysX86 & "\Microsoft Visual Studio " & strVSVersionPath & "\Common7\"
  strPSInstall      = """" & strDirSys & "\Microsoft.NET\Framework\v4.0.30319\InstallUtil.exe"""
  strRegPath        = GetRegPath(strHKLMSQL, strInstRegSQL)

  strDiscoverFile   = GetParam(Null,                  "DiscoverFile",       "",                    strDirServInst)
  strDiscoverFolder = GetParam(Null,                  "DiscoverFolder",     "",                    strPathFBStart)

  Call SetFolderPath("DirDBA",            "VolDBA",      "",           "",         strDirDBA,                    "")
  Call SetFolderPath("DirDRU",            "VolDRU",      "",           strDirSQL,  "",                           "")
  Call SetFolderPath("DirMDW",            "VolMDW",      "",           strDirDBA,  "MDW.Cache",                  "")

  Select Case True
    Case strSetupTempWin = "YES"
      Call SetFolderPath("PathTemp",      "VolTempWin",  "",           "",         strDirTempWin,                "")
      Call SetFolderPath("PathTempUser",  "VolTempWin",  "",           "",         strDirTempUser,               "")
    Case Else
      Call SetFolderPath("PathTemp",      "VolTempWin",  "",           "",         "",                           objShell.ExpandEnvironmentStrings(colSysEnvVars("TEMP")))
      Call SetFolderPath("PathTempUser",  "VolTempWin",  "",           "",         "",                           objShell.ExpandEnvironmentStrings(colUsrEnvVars("TEMP")))
  End Select

  If strSetupSQLDB = "YES" Then
    Call SetFolderPath("DirData",         "VolData",     "SqlAccount", strDirSQL,  strInstNode & ".Data",        GetRegValue(GetRegPath(strRegPath, "\MSSQLServer\DefaultData")))
    Call SetFolderPath("DirLog",          "VolLog",      "SqlAccount", strDirSQL,  strInstNode & ".Log",         GetRegValue(GetRegPath(strRegPath, "\MSSQLServer\DefaultLog")))
    Select Case True
      Case GetBuildfileValue("VolBackupSource") = "S"
        Call SetFolderPath("DirBackup",   "VolBackup",   "",           strDirSQL,  "",                           GetRegValue(GetRegPath(strRegPath, "\MSSQLServer\BackupDirectory")))
      Case Else
        Call SetFolderPath("DirBackup",   "VolBackup",   "",           strDirSQL,  strInstNode & ".Backup",      GetRegValue(GetRegPath(strRegPath, "\MSSQLServer\BackupDirectory")))
    End Select
    Select Case True
      Case strVolSysDB = strVolProg 
        Call SetFolderPath("DirSysDB",    "VolSysDB",    "SqlAccount", "",         strDirProg,                   GetRegValue(GetRegPath(strRegPath, "\Setup\SQLDataRoot")))
      Case Else
        Call SetFolderPath("DirSysDB",    "VolSysDB",    "SqlAccount", "",         strDirSQL,                    GetRegValue(GetRegPath(strRegPath, "\Setup\SQLDataRoot")))
    End Select
    Call SetFolderPath("DirTemp",         "VolTemp",     "SqlAccount", strDirSQL,  strInstNode & ".Data",        "")
    Call SetFolderPath("DirLogTemp",      "VolLogTemp",  "SqlAccount", strDirSQL,  strInstNode & ".Log",         "")
    If strSQLVersion >= "SQL2014" Then
      Call SetFolderPath("DirBPE",        "VolBPE",      "",           strDirSQL,  strInstNode & ".BPE",         "")
    End If
    Select Case True
      Case strSQLVersion < "SQL2008"
        ' Nothing
      Case strFSLevel = "0"
        ' Nothing
      Case Else
        Call SetFolderPath("DirDataFS",   "VolDataFS",   "",           strDirSQL,  strInstNode & ".Filestream",  "")
    End Select
    Call SetFolderPath("DirDataFT",       "VolDataFT",   "FtAccount",  strDirSQL,  strInstNode & ".FTData",      "")
    Call SetSystemDataBackup()
  End If

  If strSetupSQLAS = "YES" Then
    strRegPath      = GetRegPath(strHKLMSQL, strInstRegAS)
    Call SetFolderPath("DirDataAS",       "VolDataAS",   "AsAccount",  strDirSQL,  strInstNodeAS & ".Data",      GetRegValue(GetRegPath(strRegPath, "\Setup\DataDir")))
    Call SetFolderPath("DirLogAS",        "VolLogAS",    "AsAccount",  strDirSQL,  strInstNodeAS & ".Log",       "")
    Call SetFolderPath("DirTempAS",       "VolTempAS",   "AsAccount",  strDirSQL,  strInstNodeAS & ".Temp",      "")
    Select Case True
      Case GetBuildfileValue("VolBackupASSource") = "S"
        Call SetFolderPath("DirBackupAS", "VolBackupAS", "AsAccount",  strDirSQL,  "SSAS",                       "")
      Case Else
        Call SetFolderPath("DirBackupAS", "VolBackupAS", "AsAccount",  strDirSQL,  strInstNodeAS & ".Backup",    "")
    End Select
    Select Case True
      Case (strSQLVersion >= "SQL2012") And (strSPLevel < "RTM")
        strDirASDLL = strDirProg & "\" & strSQLVersionNum & "\Setup Bootstrap\" & strSQLVersion & strSPLevel & "\" & strFileArc
      Case strSQLVersion >= "SQL2017"
        strDirASDLL = strDirProg & "\" & strSQLVersionNum & "\Setup Bootstrap\" & strSQLVersion & "\" & strFileArc
      Case strSQLVersion >= "SQL2012"
        strDirASDLL = strDirProg & "\" & strSQLVersionNum & "\Setup Bootstrap\SQLServer" & Mid(strSQLVersion, 4) & "\" & strFileArc
      Case strSQLVersion >= "SQL2008"
        strDirASDLL = strDirProgX86 & "\" & strSQLVersionNum & "\Setup Bootstrap\Release\" & strFileArc
    End Select
  End If

  If strSetupSQLIS = "YES" Then
    strSSISVol      = "VolProg"
    Select Case True
      Case (strSetupSSISCluster = "YES")  And (strSetupSQLDBCluster = "YES")
        If Instr("NS", GetBuildfileValue("VolDataSource")) = 0 Then ' ie NOT Network or Share
          strSSISVol = "VolData"
        End If
      Case (strSetupSSISCluster = "YES")  And (strSetupSQLASCluster = "YES")
        If Instr("NS", GetBuildfileValue("VolDataASSource")) = 0 Then ' ie NOT Network or Share
          strSSISVol = "VolDataAS"
        End If
      Case (strSetupSSISCluster <> "YES") And (strSetupSQLDBCluster <> "YES") And (strSetupSQLDB = "YES")
        If Instr("NS", GetBuildfileValue("VolDataSource")) = 0 Then ' ie NOT Network or Share
          strSSISVol = "VolData"
        End If
      Case (strSetupSSISCluster <> "YES") And (strSetupSQLASCluster <> "YES") And (strSetupSQLAS = "YES")
        If Instr("NS", GetBuildfileValue("VolDataASSource")) = 0 Then ' ie NOT Network or Share
          strSSISVol = "VolDataAS"
        End If
    End Select
    Call SetFolderPath("DirDataIS",       strSSISVol,    "IsAccount",  strDirSQL,  strInstNodeIS & ".Data",      "")
  End If

  strRegPath        = GetRegPath(strHKLMSQL, strInstRegRS)
  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case strRegPath = ""
      ' Nothing
    Case strSetupPowerBI = "YES"
      Call SetFolderPath("PathSSRS",      "VolProg",     "RsAccount",  "",         strDirProgSys & "\Microsoft Power BI Report Server\" & strInstRSSQL, GetRegValue(GetRegPath(strRegPath, "\Setup\InstallRootDirectory")) & "\" & strInstRSSQL)
    Case strSQLVersion >= "SQL2017"
      Call SetFolderPath("PathSSRS",      "VolProg",     "RsAccount",  "",         strDirProg & "\" & strInstRegRS & "\Reporting Services", GetRegValue(GetRegPath(strRegPath, "\Setup\InstallRootDirectory")) & "\" & strInstRSSQL)
    Case Else
      Call SetFolderPath("PathSSRS",      "VolProg",     "RsAccount",  "",         strDirProg & "\" & strInstRegRS & "\Reporting Services", GetRegValue(GetRegPath(strRegPath, "\Setup\SQLPath")))
  End Select

  Select Case True
    Case strSetupMDS <> "YES"
      ' Nothing
    Case strSQLVersion = "SQL2008R2"
      Call SetFolderPath("PathMDS",       "VolProg",     "MDSAccount", "",         strDirProg & "\Master Data Services", "")
    Case Else
      Call SetFolderPath("PathMDS",       "VolProg",     "MDSAccount", "",         strDirProg & "\" & strSQLVersionNum & "\Master Data Services", "")
  End Select
  
End Sub


Function GetMediaPath(strMedia)
  Call DebugLog("GetMediaPath: " & strMedia)
  Dim strPath, strPathAlt

  strPath           = strMedia
  Select Case True
    Case Len(strPath) = 1
      strPath       = strPath & ":"
    Case Right(strPath, 1) = "\" 
      strPath       = Left(strPath, Len(strPath) - 1)
  End Select

  strPathAlt        = strPath
  Select Case True
    Case Left(strPath, 2) = "\\"
      ' Nothing
    Case (Left(strPath, 3) = "..\") And (Right(strPathFB, 1) = "\")
      strPath       = strPathFB & strPath
      Select Case True
        Case Instr(strPathAlt, "_") > 0
          strPathAlt = strPathFB & "..\" & Left(strPathAlt, Instr(strPathAlt, "_") - 1) & Mid(strPathAlt, 3)
        Case Else
          strPathAlt = strPath
      End Select
    Case Left(strPath, 3) = "..\"
      strPath       = strPathFB & "\" & strPath
      Select Case True
        Case Instr(strPathAlt, "_") > 0
          strPathAlt = strPathFB & "\..\" & Left(strPathAlt, Instr(strPathAlt, "_") - 1) & Mid(strPathAlt, 3)
        Case Else
          strPathAlt = strPath
      End Select
    Case (Mid(strPath, 2, 1) = ":") And (GetBuildfileValue("Vol" & Left(strPath, 1) & "Source") = "N")
      strPath       = GetBuildfileValue("Vol" & Left(strPath, 1) & "Path") & Mid(strPath, 3)
      strPathAlt    = strPath
  End Select

  strPath           = CheckMediaPath(strPath)
  If strPath = "" Then
    strPath         = CheckMediaPath(strPathAlt)
  End If

  Call DebugLog(" Path found: " & strPath)
  GetMediaPath      = strPath

End Function


Function CheckMediaPath(strPath)
  Call DebugLog("CheckMediaPath:" & strPath)
  Dim strPathWork

  strPathWork       = strPath
  Select Case True
    Case objFSO.FolderExists(strPathWork)
      Set objFolder = objFSO.GetFolder(strPathWork)
      strPathWork   = objFolder.Path
      Select Case True
        Case Left(strPathWork, 2) = "\\"
          ' Nothing
        Case GetBuildfileValue("Vol" & Left(strPathWork, 1) & "Source") = "N" 
          strPathWork = GetBuildfileValue("Vol" & Left(strPathWork, 1) & "Path") & Mid(strPathWork, 3)
      End Select
    Case Else
      strPathWork   = ""
  End Select

  Select Case True
    Case strPathWork = ""
      ' Nothing
    Case Right(strPathWork, 1) <> "\" 
      strPathWork   = strPathWork & "\"
  End Select

  CheckMediaPath    = strPathWork

End Function


Function GetRegPath(strRegBase, strRegItem)
  Call DebugLog("GetRegPath: " & strRegItem)
  Dim strRegPath

  Select Case True
    Case (strRegBase = "") Or (IsNull(strRegBase))
      strRegPath    = ""
    Case (strRegItem = "") Or (IsNull(strRegItem))
      strRegPath    = ""
    Case Else
      strRegPath    = strRegBase & strRegItem
  End Select

  Call DebugLog("Reg path: " & strRegPath)
  GetRegPath       = strRegPath

End Function


Function GetRegValue(strRegPath)
  Call DebugLog("GetRegValue: " & strRegPath)
  Dim strRegBase, strRegItem, strRegValue

  Select Case True
    Case strRegPath = ""
      strRegValue   = ""
    Case Else
      strRegBase    = Left(strRegPath, InStrRev(strRegPath, "\") - 1)
      strRegItem    = Mid(strRegPath, InStrRev(strRegPath, "\") + 1)
      If Left(strRegBase, 5) = "HKLM\" THen
        strRegBase  = Mid(strRegBase, 6)
      End If
      objWMIReg.GetStringValue strHKLM,strRegBase,strRegItem,strRegValue
      If IsNull(strRegValue) Then
        strRegValue = ""
      End If
  End Select

  Call DebugLog("Reg value: " & strRegValue)
  GetRegValue       = strRegValue

End Function


Sub SetFolderPath(strVarName, strVolVar, strAccountVar, strRoot, strPath, strAltPath)
  Call DebugLog("SetFolderPath: " & strVarName & " for " & strPath)
  Dim arrVolumes
  Dim strAccountParm, strAccountType, strPathBase, strPathWork, strPathDir, strRootWork, strVolume, strVolSource, strVolWork

  strAccountParm    = "/" & UCase(GetBuildfileValue(strAccountVar & "Name")) & ":"
  strAccountType    = GetBuildfileValue(strAccountVar & "Type")
  strVolume         = GetBuildfileValue(strVolVar)
  strVolSource      = GetBuildfileValue(strVolVar & "Source")

  Select Case True
    Case strVolume = ""
      Exit Sub ' Supplied volume cannot be found, see previous error message
    Case strAltPath = ""
      strPathWork   = strPath
      strRootWork   = "\" & strRoot
      strVolWork    = GetBuildfileValue(strVolVar)
    Case CheckCSV(strVarName, strAltPath) <> 0
      strPathWork   = Mid(strAltPath, Len(strCSVRoot) + 1)
      strVolWork    = strCSVRoot & Left(strPathWork, Instr(strPathWork, "\") - 1)
      strPathWork   = Mid(strPathWork, Instr(strPathWork, "\") + 1)
      strRootWork   = "\" & Left(strPathWork, Instr(strPathWork, "\"))
      strPathWork   = Mid(strPathWork, Instr(strPathWork, "\") + 1)
    Case CheckShare(strVarName, strAltPath) <> 0
      strPathWork   = Mid(strAltPath, Len(strRemoteRoot) + 2)
      strVolWork    = strRemoteRoot
      strRootWork   = ""
    Case Else
      strPathWork   = strAltPath
      strVolWork    = ""
      strRootWork   = ""
      If Mid(strPathWork, 2, 2) = ":\" Then
        strVolWork  = Left(strPathWork, 1)
        strRootWork = "\"
      End If
  End Select
  Select Case True
    Case strRootWork = ""
      ' Nothing
    Case Right(strRootWork, 1) = "\"
      ' Nothing
    Case Else
      strRootWork   = strRootWork & "\"
  End Select
  If Mid(strPathWork, 2, 2) = ":\" Then
    strPathWork     = Mid(strPathWork, 4)
  End If

  arrVolumes        = Split(strVolWork, ",")
  Select Case True
    Case strVolSource = "D"
      strPathBase   = ":" & strRootWork & strPathWork
      strPathDir    = Left(strVolWork, 1) & strPathBase
    Case strVolSource = "C"
      strPathBase   = strRootWork & strPathWork
      strPathDir    = Trim(arrVolumes(0)) & strPathBase
      Select Case True
        Case Instr(UCase(" VolDataAS VolLogAS VolTempAS "), UCase(" " & strVolVar & " ")) > 0
          Call SetBuildMessage(strMsgErrorConfig, "Analysis Services can not be installed to a Cluster Shared Volume (CSV): " & strVolVar)
        Case strSQLVersion <= "SQL2008"
          Call SetBuildMessage(strMsgErrorConfig, "Cluster Shared Volume cannot be used for " & strSQLVersion)
        Case strSetupSQLDBCluster = "YES"
          ' Nothing
        Case (strSetupAlwaysOn = "YES") And (Instr(UCase(" VolSysDB VolTemp VolLogTemp "), UCase(" " & strVolVar & " ")) > 0)
          Call SetBuildMessage(strMsgErrorConfig, "Cluster Shared Volume cannot be used for /" & strVolVar & ": with Always On")
      End Select
    Case (strVolSource = "M") And (strVolVar = strVolRootAS)
      strPathBase   = Mid(strMountRoot, 2) & strPathWork
      strPathDir    = strRootAS & strPathBase
    Case strVolSource = "M" 
      strPathBase   = strRootWork & strPathWork
      strPathDir    = strMountRoot & strPathBase
    Case strVolSource = "N" 
      strPathBase   = strRootWork & strPathWork
      strPathDir    = strMountRoot & strPathBase
      Select Case True
        Case Instr(UCase(" VolDataAS VolLogAS VolTempAS "), UCase(" " & strVolVar & " ")) > 0
          Call SetBuildMessage(strMsgErrorConfig, "Analysis Services can not be installed to a Network Drive: " & strVolVar)
        Case strSQLVersion > "SQL2008R2"
          ' Nothing
        Case (Instr(UCase(" VolData VolLog VolTemp VolTempLog "), UCase(" " & strVolVar & " ")) > 0)
          Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " can not be installed to a Network Drive: " & strVolVar)
        Case strAccountVar = "FtAccount"
          ' Nothing
        Case (strAccountType = "L") Or (strAccountType = "S")
          Call SetBuildMessage(strMsgErrorConfig, "/" & strAccountParm & ": parameter must be a domain account")
      End Select
    Case strVolSource = "S" 
      strPathBase   = "\" & strPathWork
      strVolWork    = Trim(arrVolumes(0))
      If Len(strVolWork) = 1 Then
        strVolWork  = strVolWork & ":"
      End If
      strPathDir    = strVolWork & strPathBase
      Select Case True
        Case Instr(UCase(" VolDataAS VolLogAS VolTempAS "), UCase(" " & strVolVar & " ")) > 0 
          Call SetBuildMessage(strMsgErrorConfig, "Analysis Services can not be installed to a File Share: " & strVolVar)
        Case strSQLVersion > "SQL2008R2"
          ' Nothing
        Case (Instr(UCase(" VolData VolLog VolTemp VolTempLog "), UCase(" " & strVolVar & " ")) > 0) 
          Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " can not be installed to a File Share " & strVolVar)
        Case strAccountVar = "FtAccount"
          ' Nothing
        Case (strAccountType = "L") Or (strAccountType = "S")
          Call SetBuildMessage(strMsgErrorConfig, "/" & strAccountParm & ": parameter must be a domain account")
      End Select
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, strVarName & " " & GetBuildfileValue(strVolVar) & " has unknown Volume Source code: " & strVolSource)
  End Select

  If Right(strPathBase, 1) = "\" Then
    strPathBase     = Left(strPathBase, Len(strPathBase) - 1)
    strPathDir      = Left(strPathDir, Len(strPathDir) - 1)
  End If

  Call SetBuildfileValue(strVarName, strPathDir)
  Call SetBuildfileValue(strVarName & "Base", strPathBase)

End Sub


Sub SetSystemDataBackup()
  Call DebugLog("SetSystemDataBackup:")
  Dim strAGDagServer, strDirBackup

  strPath           = GetBuildfileValue("DirBackup")

  If Instr(strPath, "AdHocBackup") > 0 Then
    strPath         = Left(strPath, Instr(strPath, "AdHocBackup") - 1)
  End If
  If Right(strPath, 1) <> "\" Then
    strPath         = strPath & "\"
  End If

  Select Case True
    Case strSetupSQLDBCluster = "YES"
      strDirBackup  = strClusterNameSQL
    Case Else
      strDirBackup  = strDirServInst
  End Select

  Call SetBuildfileValue("DirSystemDataBackup", strPath & "SystemDataBackup\" & strDirBackup)

  Select Case True
    Case strSetupAlwaysOn = "YES"
      strDirBackup  = strGroupAO
  End Select

  Call SetBuildfileValue("DirSystemDataShared", strPath & "SystemDataBackup\" & strDirBackup)

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case strAGDagName = ""
      ' Nothing
    Case GetStatefileValue(strAGDagName) = ""
      ' Nothing
    Case Else
      strDirBackup  = GetStatefileValue(strAGDagName)
  End Select

  Call SetBuildfileValue("DirSystemDataPrimary",  strPath & "SystemDataBackup\" & strDirBackup)

End Sub


Sub GetDTCVolumes()
  Call DebugLog("GetDTCVolumes:")
  Dim strGroup

  strGroup          = GetBuildfileValue("ClusterGroupDTC")
  strVolDTC         = GetVolumes("VolDTC",       "",  "Y", strVolSys,  1, strActionDTC)

  Select Case True
    Case strGroup = strClusterGroupSQL
      strClusterGroupDTC   = strClusterGroupSQL
      strClusterNetworkDTC = strClusterNetworkSQL
      strLabDTC            = strLabData
    Case strGroup = strClusterGroupAS
      strClusterGroupDTC   = strClusterGroupAS
      strClusterNetworkDTC = strClusterNetworkAS
      strLabDTC            = strLabDataAS
    Case strVolDTC = GetVolumes("VolData",      "",  "Y", strVolData,   1, strActionSQLDB) 
      strClusterGroupDTC   = strClusterGroupSQL
      strClusterNetworkDTC = strClusterNetworkSQL
      strLabDTC            = strLabData
    Case strVolDTC = GetVolumes("VolDataAS",    "",  "Y", strVolDataAS, 1, strActionSQLAS) 
      strClusterGroupDTC   = strClusterGroupAS
      strClusterNetworkDTC = strClusterNetworkAS
      strLabDTC            = strLabDataAS
    Case Else
      strClusterGroupDTC   = strClusterNameDTC
      strClusterNetworkDTC = "DTC " & strClusterNameDTC
  End Select

End Sub


Sub SetupInstRS()
  Call DebugLog("SetupInstRS:")
  Dim strCapMemory, strItemReg, strWorkRSMode

  Select Case True
    Case strtype = "CLIENT"
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListType)
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListSQLVersion)
    Case strOSVersion < "6.2"
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListOSVersion)
    Case strFileArc = "X86"
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListOSVersion)
  End Select

  Select Case True
    Case strSetupSQLRSCluster = "YES"
      strSetupSQLRS = "YES"
    Case strSetupPowerBI = "YES"
      strSetupSQLRS = "YES"
    Case strSetupSQLRS = ""
      strSetupSQLRS = "YES"
  End Select

  Select Case True
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "N/A", "", strListSQLRS)
    Case (strSQLVersion < "SQL2017") And (strSetupPowerBI <> "")
      ' Nothing
    Case strSQLVersion < "SQL2017"
      strSetupPowerBI = "NO"
    Case Instr("DATA CENTER ENTERPRISE EVALUATION DEVELOPER", strEdition) = 0
      strSetupPowerBI = "NO"
    Case Else
      Call SetParam("SetupPowerBI",          strSetupPowerBI,          "YES", "PowerBI recommended for " & strSQLVersion, "")
  End Select

  Select Case True
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "N/A", "", strListCore)
    Case strSetupSQLRS <> ""
      ' Nothing
    Case strMainInstance = "YES"
      strSetupSQLRS = "YES"
    Case strSetupSQLRSCluster = "YES"
      strSetupSQLRS = "YES"
    Case Else
      Call SetParam("SetupSQLRS",            strSetupSQLRS,            "NO",  "SSRS not installed by default with secondary SQL Instances on server", "")
  End Select

  Select Case True
    Case strClusterHost <> "YES"
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListCluster)
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListSQLRS)
    Case (strClusterAction <> "") And (strSetupSQLRSCluster = "")
      strSetupSQLRSCluster = "YES"
    Case strSetupSQLRSCluster = ""
      strSetupSQLRSCluster = "NO"
  End Select

  Select Case True
    Case strInstance = "MSSQLSERVER"
      strRSDBName   = "ReportServer"
    Case Else
      strRSDBName   = "ReportServer" & "$" & strInstance
  End Select
  strSetupRSDB      = CheckSetupRSDB()

  strRSVersion      = Left(strSQLVersionNet, Instr(strSQLVersionNet, ".") - 1)

  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupPowerBI = "YES"
      Call SetFileData("SQLRSexe",     "PowerBIexe",                        "", "")
      strRSVersionNum = CheckRSVersion()
      strInstRSDir  = "\Setup\InstallRootDirectory"
      strInstRSSQL  = "PBIRS"
      strInstRS     = "PowerBIReportServer"
      strInstRSURL  = "ReportServer"
      strRSFullURL  = strHTTP & "://" & strServer & "/Reports" & "/Pages/Folder.aspx"
      strRSURLSuffix = ""
    Case strSQLVersion >= "SQL2017"
      Call SetFileData("SQLRSexe",     "",                                  "", "SQLServerReportingServices.exe")
      strInstRSDir  = "\Setup\InstallRootDirectory"
      strInstRSSQL  = "SSRS"
      strInstRS     = "SQLServerReportingServices"
      strInstRSURL  = "ReportServer"
      strRSFullURL  = strHTTP & "://" & strServer & "/Reports" & "/Pages/Folder.aspx"
      strRSURLSuffix = ""
    Case (strSetupSQLRSCluster = "YES") And (strSQLVersion <= "SQL2005")
      strInstRSDir  = "\Setup\SQLPath"
      strInstRSSQL  = strClusterBase & strClusRSSuffix
      strInstRS     = "ReportServer$" & strInstRSSQL
      strInstRSURL  = "ReportServer"
      strRSFullURL  = strHTTP & "://" & strClusterNameRS & "/Reports/Pages/Folder.aspx"
      strRSURLSuffix = ""
    Case strSetupSQLRSCluster = "YES"
      strInstRSDir  = "\Setup\SQLPath"
      strInstRSSQL  = strClusterBase & strClusRSSuffix
      strInstRS     = "ReportServer$" & strInstRSSQL
      strInstRSURL  = "ReportServer_" & strInstRSSQL
      strRSFullURL  = strHTTP & "://" & strClusterNameRS & "/Reports_" & strInstRSSQL & "/Pages/Folder.aspx"
      strRSURLSuffix = ""
    Case strInstance = "MSSQLSERVER"
      strInstRSDir  = "\Setup\SQLPath"
      strInstRSSQL  = strInstance
      strInstRS     = "ReportServer"
      strInstRSURL  = "ReportServer"
      strRSFullURL  = strHTTP & "://" & strServer & "/Reports" & "/Pages/Folder.aspx"
      strRSURLSuffix = ""
    Case Else
      strInstRSDir  = "\Setup\SQLPath"
      strInstRSSQL  = strInstance
      strInstRS     = "ReportServer$" & strInstRSSQL
      strInstRSURL  = "ReportServer_" & strInstRSSQL
      strRSFullURL  = strHTTP & "://" & strServer & "/Reports_" & strInstRSSQL & "/Pages/Folder.aspx"
      strRSURLSuffix = "_" & strInstRSSQL
  End Select

  strActionSQLRS    = GetItemAction(strType, strAction, "SQLRS", strSetupSQLRSCluster)
  strInstRSWMI      = "RS_" & Replace(Replace(Replace(strInstRSSQL, "_", "_5f"), "$", "_24"), "@", "_40")

  Select Case True
    Case strInstRegRS <> ""
      ' Nothing
    Case strType = "CLIENT"
      strInstRegRS = strSQLVersionNum & "\Tools"
    Case Else
      strPath       = Mid(strHKLMSQL, 6) & "Instance Names\RS\"
      objWMIReg.GetStringValue strHKLM, strPath, strInstance, strInstRegRS
      If IsNull(strInstRegRS) Then
        strInstRegRS  = ""
      End If
  End Select

  Select Case True
    Case strEdition = "EXPRESS" 
      strInstRSHost = strServer & ":80"
    Case Else
      strInstRSHost = strServer
  End Select

  Select Case True
    Case strInstRegRS = ""
      ' Nothing
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

  Select Case True
    Case strSQLVersion <= "SQL2005"
      strWorkRSMode        = GetParam(colGlobal,      "RSConfiguration",    "",                    "Default")
    Case strSQLVersion < "SQL2012"
      strWorkRSMode        = GetParam(colGlobal,      "RSInstallMode",      "",                    "DefaultNativeMode")
    Case strSQLVersion < "SQL2017"
      strWorkRSMode        = GetParam(colGlobal,      "RSInstallMode",      "",                    "DefaultNativeMode")
      strRSShpInstallMode  = GetParam(colGlobal,      "RSShpInstallMode",   "",                    "DefaultSharePointMode")
    Case Else
      strWorkRSMode        = GetParam(colGlobal,      "RSInstallMode",      "",                    "DefaultNativeMode")
  End Select
  strRSInstallMode   = GetBuildfileValue("RSInstallMode")
  If strRSInstallMode = "" Then
    strRSInstallMode = strWorkRSMode
  End If

  strCapMemory      = 6000
  Select Case True
    Case Not IsNumeric("0" & strSetWorkingSetMaximum)
      Call SetBuildMessage(strMsgErrorConfig, "/SetWorkingSetMaximum: value must be Numeric")
    Case strSetWorkingSetMaximum <> ""
      ' Nothing
    Case strServerMB <= CLng((strCapMemory / 80) * 100)
      strSetWorkingSetMaximum = "0"
    Case Else
      strSetWorkingSetMaximum = Int(strCapMemory * 1000)
      Call ParamListAdd("ListNonDefault", "SetWorkingSetMaximum")
  End Select

End Sub


Function CheckSetupRSDB()
  Call DebugLog("CheckSetupRSDB:")
  Dim strSetup

  Select Case True
    Case strActionDAG = "ADDNODE"
      strSetup      = "NO"
    Case strActionSQLRS = strActionClusInst
      strSetup      = "YES"
    Case strSQLVersion >= "SQL2017"
      strSetup      = "YES"
    Case strSetupPowerBI = "YES"
      strSetup      = "YES"
    Case Else
      strSetup      = "NO"
  End Select

  Select Case True
    Case (strCatalogServerName = strClusterNameSQL) And (strClusterNameSQL <> "") And (strCatalogInstance = strInstance)
      ' Nothing
    Case (strCatalogServerName = strGroupAO) And (strGroupAO <> "")
      ' Nothing
    Case (strCatalogServerName = strServer) And (strCatalogInstance = strInstance)
      ' Nothing
    Case strCatalogServerName = strRSAlias
      ' Nothing
    Case Else
      strSetup      = "NO"
  End Select

  CheckSetupRSDB    = strSetup

End Function


Function CheckRSVersion()
  Call DebugLog("CheckRSVersion:")
  Dim strRSVersion

  strRSVersion      = strRSVersionNum
  strPath           = FormatFolder(strPathAddComp) & "\" & GetBuildfileValue("SQLRSexe")
  Select Case True
    Case strSetupPowerBI <> "YES"
      ' Nothing
    Case Not objFSO.FileExists(strPath)
      ' Nothing
    Case objFSO.GetFileVersion(strPath) > "1.2"
      strRSVersion  = Max(strRSVersion, "15")
    Case Else
      strRSVersion  = Max(strRSVersion, "14")
  End Select

  CheckRSVersion    = strRSVersion

  strRSVersion      = Left(strSQLVersionNet, Instr(strSQLVersionNet, ".") - 1)

End Function


Sub GetFileData()
  Call SetProcessId("0BL", "Get data for Install files")

  Select Case True
    Case strWOWX86 = "TRUE"
      strSPFile     = strSPLevel & "X86"
      strCUFile     = strSPLevel & "X86" & strSPCULevel
    Case Else
      strSPFile     = strSPLevel & strFileArc
      strCUFile     = strSPLevel & strFileArc & strSPCULevel
  End Select
  If (strSQLVersion = "SQL2005") And (strEdition = "EXPRESS") Then
    strSPFile       = strSPFile & "Exp"
  End If

  Call SetFileData("ABEmsi",           "ABE" & strFileArc & "msi",          "", "%")
  Call SetFileData("AccidentalDBAzip", "AccidentalDBAzip",                  "", "Accidental_DBA_EBook.zip")
  Call SetFileData("BIDSexe",          "BIDSexe",                           "", "%")
  Call SetFileData("BOLexe",           "BOLexe",                            "", "%")
  Call SetFileData("BOLmsi",           "BOLmsi",                            "", "%")
  Call SetFileData("BPAmsi",           "BPA" & strFileArc & "msi",          "", "%")
  Call SetFileData("CacheManagerZip",  "CacheManagerZip",                   "", "%")
  Call SetFileData("CUFile",           strCUFile,                           "", "%")
  Call SetFileData("DB2exe",           "DB2exe",                            "", "%")
  Call SetFileData("DB2OLEmsi",        "DB2OLE" & strFileArc & "msi",       "", "%")
  Call SetFileData("DBAManagementbat", "DBAManagementbat",                  "", "Install.bat")
  Call SetFileData("DBAManagementCab", "DBAManagementCab",                  "", "SqlDBAManagement.cab")
  Call SetFileData("DimensionSCDZip",  "DimensionSCDZip",                   "", "%")
  Call SetFileData("DimensionSCDmsi",  "DimensionSCD" & strFileArc & "msi", "", "%")
  Call SetFileData("DTSBackupmsi",     "DTSBackupmsi",                      "", "%")
  Call SetFileData("DTSmsi",           "DTSmsi",                            "", "%")
  Call SetFileData("DTSFix",           "DTSFix",                            "", "%")
  Call SetFileData("GenMaintCab",      "GenMaintCab",                       "", "GenericMaintenance.cab")
  Call SetFileData("GenMaintSql",      "GenMaintSql",                       "", "MaintenanceSolution.sql")
  Call SetFileData("GenMaintVbs",      "GenMaintVbs",                       "", "INSTALL.VBS")
  Call SetFileData("GovernorSql",      "GovernorSql",                       "", "Set-ResourceGovernor.sql")
  Call SetFileData("IntViewermsi",     "IntViewermsi",                      "", "%")
  Call SetFileData("Javaexe",          "Javaexe",                           "", "%")
  Select Case True
    Case strSQLVersion >= "SQL2019"
      strPath       = FormatFolder(strPathSQLMedia) & "x64\Setup"
      Call SetFileData("JREexe",        "",                                 "", strPath & "sql_azul_*.msi")
    Case Else
      Call SetFileData("JREexe",       "JRE" & strFileArc & "exe",          "", "%")
  End Select
  Call SetFileData("KB925336exe",      "KB925336" & strFileArc & "exe",     "", "%")
  Call SetFileData("KB932232exe",      "KB932232exe",                       "", "%")
  Call SetFileData("KB933789exe",      "KB933789" & strFileArc & "exe",     "", "%")
  Call SetFileData("KB937444exe",      "KB937444" & strFileArc & "exe",     "", "%")
  Call SetFileData("KB954961exe",      "KB954961exe",                       "", "%")
  Call SetFileData("KB956250msu",      "KB956250" & strFileArc & "msu",     "", "%")
  Call SetFileData("KB2549864exe",     "KB2549864exe",                      "", "%")
  Call SetFileData("KB2781514exe",     "KB2781514exe",                      "", "%")
  Call SetFileData("KB2919355msu",     "KB2919355" & strFileArc & "msu",    "", "%")
  Call SetFileData("KB2919442msu",     "KB2919442" & strFileArc & "msu",    "", "%")
  Call SetFileData("KB3090973msu",     "KB3090973" & strFileArc & "msu",    "", "%")
  Call SetFileData("MBCAmsi",          "MBCA" & strFileArc & "msi",         "", "%")
  If strOSVersion >= "6.2" Then
    Call SetFileData("MBCAmsi",        "MBCAWin8" & strFileArc & "msi",     "", "%")
  End If
  Call SetFileData("MDXexe",           "MDXexe",                            "", "MDXStudio.exe")
  Select Case True
    Case strSQLVersion <= "SQL2008R2"
      Call SetFileData("MDXZip",       "MDXZip",                            "MDXStudio_0_4_15.zip", "%")
    Case Else
      Call SetFileData("MDXZip",       "MDXZip",                            "", "%")
  End Select
  Select Case True
    Case strOSVersion <= "6" 
      ' Nothing
    Case strOSVersion = "6.0" 
      Call SetFileData("Net4Xexe",     "Net4Xexe",                          "NDP461-KB3151800-x86-x64-AllOS-ENU.exe", "")
    Case (strOSVersion = "6.2" ) And (strOSType = "CLIENT")
      Call SetFileData("Net4Xexe",     "Net4Xexe",                          "NDP461-KB3151800-x86-x64-AllOS-ENU.exe", "")
    Case (strSQLVersion = "SQL2016") And (Instr(strOSType, "CORE") > 0)
      Call SetFileData("Net4Xexe",     "Net4Xexe",                          "NDP462-KB3151800-x86-x64-AllOS-ENU.exe", "")
    Case Else
      Call SetFileData("Net4Xexe",     "Net4Xexe",                          "", "%")
  End Select
  Call SetFileData("PBMBat",           "PBMBat",                            "", "FineBuildPBM.BAT")
  Call SetFileData("PBMCab",           "PBMCab",                            "", "FineBuildPBM.cab")
  Call SetFileData("PDFexe",           "PDFexe",                            "", "%")
  Call SetFileData("PDFreg",           "PDFreg",                            "", "SumatraPDF")
  Call SetFileData("PerfDashmsi",      "PerfDashmsi",                       "", "%")
  Call SetFileData("PlanExpexe",       "PlanExp" & strFileArc & "exe",      "", "%")
  Call SetFileData("PlanExpAddinmsi",  "PlanExpAddinmsi",                   "", "SQLSentryPlanExplorerSSMSAddinSetup.msi")
  Call SetFileData("ProcExpDir",       "ProcExpDir",                        "", "ProcessExplorer")
  Call SetFileData("ProcExpexe",       "ProcExpexe",                        "", "Procexp.exe")
  Call SetFileData("ProcExpZip",       "ProcExpZip",                        "", "ProcessExplorer.zip")
  Call SetFileData("ProcMonDir",       "ProcMonDir",                        "", "ProcessMonitor")
  Call SetFileData("ProcMonexe",       "ProcMonexe",                        "", "Procmon.exe")
  Call SetFileData("ProcMonZip",       "ProcMonZip",                        "", "ProcessMonitor.zip")
  Select Case True
    Case (Instr(Ucase(strOSName), " XP") > 0) And (strProcArc  = "X86")    ' Windows XP and 32-bit
      Call SetFileData("PS1File",      "PS1XPX86",                          "", "%")
      Call SetFileData("PS2File",      "PS2XPX86",                          "", "%")
    Case (strOSVersion < "6.0") And (strProcArc  = "X86")                  ' Windows 2003 32-bit
      Call SetFileData("PS1File",      "PS1W2003X86",                       "", "%")
      Call SetFileData("PS2File",      "PS2W2003" & strFileArc,             "", "%")
    Case strOSVersion < "6.0"                                              ' Windows 2003 or XP 64-bit
      Call SetFileData("PS1File",      "PS1XPW2003X64",                     "", "%")
      Call SetFileData("PS2File",      "PS2W2003" & strFileArc,             "", "%")
    Case (strOSVersion = "6.0") And (Instr(Ucase(strOSName), "VISTA") > 0) ' Windows Vista
      Call SetFileData("PS1File",      "PS1Vista" & strFileArc,             "", "%")
      Call SetFileData("PS2File",      "PS2W2008" & strFileArc,             "", "%")
      Call SetFileData("KB2862966File","KB2862966W2008" & strFileArc & "msu", "", "%")
    Case strOSVersion = "6.0"                                              ' Windows 2008
      Call SetFileData("PS1File",      "",                                  "", "PKGMGR")
      Call SetFileData("PS2File",      "PS2W2008" & strFileArc,             "", "%")
      Call SetFileData("KB2862966File","KB2862966W2008" & strFileArc & "msu", "", "%")
    Case strOSVersion = "6.1"  
      Call SetFileData("PS1File",      "",                                  "", "PKGMGR")
      Call SetFileData("PS2File",      "",                                  "", "PKGMGR")
      Call SetFileData("KB2862966File","KB2862966W2008R2" & strFileArc & "msu", "", "%")
    Case Else
      Call SetFileData("PS1File",      "",                                  "", "PKGMGR")
      Call SetFileData("PS2File",      "",                                  "", "PKGMGR")
      Call SetFileData("KB2862966File","KB2862966" & strFileArc & "msu",    "", "%")
  End Select
  Call SetFileData("RawReaderexe",     "RawReaderexe",                      "", "RawFileReader.exe")
  Call SetFileData("ReportViewerexe",  "ReportViewerexe",                   "", "ReportViewer.exe")
  Call SetFileData("RMLToolsmsi",      "RMLTools" & strFileArc & "msi",     "", "%")
  Call SetFileData("RptTaskPadRdl",    "RptTaskPadRdl",                     "", "SQL2000 Taskpad View.rdl")
  Call SetFileData("RSKeepAliveCab",   "RSKeepAliveCab",                    "", "RSKeepAlive.cab")
  Call SetFileData("RSScripterZip",    "RSScripterZip",                     "", "RSScripter.zip")
  Call SetFileData("RSLinkGenZip",     "RSLinkGenZip",                      "", "RSLinkgen.zip")
  Call SetFileData("Samplesmsi",       "Samples" & strFileArc & "msi",      "", "%")
  strSNACFile       = FormatFolder(GetMediaPath(strPathSQLSPOrig & "\" & strSPLevel))
  Call SetFileData("SNACFile",         strSNACFile,                         "", "SQLNCLI*" & strFileArc & ".msi")
  Call SetFileData("SPFile",           strSPFile,                           "", "")
  Call SetFileData("SQLBCmsi",         "SQLBC" & strFileArc & "msi",        "", "%")
  Call SetFileData("SQLCEexe",         "SQLCE" & strFileArc & "exe",        "", "%")
  Call SetFileData("SQLNexuszip",      "SQLNexuszip",                       "", "%")
  Call SetFileData("SQLNSmsi",         "SQLNS" & strFileArc & "msi",        "", "%")
  Call SetFileData("SSDTBIexe",        "SSDTBIexe",                         "", "%")
  Call SetFileData("SSMSexe",          "SSMSexe",                           "", "%")
  Call SetFileData("SysManagementbat", "SysManagementbat",                  "", "Install.bat")
  Call SetFileData("SysManagementCab", "SysManagementCab",                  "", "SqlSysManagement.cab")
  Call SetFileData("SystemViewsPDF",   "SystemViewsPDF",                    "", "%")
  Call SetFileData("TroublePDF",       "TroublePDF",                        "", "%")
  Call SetFileData("VS2005SP1exe",     "VS2005SP1exe",                      "", "SETUP.EXE")
  Call SetFileData("VS2010SP1exe",     "VS2010SP1exe",                      "", "SETUP.EXE")
  Call SetFileData("XEventsmsi",       "XEventsmsi",                        "", "%")
  Call SetFileData("XMLmsi",           "XMLmsi",                            "", "XmlNotepad.msi")
  Call SetFileData("ZoomItDir",        "ZoomItDir",                         "", "ZoomIt")
  Call SetFileData("ZoomItExe",        "ZoomItExe",                         "", "ZoomIt.exe")
  Call SetFileData("ZoomItZip",        "ZoomItZip",                         "", "ZoomIt.zip")

  strSSMSexe        = GetBuildfileValue("SSMSexe")

End Sub


Sub SetFileData(strBuild, strParam, strMaxFile, strDefault)
  Call DebugLog("SetFileData: " & strBuild)
  Dim colREFiles
  Dim objFolder
  Dim strFolder, strREPath, strStorePath

  Select Case True
    Case strDefault = ""
      strDefault    = strUnknown
  End Select

  Select Case True
    Case strParam = ""
      strFolder     = ""
      strPath       = strDefault
    Case Mid(strParam, 2, 1) = ":"
      strFolder     = strParam
      strPath       = strDefault
    Case Else
      strFolder     = strPathAddComp
      strPath       = GetParam(colFiles,              strParam,            "",                     strDefault)
  End Select

  Select Case True
    Case strFolder = ""
      ' Nothing
    Case Not objFSO.FolderExists(strFolder)
      strPath       = ""
    Case Instr(strPath, "*") > 0 
      strREPath     = "^" & Replace(strPath, "*", ".*") & "$"
      strPath       = ""
      objRE.Pattern = strREPath
      Set objFolder = objFSO.GetFolder(strFolder)
      Set colREFiles = objFolder.Files
      For Each objFile In colREFiles
        Select Case True
          Case Not objRE.Test(objFile.Name)
            ' Nothing
          Case (UCase(objFile.Name) > UCase(strMaxFile)) And (strMaxFile <> "")
            ' Nothing
          Case strPath < objFile.name
            strPath = objFile.name
        End Select
      Next
  End Select

  If strPath = "" Then
    strPath         = strUnknown
  End If
  Call SetBuildfileValue(strBuild, strPath)

  Set colREFiles    = Nothing
  Set objFolder     = Nothing

End Sub


Sub GetSetupData()
  Call SetProcessId("0BM", "Get Setup Parameter data")

  Select Case True
    Case Instr(" CONFIG DISCOVER FIX REBUILD ", strType) > 0
      ' Nothing
    Case strSetupCompliance = "YES"
      Call SetupDataCompliance
  End Select

  Call SetupDataDep0

  Call SetupDataDep1

  Call SetupDataDep2

  Call SetupDataDep3

End Sub


Sub SetupDataCompliance()
  Call SetProcessId("0BMA", "Setup Parameter Data for Compliance")

  If strOSVersion < "6.3A" Then
    Call SetParam("SetupABE",                strSetupABE,              "YES", "", strListCompliance)
  End If
  Call SetParam("SetupDCOM",                 strSetupDCOM,             "YES", "", strListCompliance)
  Call SetParam("SetupFirewall",             strSetupFirewall,         "YES", "", strListCompliance)
  Call SetParam("SetupNetwork",              strSetupNetwork,          "YES", "", strListCompliance)
  Call SetParam("SetupNoTCPNetBios",         strSetupNoTCPNetBios,     "YES", "", strListCompliance)
  Call SetParam("SetupNoWinGlobal",          strSetupNoWinGlobal,      "YES", "", strListCompliance)
  Call SetParam("SetupWinAudit",             strSetupWinAudit,         "YES", "", strListCompliance)

  If strSQLVersion >= "SQL2008" Then
    Call SetParam("SetupNoSSL3",             strSetupNoSSL3,           "YES", "", strListCompliance)
    Call SetParam("SetupTLS12",              strSetupTLS12,            "YES", "", strListCompliance)
  End If

  If strType <> "CLIENT" Then
'    Call SetParam("SetupSSL",                strSetupSSL,               "YES", "", strListCompliance)
    Call SetParam("SetupKerberos",           strSetupKerberos,          "YES", "", strListCompliance)
  End If

  If strSetupSQLAS = "YES" Then
    Call SetParam("SetupOLAP",               strSetupOLAP,             "YES", "", strListCompliance)
  End If

  If strSetupSQLDB = "YES" Then
    Call SetParam("SetupDisableSA",          strSetupDisableSA,        "YES", "", strListCompliance)
    Call SetParam("SetupNonSAAccounts",      strSetupNonSAAccounts,    "YES", "", strListCompliance)
    Call SetParam("SetupOldAccounts",        strSetupOldAccounts,      "YES", "", strListCompliance)
    Call SetParam("SetupStdAccounts",        strSetupStdAccounts,      "YES", "", strListCompliance)
    Call SetParam("SetupSAAccounts",         strSetupSAAccounts,       "YES", "", strListCompliance)
    Call SetParam("SetupSAPassword",         strSetupSAPassword,       "YES", "", strListCompliance)
    Call SetParam("SetupSQLTools",           strSetupSQLTools,         "NO",  "SQL Tools must not be installed for SQL Compliance", "")
    Call SetParam("AuditLevel",              strAuditLevel,            "3",   "Audit Level must be FULL for SQL Compliance", "")
  End If

  If strSetupSQLRS = "YES" Then
    Call SetParam("SetupRSAdmin",            strSetupRSAdmin,          "YES", "", strListCompliance)
  End If

End Sub


Sub SetupDataDep0()
  Call SetProcessId("0BMB", "Setup Parameter Data for Dependency Level 0")
  Dim strRestartSave

  strRestartSave    = GetBuildfileValue("RestartSave")

  Select Case True
    Case Instr(" MULTIDIMENSIONAL TABULAR ", " " & UCase(strASServerMode) & " ") = 0
      Call SetBuildMessage(strMsgErrorConfig, "/ASServerMode: must be either MULTIDIMENSIONAL or TABULAR") 
    Case (strSQLVersion <= "SQL2008R2") And (UCase(strASServerMode) = "TABULAR")
      Call SetBuildMessage(strMsgErrorConfig, "/ASServerMode:TABULAR cannot be used with " & strSQLVersion) 
  End Select

  Select Case True
    Case strType = "CLIENT"
      Call SetParam("SetupSQLDBAG",          strSetupSQLDBAG,          "N/A", "", strListType)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLDBAG",          strSetupSQLDBAG,          "N/A", "", strListSQLDB)
      Call SetParam("SetupSQLAgent",         strSetupSQLAgent,         "N/A", "", strListSQLDB)
    Case strEdition <> "EXPRESS"
      Call SetParam("SetupSQLDBAG",          strSetupSQLDBAG,          "YES", "SQL Agent always installed for Edition " & strEdition, "")
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupCmdshell",         strSetupCmdshell,         "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupDBMail",           strSetupDBMail,           "N/A", "", strListAddNode)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDBMail",           strSetupDBMail,           "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDisableSA",        strSetupDisableSA,        "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strOSVersion <= "5.1"               ' Windows XP
      Call SetParam("SetupABE",              strSetupABE,              "N/A", "", strListOSVersion)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupABE",              strSetupABE,              "N/A", "", strListCore)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strActionAO = "ADDNODE"
      Call SetParam("SetupAODB",             strSetupAODB,             "N/A", "due to Always On action ADDNODE", "")
    Case strActionDAG = "ADDNODE"
      Call SetParam("SetupAODB",             strSetupAODB,             "N/A", "due to Distributed Availability Group action ADDNODE", "")
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strClusterAction = ""
      Call SetParam("SetupClusterShares",    strSetupClusterShares,    "N/A", "", strListCluster)
    Case strSetupClusterShares = ""
      strSetupClusterShares = "NO"
  End Select

  Select Case True 
    Case strClusterAction = ""
      Call SetParam("SetupAPCluster",        strSetupAPCluster,        "N/A", "", strListCluster)
    Case strSetupAPCluster = ""
      strSetupAPCluster = "NO"
  End Select

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case strClusterHost = "YES"
      ' Nothing
    Case strAGName <> ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/AGName: parameter must be specified for AlwaysOn with no Cluster")
  End Select

  Select Case True 
    Case strClusterHost <> "YES"
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "N/A", "", strListCluster)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupAlwaysOn = "YES"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "N/A", "", strListCluster)
    Case strSetupAPCluster = "YES"
      Call SetParam("SetupPolyBaseCluster",  strSetupPolyBaseCluster,  "YES", "PolyBase Cluster mandatory for /SetupAPCluster:" & strSetupAPCluster, "")
  End Select

  Select Case True
    Case strSetupPolyBase = ""
      strSetupPolyBase        = "NO"
      strSetupPolyBaseCluster = "N/A"
    Case strSetupPolyBase <> "YES"
      strSetupPolyBaseCluster = "N/A"
    Case strSetupSQLDBCluster <> "YES"
      strSetupPolyBaseCluster = "N/A"
    Case strSetupPolyBaseCluster = ""
      strSetupPolyBaseCluster = "YES"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSQLVersion >= "SQL2019"
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListSQLVersion)
      If strSetupSQLRS = "YES" Then
        Call SetParam("SetupNet3",           strSetupNet3,             "YES", ".Net3 mandatory for configuration of SSRS", "")
      End If
      If strSetupNet3 = "" Then
        Call SetParam("SetupNet3",           strSetupNet3,             "N/A", "", strListSQLVersion)
      End If
    Case strSQLVersion >= "SQL2016"
      Call SetParam("SetupNet4x",            strSetupNet4x,            "YES", ".Net4.6.1 or above mandatory for " & strSQLVersion, "")
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListSQLVersion)
      If strSetupSQLRS = "YES" Then
        Call SetParam("SetupNet3",           strSetupNet3,             "YES", ".Net3 mandatory for configuration of SSRS", "")
      End If
      If strSetupNet3 = "" Then
        Call SetParam("SetupNet3",           strSetupNet3,             "N/A", "", strListSQLVersion)
      End If
    Case strSQLVersion = "SQL2014"
      Call SetParam("SetupNet3",             strSetupNet3,             "YES", ".Net3 mandatory for " & strSQLVersion, "")
      Call SetParam("SetupNet4",             strSetupNet4,             "YES", ".Net4 mandatory for " & strSQLVersion, "")
      Call SetParam("SetupNet4x",            strSetupNet4x,            "YES", ".Net4.5 or above mandatory for " & strSQLVersion, "")
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListSQLVersion)
    Case strSQLVersion = "SQL2012"
      Call SetParam("SetupNet3",             strSetupNet3,             "YES", ".Net3 mandatory for " & strSQLVersion, "")
      Call SetParam("SetupNet4",             strSetupNet4,             "YES", ".Net4 mandatory for " & strSQLVersion, "")
      Call SetParam("SetupNet4x",            strSetupNet4x,            "YES", ".Net4.5 or above is recommended for " & strSQLVersion, "")
    Case strSQLVersion = "SQL2008"
      Call SetParam("SetupNet3",             strSetupNet3,             "YES", ".Net3 mandatory for " & strSQLVersion, "")
    Case strSQLVersion = "SQL2008R2"
      Call SetParam("SetupNet3",             strSetupNet3,             "YES", ".Net3 mandatory for " & strSQLVersion, "")
    Case strSQLVersion <= "SQL2005"
      If strOSVersion >= "6.1" Then
        Call SetParam("SetupNet3",           strSetupNet3,             "YES", ".Net3 mandatory for " & strSQLVersion, "")
      End If 
  End Select

  Select Case True
    Case strOSVersion > "6.0"
      ' Nothing
    Case strVersionNet3 > ""
      ' Nothing
    Case strVersionNet4 > ""
      Call SetBuildMessage(strMsgErrorConfig, ".Net3 must be installed before .Net 4")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListSQLVersion)
    Case strEditionEnt <> "YES"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListEdition)
    Case strOSVersion < "6.2"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListOSVersion)
    Case strSQLVersion >= "SQL2017"
      ' Nothing
    Case strClusterHost <> "YES"
      Call SetParam("SetupAlwaysOn",         strSetupAlwaysOn,         "N/A", "", strListCluster)
  End Select
  If strSetupAlwaysOn = "" Then
    strSetupAlwaysOn = "NO"
  End If
  If strSetupAlwaysOn <> "YES" Then
    strSetupAOAlias = "N/A"
    strSetupAODB    = "N/A"
    strSetupAOProcs = "N/A"
  End If

  Select Case True
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupAnalytics",        strSetupAnalytics,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strExtSvcAccount = "" 
      Call SetBuildMessage(strMsgErrorConfig, "/EXTSVCACCOUNT: parameter must be specified for Analytics")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupBIDS",             strSetupBIDS,             "N/A", "", strListSQLTools)
      Call SetParam("SetupBOL",              strSetupBOL,              "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupBPAnalyzer",       strSetupBPAnalyzer,       "N/A", "", strListSQLTools)
    Case strSQLVersion = "SQL2008" 
      Call SetParam("SetupBPAnalyzer",       strSetupBPAnalyzer,       "N/A", "", strListSQLVersion)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2014" 
      Call SetParam("SetupBPE",              strSetupBPE,              "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupBPE",              strSetupBPE,              "N/A", "", strListSQLDB)
  End Select
  If strSetupBPE <> "YES" Then
    Call SetBuildfileValue("DirBPE", "")
  End If

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupCacheManager",     strSetupCacheManager,     "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
   Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDB2OLE",           strSetupDB2OLE,           "N/A", "", strListSQLDB)
    Case strEditionEnt <> "YES"
      Call SetParam("SetupDB2OLE",           strSetupDB2OLE,           "N/A", "", strListEdition)
    Case (strSQLVersion <= "SQL2008R2") And (strOSVersion >= "6.3A")
      Call SetParam("SetupDB2OLE",           strSetupDB2OLE,           "N/A", "", strListOSVersion)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupDB2OLE",           strSetupDB2OLE,           "N/A", "", strListOSVersion)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDBAManagement",    strSetupDBAManagement,    "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDBOpts",           strSetupDBOpts,           "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
   Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDistributor",      strSetupDistributor,      "N/A", "", strListSQLDB)
    Case strType = "CLIENT"
      Call SetParam("SetupDistributor",      strSetupDistributor,      "N/A", "", strListType)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupDistributor",      strSetupDistributor,      "N/A", "", strListAddNode)
  End Select

  Select Case True  
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListSQLDB)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListCore)
    Case strMainInstance <> "YES"
      Call SetParam("SetupDQ",               strSetupDQ,               "NO",  "", strListMain)
    Case strEdition = "EXPRESS"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListEdition)
    Case strEdition = "WORKGROUP"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListEdition)
    Case strType = "CLIENT"
      Call SetParam("SetupDQ",               strSetupDQ,               "N/A", "", strListType)
  End Select

  Select Case True              
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing 
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListCore)
    Case strEdition = "EXPRESS"
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListEdition)
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupDQC",              strSetupDQC,              "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDRUClt",           strSetupDRUClt,           "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupDRUClt",           strSetupDRUClt,           "N/A", "", strListCore)
    Case strSetupDRUClt <> "YES"
      strSetupDRUClt = "NO"
  End Select
  If strSetupDRUClt <> "YES" Then
    Call SetBuildfileValue("DirDRU", "")
  End If

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupDRUCtlr",          strSetupDRUCtlr,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupDRUCtlr = "N/A"
      ' Nothing
    Case strSetupDRUCtlr <> "YES"
      strSetupDRUCtlr = "NO"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strClusterAction = ""
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "N/A", "", strListCluster)
    Case strOSVersion >= "6.0"
      ' Nothing
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strDTCClusterRes > ""
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "N/A", "DTC Cluster already exists", "")
    Case Else
      Call SetParam("SetupDTCCluster",       strSetupDTCCluster,       "YES", "DTC Cluster is mandatory for Cluster install", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case (strOSVersion < "6.0") And (strDTCClusterRes > "")
      ' Nothing
    Case strSetupDTCCluster = "YES"
      Call SetParam("SetupDTCNetAccess",     strSetupDTCNetAccess,     "YES", "DTC Network Access mandatory for Cluster install", "")
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupDTCNetAccess",     strSetupDTCNetAccess,     "YES", "DTC Network Access mandatory for " & strSQLVersion & " install", "")
  End Select

  Select Case True
    Case strSetupDTCNetAccess = "YES"
      ' Nothing
    Case strSetupDTCNetAccessStatus = strStatusComplete
      strSetupDTCNetAccessStatus = strStatusPreConfig
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupDTSDesigner",      strSetupDTSDesigner,      "N/A", "", strListSQLTools)
    Case strSQLVersion >= "SQL2012" 
      Call SetParam("SetupDTSDesigner",      strSetupDTSDesigner,      "N/A", "", strListSQLVersion)
    Case strOSVersion >= "6.3A"
      Call SetParam("SetupDTSDesigner",      strSetupDTSDesigner,      "N/A", "", strListOSVersion)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupDimensionSCD",     strSetupDimensionSCD,     "N/A", "", strListSSIS)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupDimensionSCD",     strSetupDimensionSCD,     "N/A", "", strListCore)
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupGovernor",         strSetupGovernor,         "N/A", "", strListSQLDB)
    Case strSetupGovernor <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupGovernor",         strSetupGovernor,         "YES", "Resource Governor recommended for " & strEdition & " Edition", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case (Instr(Ucase(strOSName), " XP") > 0) And (Instr(strOSType, "STARTER") > 0)
      Call SetParam("SetupGenMaint",         strSetupGenMaint,         "N/A", "", strListOSVersion)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupGenMaint",         strSetupGenMaint,         "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2008R2" 
      Call SetParam("SetupIntViewer",        strSetupIntViewer,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupIntViewer",        strSetupIntViewer,        "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2017" 
      Call SetParam("SetupISMaster",         strSetupISMaster,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupISMaster",         strSetupISMaster,         "N/A", "", strListType)
    Case strSetupISMaster = ""
      strSetupISMaster = "NO"
    Case strSetupISMaster <> "YES"
      ' Nothing
    Case Else
      Call SetParam("SecurityMode",          strSecurityMode,          "SQL", "/SecurityMode:SQL mandatory for ISMaster", "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2017" 
      Call SetParam("SetupISWorker",         strSetupISWorker,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupISWorker",         strSetupISWorker,         "N/A", "", strListType)
    Case strSetupISWorker <> ""
      ' Nothing
    Case strSetupISMaster = "YES"
      strSetupISWorker = "YES"
    Case Else
      strSetupISWorker = "NO"
  End Select
  Select Case True
    Case strSetupISWorker <> "YES"
      ' Nothing
    Case Else
      Call SetParam("SecurityMode",          strSecurityMode,          "SQL", "/SecurityMode:SQL mandatory for ISWorker", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupJavaDBC",          strSetupJavaDBC,          "N/A", "", strListSQLDB)
  End Select

  Select Case True   
    Case strType = "REBUILD"
      ' Nothing            
    Case strSetupPolyBase <> "YES"
      ' Nothing
    Case Else
      Call SetParam("SetupJRE",              strSetupJRE,              "YES", "JRE mandatory if PolyBase is installed", "")
  End Select

  Select Case True   
    Case strType = "REBUILD"
      ' Nothing            
    Case strSetupSQLDB = "YES"
      ' Nothing
    Case Else
      Call SetParam("SetupStartJob",         strSetupStartJob,         "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      Call SetParam("SetupKB925336",         strSetupKB925336,         "N/A", "", strListOSVersion)
    Case (Instr(Ucase(strOSName), " XP") > 0) And (strFileArc = "X86")
      Call SetParam("SetupKB925336",         strSetupKB925336,         "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case Else
      Call SetParam("SetupKB925336",         strSetupKB925336,         "YES", "KB925336 mandatory on " & strOSName, "")
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      Call SetParam("SetupKB933789",         strSetupKB933789,         "N/A", "", strListOSVersion)
    Case (Instr(Ucase(strOSName), " XP") > 0) And (strFileArc = "X86")
      Call SetParam("SetupKB933789",         strSetupKB933789,         "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case Else
      Call SetParam("SetupKB933789",         strSetupKB933789,         "YES", "KB933789 mandatory on " & strOSName, "")
  End Select

  Select Case True
    Case strOSVersion >= "6.0"
      Call SetParam("SetupKB937444",         strSetupKB937444,         "N/A", "", strListOSVersion)
    Case (Instr(Ucase(strOSName), " XP") > 0) And (strFileArc = "X86")
      Call SetParam("SetupKB937444",         strSetupKB937444,         "N/A", "", strListOSVersion)
    Case strSQLVersion <= "SQL2005" 
      Call SetParam("SetupKB937444",         strSetupKB937444,         "N/A", "", strListSQLVersion)
    Case strType = "CLIENT"
      Call SetParam("SetupKB937444",         strSetupKB937444,         "N/A", "", strListType)
    Case strType = "REBUILD"
      ' Nothing
    Case Else
      Call SetParam("SetupKB937444",         strSetupKB937444,         "YES", "KB937444 mandatory on " & strOSName, "")
  End Select

  Select Case True
    Case strOSVersion <> "6.3"
      Call SetParam("SetupKB2919355",        strSetupKB2919355,        "N/A", "", strListOSVersion)
      Call SetParam("SetupKB2919442",        strSetupKB2919442,        "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case Else
      Call SetParam("SetupKB2919355",        strSetupKB2919355,        "YES", "KB92919355 mandatory on " & strOSName, "")
      Call SetParam("SetupKB2919442",        strSetupKB2919442,        "YES", "KB92919422 mandatory on " & strOSName, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008R2"
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListSQLVersion)
      Call SetParam("SetupMDSC",             strSetupMDSC,             "N/A", "", strListSQLVersion)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListCore)
      Call SetParam("SetupMDSC",             strSetupMDSC,             "N/A", "", strListCore)
    Case strFileArc <> "X64"
      Call SetParam("SetupMDS",              strSetupMDS,              "NO",  "Master Data Services can only installed on X64", "")
    Case strType = "REBUILD"
      ' Nothing
   Case strSetupSQLDB <> "YES"
      Call SetParam("SetupMDS",              strSetupMDS,              "N/A", "", strListSQLDB)
    Case strMainInstance <> "YES"
      Call SetParam("SetupMDS",              strSetupMDS,              "NO",  "", strListMain)
  End Select
  Select Case True
    Case strSetupMDS <> "YES"
      ' Nothing
    Case Else
      If strMDSPort = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/MDSPort: parameter must be specified for MDS")
      End If
      If strMDSSite = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/MDSSite: parameter must be specified for MDS")
      End If
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupMDXStudio",        strSetupMDXStudio,        "N/A", "", strListSQLTools)
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      Call SetParam("SetupMDXStudio",        strSetupMDXStudio,        "N/A", "", strListSSAS)
  End Select

  Select Case True 
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A", "", strListType)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupManagementDW",     strSetupManagementDW,     "N/A", "", strListAddNode)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupMyDocs",           strSetupMyDocs,           "N/A", "", strListSQLTools)
    Case strEdition = "EXPRESS"
      Call SetParam("SetupMyDocs",           strSetupMyDocs,           "N/A", "", strListEdition)
  End Select

  If strSetupNetBind = "" Then
    Call SetParam("SetupNetBind",            strSetupNetBind,          "N/A", "NetBind processing not required", "")
  End If

  If strSetupNetName = "" Then
    Call SetParam("SetupNetName",            strSetupNetName,          "N/A", "NetName processing not required", "")
  End If

  
  Select Case True
    Case strType <> "REBUILD"
      ' Nothing
    Case strRestartSave = ""
      Call SetBuildMessage(strMsgErrorConfig, "/Restart: parameter must be specified for /Type:REBUILD")
    Case strRestartSave = "YES"
      Call SetBuildMessage(strMsgErrorConfig, "/Restart: ProcessId must be specified for /Type:REBUILD")
    Case strStopAt = ""
      Call SetParam("StopAt",                strStopAt,                strRestartSave,"/StopAt: set to automatic stop", "")
    Case strStopAt = "YES"
      Call SetParam("StopAt",                strStopAt,                strRestartSave,"/StopAt: set to automatic stop", "")
  End Select

  Select Case True 
    Case strOSVersion < "6.0"
      Call SetParam("SetupNoDefrag",         strSetupNoDefrag,         "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupNoDefrag <> ""
      ' Nothing
    Case strType = "WORKSTATION"
      Call SetParam("SetupNoDefrag",         strSetupNoDefrag,         "NO",  "", strListType)
    Case Else
      strSetupNoDefrag = "YES"
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2008"
      Call SetParam("SetupNoSSL3",           strSetupNoSSL3,           "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupNoSSL3 <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupNoSSL3",           strSetupNoSSL3,           "YES", "No SSL3 is reccommended for Security Compliance", "")
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupNonSAAccounts",    strSetupNonSAAccounts,    "N/A", "", strListType)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupNonSAAccounts",    strSetupNonSAAccounts,    "N/A", "", strListSQLDB)
    Case strGroupDBANonSA = ""
      Call SetParam("SetupNonSAAccounts",    strSetupNonSAAccounts,    "NO",  "Non-sa Accounts can not be configured when /GroupDBANonSA: is blank", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupOLAP",             strSetupOLAP,             "N/A", "", strListType)
    Case strSetupSQLAS <> "YES"
      Call SetParam("SetupOLAP",             strSetupOLAP,             "N/A", "", strListSSAS)
    Case strActionSQLAS = "ADDNODE"
      Call SetParam("SetupOLAP",             strSetupOLAP,             "N/A", "", strListAddNode)
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupOLAPAPI",          strSetupOLAPAPI,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupOLAPAPI",          strSetupOLAPAPI,          "N/A", "", strListType)
    Case strSetupSQLAS <> "YES"
      Call SetParam("SetupOLAPAPI",          strSetupOLAPAPI,          "N/A", "", strListSSAS)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupOldAccounts",      strSetupOldAccounts,      "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupParam",            strSetupParam,            "N/A", "", strListSQLDB)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupParam",            strSetupParam,            "N/A", "", strListAddNode)
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupPBM",              strSetupPBM,              "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupPBM",              strSetupPBM,              "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupPDFReader",        strSetupPDFReader,        "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2016"
      Call SetParam("SetupPerfDash",         strSetupPerfDash,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupPerfDash",         strSetupPerfDash,         "N/A", "", strListSQLTools)
    Case strUseFreeSSMS = "YES"
      Call SetParam("SetupPerfDash",         strSetupPerfDash,         "N/A", "Performance Dashboard included with Free SSMS", "")
  End Select

  Select Case True
    Case strOSVersion < "6"
      Call SetParam("SetupPlanExplorer",     strSetupPlanExplorer,     "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupPlanExplorer",     strSetupPlanExplorer,     "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupPlanExpAddin",     strSetupPlanExpAddin,     "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case Instr(Ucase(strOSName), " XP") > 0
      Call SetParam("SetupPowerCfg",         strSetupPowerCfg,         "N/A", "", strListOSVersion)
    Case Instr(Ucase(strOSName), "Vista") > 0
      Call SetParam("SetupPowerCfg",         strSetupPowerCfg,         "N/A", "", strListOSVersion)
    Case Instr(Ucase(strOSName), "Windows 7") > 0
      Call SetParam("SetupPowerCfg",         strSetupPowerCfg,         "N/A", "", strListOSVersion)
    Case Instr(Ucase(strOSName), "Windows 8") > 0
      Call SetParam("SetupPowerCfg",         strSetupPowerCfg,         "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPowerCfg <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupPowerCfg",         strSetupPowerCfg,         "YES", "Power Configuration recommended with " & strOSName, "")
  End Select

  Select Case True
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupProcExp",          strSetupProcExp,          "N/A", "", strListCore)
  End Select

  Select Case True
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupProcMon",          strSetupProcMon,          "N/A", "", strListCore)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupPowerBIDesktop",   strSetupPowerBIDesktop,   "N/A", "", strListSQLVersion)
    Case strOSVersion < "6.2"
      Call SetParam("SetupPowerBIDesktop",   strSetupPowerBIDesktop,   "N/A", "", strListOSVersion)
    Case strFileArc = "X86"
      Call SetParam("SetupPowerBIDesktop",   strSetupPowerBIDesktop,   "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupPowerBIDesktop",   strSetupPowerBIDesktop,   "N/A", "", strListSQLTools)
    Case strSetupPowerBIDesktop <> ""
      ' Nothing
    Case Else
      strSetupPowerBIDesktop = "NO"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2017"
      Call SetParam("SetupPython",           strSetupPython,           "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupPython",           strSetupPython,           "N/A", "", strListSQLDB)
    Case strSetupPython <> ""
      ' Nothing
    Case Else
      strSetupPython = "NO"
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2012" 
      Call SetParam("SetupRawReader",        strSetupRawReader,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupRawReader",        strSetupRawReader,        "N/A", "", strListSQLTools)
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupRawReader",        strSetupRawReader,        "N/A", "", strListSSIS)
    Case strSetupBIDS <> "YES"
      Call SetParam("SetupRawReader",        strSetupRawReader,        "NO",  "SSIS Raw File Reader can not be installed when BIDS is not installed", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupRMLTools",         strSetupRMLTools,         "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupRptTaskPad",       strSetupRptTaskPad,       "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupRSAdmin",          strSetupRSAdmin,          "N/A", "", strListSQLRS)
      Call SetParam("SetupRSAlias",          strSetupRSAlias,          "N/A", "", strListSQLRS)
      Call SetParam("SetupRSExec",           strSetupRSExec,           "N/A", "", strListSQLRS)
      Call SetParam("SetupRSIndexes",        strSetupRSIndexes,        "N/A", "", strListSQLRS)
      Call SetParam("SetupRSKeepAlive",      strSetupRSKeepAlive,      "N/A", "", strListSQLRS)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupRSIndexes",        strSetupRSIndexes,        "N/A", "", strListSQLDB)
      Call SetParam("SetupRSKeepAlive",      strSetupRSKeepAlive,      "N/A", "", strListSQLDB)
    Case Else
      If strSetupRSIndexes = "" Then
        Call SetParam("SetupRSIndexes",      strSetupRSIndexes,        "YES", "RSIndexes Recommended when SSRS installed", "")
      End If
      If strSetupRSKeepAlive = "" Then
        Call SetParam("SetupRSKeepAlive",    strSetupRSKeepAlive,      "YES", "RSKeepAlive Recommended when SSRS installed", "")
      End If
  End Select

  Select Case True
    Case Instr(strOSType, "SERVER") > 0
      Call SetParam("SetupRSAT",             strSetupRSAT,             "N/A", "", strListOSVersion)
    Case strOSVersion < "6.0"
      Call SetParam("SetupRSAT",             strSetupRSAT,             "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupRSAT > ""
      ' Nothing
    Case Else
      Call SetParam("SetupRSAT",             strSetupRSAT,             "YES", "RSAT is recommended for Client OS", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupRSLinkGen",        strSetupRSLinkGen,        "N/A", "", strListSQLTools)
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupRSLinkGen",        strSetupRSLinkGen,        "N/A", "", strListSQLRS)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupRSScripter",       strSetupRSScripter,       "N/A", "", strListSQLTools)
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupRSScripter",       strSetupRSScripter,       "N/A", "", strListSQLRS)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupRServer",          strSetupRServer,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupRServer",          strSetupRServer,          "N/A", "", strListSQLDB)
    Case strSetupRServer <> ""
      ' Nothing
    Case Else
      strSetupRServer = "NO"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupSemantics",        strSetupSemantics,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupSemantics",        strSetupSemantics,        "N/A", "", strListType)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSemantics",        strSetupSemantics,        "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupSQLAgent",         strSetupSQLAgent,         "N/A", "", strListType)
    Case strSetupSQLDBAG <> "YES"
      Call SetParam("SetupSQLAgent",         strSetupSQLAgent,         "NO",  "SQL Agent can not be configured when it is not installed", "")
  End Select

  Select Case True      
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupSQLBC",            strSetupSQLBC,            "N/A", "", strListCore)
   End Select

  Select Case True
    Case strSQLVersion < "SQL2008"
      Call SetParam("SetupSQLDBFS",          strSetupSQLDBFS,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLDBFS",          strSetupSQLDBFS,          "N/A", "", strListSQLDB)
    Case strSetupSQLDBFS <> "YES"
      strSetupSQLDBFS  = "NO"
    Case strFSLevel = "0"
      strSetupSQLDBFS  = "NO"
    Case (strFileArc = "X64") And (strWOWX86 = "TRUE")
      Call SetParam("SetupSQLDBFS",          strSetupSQLDBFS,          "N/A", "Filestream not available on WOW install", "")
    Case strSQLVersion >= "SQL2017"
      ' Nothing
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("VolDataFSType") <> "C" 
      Call SetBuildMessage(strMsgErrorConfig, "/VolDataFS: parameter must specify a clustered disk")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLDBFT",          strSetupSQLDBFT,          "N/A", "", strListSQLDB)
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strSetupSQLDBFT <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupSQLDBFT",          strSetupSQLDBFT,          "YES", "SQL Full Text is recommended for Cluster install", "")
  End Select

  Select Case True
    Case strSetupFirewall <> "YES"
      Call SetParam("SetupSQLDebug",         strSetupSQLDebug,         "N/A", "/SetupSQLDebug: not available for /SetupFirewall:" & strSetupFireWall, "")
    Case strSetupSQLDebug <> ""
      ' Nothing
    Case strEdition = "DEVELOPER"
      strSetupSQLDebug = "YES"
    Case Else
      strSetupSQLDebug = "NO"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupSQLInst",          strSetupSQLInst,          "N/A", "", strListType)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupSQLInst",          strSetupSQLInst,          "N/A", "", strListAddNode)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLInst",          strSetupSQLInst,          "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strFileArc <> "X86"
      Call SetParam("SetupSQLMail",          strSetupSQLMail,          "N/A", "", strListOSVersion)
    Case strSQLVersion >= "SQL2012"
      Call SetParam("SetupSQLMail",          strSetupSQLMail,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLMail",          strSetupSQLMail,          "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strOSVersion >= "6.2"
      Call SetParam("SetupSQLNS",            strSetupSQLNS,            "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strMainInstance <> "YES"
      Call SetParam("SetupSQLNS",            strSetupSQLNS,            "NO",  "", strListMain)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLDBRepl",        strSetupSQLDBRepl,        "N/A", "", strListSQLDB)
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strSetupSQLDBRepl <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupSQLDBRepl",        strSetupSQLDBRepl,        "YES", "SQL Replication is recommended for Cluster install", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupSQLNexus",         strSetupSQLNexus,         "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListSQLRS)
    Case strClusterAction = ""
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "N/A", "", strListCluster)
    Case strSetupSQLRSCluster <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupSQLRSCluster",     strSetupSQLRSCluster,     "YES", "SSRS Cluster will be installed automatically", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "CLIENT"
      Call SetParam("SetupSQLServer",        strSetupSQLServer,        "N/A", "", strListType)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupSQLServer",        strSetupSQLServer,        "N/A", "", strListAddNode)
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSQLServer",        strSetupSQLServer,        "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupSSDTBI",           strSetupSSDTBI,           "N/A", "", strListSQLVersion)
    Case strSQLVersion >= "SQL2017"
      Call SetParam("SetupSSDTBI",           strSetupSSDTBI,           "N/A", "", strListSQLVersion)
    Case Instr(strOSType, "CORE") > 0
      Call SetParam("SetupSSDTBI",           strSetupSSDTBI,           "N/A", "", strListCore)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupSSDTBI",           strSetupSSDTBI,           "N/A", "", strListSQLTools)
    Case strSetupSSDTBI = ""
      Call SetParam("SetupSSDTBI",           strSetupSSDTBI,           "YES", "SSDTBI Recommended for " & strSQLVersion, "")
  End Select

  Select Case True 
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupSSISDB",           strSetupSSISDB,           "N/A", "", strListSSIS)
    Case strSQLVersion <= "SQL2008R2"
      Call SetParam("SetupSSISDB",           strSetupSSISDB,           "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSSISDB",           strSetupSSISDB,           "N/A", "", strListSQLDB)
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupSSISDB",           strSetupSSISDB,           "N/A", "", strListAddNode)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupSSMS",             strSetupSSMS,             "N/A", "", strListSQLTools)
    Case strSetupSSMS <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupSSMS",             strSetupSSMS,             "YES", "SSMS is recommended for " & strSQLVersion, "")
  End Select

  Select Case True
    Case strSetupSSMS <> "YES"
      ' Nothing
    Case strOSVersion <= "6.0"
      Call SetParam("UseFreeSSMS",           strUseFreeSSMS,           "N/A", "", strListOSVersion)
    Case strSSMSexe = ""
      Call SetParam("UseFreeSSMS",           strUseFreeSSMS,           "NO",  "SSMS install file not found", "")
    Case strUseFreeSSMS <> ""
      ' Nothing
    Case strSQLVersion >= "SQL2016"
      Call SetParam("UseFreeSSMS",           strUseFreeSSMS,           "YES", "/UseFreeSSMS: is required for " & strSQLVersion, "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSSMS <> "YES"
      ' Nothing
    Case strUseFreeSSMS <> "YES"
      ' Nothing
    Case Else
      Call SetParam("SetupNet4",             strSetupNet4,             "YES", ".Net4.0 or above mandatory for SSMS", "")
      Call SetParam("SetupNet4x",            strSetupNet4x,            "YES", ".Net4.5 or above mandatory for SSMS", "")
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSAAccounts",       strSetupSAAccounts,       "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupStdAccounts",      strSetupStdAccounts,      "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008R2"
      Call SetParam("SetupStreamInsight",    strSetupStreamInsight,    "N/A", "", strListSQLVersion)
    Case strOSVersion < "6.0"
      Call SetParam("SetupStreamInsight",    strSetupStreamInsight,    "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupStreamInsight <> "YES"
      ' Nothing
    Case strEdition = "ENTERPRISE EVALUATION"
      ' Nothing
    Case strType = "CLIENT"
      ' Nothing
    Case strPID = ""
      Call SetBuildMessage(strMsgErrorConfig, "/PID: is mandatory for StreamInsight")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      Call SetParam("SetupStretch",          strSetupStretch,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupStretch",          strSetupStretch,          "N/A", "", strListSQLDB)
    Case strSetupStretch = ""
      Call SetParam("SetupStretch",          strSetupStretch,          "NO",  "Stretch Database default is No", "")
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSysDB",            strSetupSysDB,            "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSysIndex",         strSetupSysIndex,         "N/A", "", strListSQLDB)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupSysManagement",    strSetupSysManagement,    "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupSystemViews",      strSetupSystemViews,      "N/A", "", strListSQLTools)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      Call SetParam("SetupTempDB",           strSetupTempDB,           "N/A", "", strListSQLDB)
  End Select

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupTLS12",            strSetupTLS12,            "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
   Case strSetupTLS12 = ""
      strSetupTLS12 = "YES"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupTrouble",          strSetupTrouble,          "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strOSVersion < "6.0"
      Call SetParam("SetupVC2010",           strSetupVC2010,           "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupDB2OLE = "YES"
      Call SetParam("SetupVC2010",           strSetupVC2010,           "YES", "Visual C 2010 is mandatory for DB2OLEDB", "")
    Case strSetupVC2010 <> ""
      ' Nothing
    Case Else
      strSetupVC2010 = "NO"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupVS",               strSetupVS,               "N/A", "", strListSQLTools)
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupVS",               strSetupVS,               "N/A", "", strListSSIS)
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupWinAudit <> ""
      ' Nothing
    Case strType = "WORKSTATION"
      Call SetParam("SetupWinAudit",         strSetupWinAudit,         "N/A", "", strListType)
    Case Else
      strSetupWinAudit = "YES"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupWindows",          strSetupWindows,          "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008"
      Call SetParam("SetupXEvents",          strSetupXEvents,          "N/A", "", strListSQLVersion)
    Case strSQLVersion >= "SQL2012" 
      Call SetParam("SetupXEvents",          strSetupXEvents,          "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupXEvents",          strSetupXEvents,          "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupXMLNotepad",       strSetupXMLNotepad,       "N/A", "", strListSQLTools)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupZoomIt",           strSetupZoomIt,           "N/A", "", strListSQLTools)
  End Select

  Select Case True ' See https://docs.microsoft.com/en-us/visualstudio/debugger/remote-debugger-port-assignments
    Case strTCPPortDebug <> ""
      ' Nothing
    Case strSQLVersion = "SQL2012" ' Default for VS 2012 4016
      strTCPPortDebug = "4016,4018,4020"
    Case strSQLVersion = "SQL2014" ' Default for VS 2013 4018
      strTCPPortDebug = "4016,4018,4020,4022"
    Case strSQLVersion = "SQL2016" ' Default for VS 2015 4020
      strTCPPortDebug = "4018,4020,4022"
    Case strSQLVersion = "SQL2017" ' Default for VS 2017 4022
      strTCPPortDebug = "4020,4022, 4024"
    Case strSQLVersion = "SQL2019" ' Default for VS 2019 4024
      strTCPPortDebug = "4022, 4024, 4026"
    Case strSQLVersion > "SQL2019" ' Default for VS vNext 4026
      strTCPPortDebug = "4024, 4026, 4028"
  End Select
  If strTCPPortDebug <> "" THen
    Call ParamListAdd("ListNonDefault", "TCPPortDebug")
  End If

End Sub


Sub SetupDataDep1()
  Call SetProcessId("0BMC", "Setup Parameter Data for Dependency Level 1")

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupBIDSHelper",       strSetupBIDSHelper,       "N/A", "", strListSQLTools)
    Case (strSQLVersion >= "SQL2019")
      ' Nothing
    Case (strSQLVersion <= "SQL2008R2") And (strSetupBIDS <> "YES")
      Call SetParam("SetupBIDSHelper",       strSetupBIDSHelper,       "NO",  "BIDS Helper can not be installed when BIDS is not installed", "")
    Case (strSQLVersion >= "SQL2012") And (strSetupSSDTBI <> "YES")
      Call SetParam("SetupBIDSHelper",       strSetupBIDSHelper,       "NO",  "BIDS Helper can not be installed when SSDTBI is not installed", "")
  End Select
  Select Case True
    Case strType = "CLIENT"
      ' Nothing
    Case strSetupSQLIS <> "YES"
      Call SetParam("SetupBIDSHelper",       strSetupBIDSHelper,       "N/A", "", strListSSIS)
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupDTSBackup",        strSetupDTSBackup,        "N/A", "", strListSQLTools)
    Case strSetupDTSDesigner <> "YES"
      Call SetParam("SetupDTSBackup",        strSetupDTSBackup,        "NO",  "DTS Backup can not be installed when DTS Designer is not installed", "")
  End Select

  strFSInstLevel   = strFSLevel
  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLDBFS <> "YES"
      strFSLevel     = "0"
      strFSInstLevel = strFSLevel
      strFSShareName = ""
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing
    Case strFSLevel > "0"
      Call SetParam("FSLevel",               strFSLevel,               "3",   "Required level for Filestream for " & strSQLVersion & " Cluster install", "")
      strFSInstLevel   = "0"
  End Select

  Select Case True 
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupDTSDesigner = "YES"
      Call SetParam("SetupSQLBC",            strSetupSQLBC,            "YES", "Backward Compatibility mandatory for DTS Designer", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case (strProcessId > "1") And (strProcessId < "7")
      strSetupDTCCID = GetBuildfileValue("SetupDTCCID")
    Case strSetupDTCNetAccessStatus = strStatusComplete
      Call SetParam("SetupDTCCID",           strSetupDTCCID,           "NO",  "MSDTC CID is already configured", "")
    Case strActionSQLDB = "ADDNODE"
      Call SetParam("SetupDTCCID",           strSetupDTCCID,           "N/A", "", strListAddnode)
    Case strSetupDTCNetAccess = "YES" 
      Call SetParam("SetupDTCCID",           strSetupDTCCID,           "YES", "New MSDTC CID mandatory for DTC Net Access", "")
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2005"
      ' Nothing
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLRS = "YES" 
      Call SetParam("SetupIIS",              strSetupIIS,              "YES", "IIS mandatory for Reporting Services", "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2019"
      Call SetParam("SetMemOptTempdb",        strSetMemOptTempdb,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
  End Select

  Select Case True
    Case strOSVersion < "6.1"
      Call SetParam("SetupKB2854082",        strSetupKB2854082,        "N/A", "", strListOSVersion)
    Case strOSVersion > "6.2"
      Call SetParam("SetupKB2854082",        strSetupKB2854082,        "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupAlwaysOn = "YES"
      Call SetParam("SetupKB2854082",        strSetupKB2854082,        "YES",  "KB2854082 is rquired for Availability Groups", "")
    Case strSetupKB2854082 = ""
      strSetupKB2854082 = "NO"
  End Select

  Select Case True
    Case strOSVersion < "6"
      Call SetParam("SetupKB2862966",        strSetupKB2862966,        "N/A", "", strListOSVersion)
    Case strOSVersion > "6.2"
      Call SetParam("SetupKB2862966",        strSetupKB2862966,        "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strUseFreeSSMS = "YES"
      Call SetParam("SetupKB2862966",        strSetupKB2862966,        "YES",  "KB2862966 is recommended for SSMS", "")
    Case strSetupKB2862966 = ""
      strSetupKB2862966 = "NO"
  End Select

  Select Case True
    Case strSQLVersion < "SQL2008R2"
      Call SetParam("SetupMDSC",             strSetupMDSC,             "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupMDSC",             strSetupMDSC,             "N/A", "", strListSQLTools)
    Case strSetupMDSC <> ""
      ' Nothing
    Case Else 
      Call SetParam("SetupMDSC",             strSetupMDSC,             "YES", "MDS Client is recommended", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSQLVersion < "SQL2008R2"
      ' Nothing
    Case strSetupMDS = "YES" 
      Call SetParam("SetupIIS",              strSetupIIS,              "YES", "IIS mandatory for Master Data Services", "")
      Call SetParam("SetCLREnabled",         strSetCLREnabled,         "1",   "CLR mandatory for Master Data Services", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case strSetupIIS <> "YES" 
      ' Nothing
    Case Else
      Call SetParam("SetupRSAlias",          strSetupRSAlias,          "YES", "RS Alias required if IIS installed", "")
  End Select    

  Select Case True
    Case strSQLVersion <= "SQL2005"
      Call SetParam("SetupKB954961",         strSetupKB954961,         "N/A", "", strListSQLVersion)
    Case GetBuildfileValue("MenuSQL2005Flag") <> "Y"
      Call SetParam("SetupKB954961",         strSetupKB954961,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupBIDS <> "YES"
      Call SetParam("SetupKB954961",         strSetupKB954961,         "NO",  "KB954961 can not be installed when BIDS is not installed", "")
    Case Else
      Call SetParam("SetupKB954961",         strSetupKB954961,         "YES", "KB954961 is recommended when SQL2005 previously installed", "")
  End Select

  Select Case True
    Case strOSVersion <> "6.0"
      Call SetParam("SetupKB956250",         strSetupKB956250,         "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupNet4 <> "YES"
      Call SetParam("SetupKB956250",         strSetupKB956250,         "NO",  "KB956250 can not be installed when .Net v4 is not installed", "")
    Case Else
      Call SetParam("SetupKB956250",         strSetupKB956250,         "YES", "KB956250 mandatory for .Net v4 on " & strOSName, "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupKB2781514",        strSetupKB2781514,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupKB2781514",        strSetupKB2781514,        "N/A", "", strListSQLTools)
    Case strSetupBPAnalyzer <> "YES" 
      Call SetParam("SetupKB2781514",        strSetupKB2781514,        "NO",  "KB2781514 can not be installed when Best Practice Analyzer is not installed", "")
    Case strSetupKB2781514 = ""
      Call SetParam("SetupKB2781514",        strSetupKB2781514,        "YES", "KB2781514 recommended if Best Practice Analyzer installed", "")
   End Select

  Select Case True
    Case strOSVersion <> "6.3" ' Windows 2012 R2
      Call SetParam("SetupKB3090973",        strSetupKB3090973,        "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupKB2919355 = "YES"
      Call SetParam("SetupKB3090973",        strSetupKB3090973,        "YES", "KB3090973 recommended for Windows 2012 R2 MSDTC", "")
    Case strStatusKB2919355 = ""
      Call SetParam("SetupKB3090973",        strSetupKB3090973,        "N/A", "", strListOSVersion)
    Case Else
      Call SetParam("SetupKB3090973",        strSetupKB3090973,        "YES", "KB3090973 recommended for Windows 2012 R2 MSDTC", "")
  End Select

  Select Case True
    Case strOSVersion < "6.1" ' Windows 2008 R2
      Call SetParam("SetupKB4019990",        strSetupKB4019990,        "N/A", "", strListOSVersion)
    Case (strOSVersion = "6.2" ) And (strOSType = "CLIENT")
      Call SetParam("SetupKB4019990",        strSetupKB4019990,        "N/A", "", strListOSVersion)
    Case strOSVersion > "6.2" ' Windows 2012
      Call SetParam("SetupKB4019990",        strSetupKB4019990,        "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case GetBuildfileValue("Net4Xexe") < "NDP462"
      Call SetParam("SetupKB4019990",        strSetupKB4019990,        "N/A", "", strListOSVersion)
    Case Else
      Call SetParam("SetupKB4019990",        strSetupKB4019990,        "YES", "KB4019990 mandatory for .Net 4.x", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupMBCA",             strSetupMBCA,             "N/A", "", strListSQLTools)
    Case (strSetupBPAnalyzer = "YES") And (strSQLVersion = "SQL2005")
      If strSetupMBCA = "" Then
        strSetupMBCA = "NO"
      End If
    Case (strSetupBPAnalyzer = "YES") And (Instr(Ucase(strOSName), " XP") > 0 )
      Call SetParam("SetupBPAnalyzer",       strSetupBPAnalyzer,       "N/A", "", strListOSVersion)
      Call SetParam("SetupMBCA",             strSetupMBCA,             "N/A", "", strListOSVersion)
    Case (strSetupBPAnalyzer = "YES") And (Instr(Ucase(strOSName), "VISTA") > 0 )
      Call SetParam("SetupBPAnalyzer",       strSetupBPAnalyzer,       "N/A", "", strListOSVersion)
      Call SetParam("SetupMBCA",             strSetupMBCA,             "N/A", "", strListOSVersion)
    Case strSetupBPAnalyzer = "YES" 
      Call SetParam("SetupMBCA",             strSetupMBCA,             "YES", "Microsoft Baseline Configuration Analyzer mandatory for SQL BPA", "")
    Case strSetupMBCA = ""
      strSetupMBCA  = "NO"
  End Select
  If strSetupMBCA = "YES" Then
    Call SetParam("SetupNet3",               strSetupNet3,             "YES", ".Net3 mandatory for Microsoft Baseline Configuration Analyzer", "")
  End If

  Select Case True
    Case strOSVersion < "6.0"
      Call SetParam("SetupNet4x",            strSetupNet4x,            "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPlanExplorer = "YES"
      Call SetParam("SetupNet4x",            strSetupNet4x,            "YES", ".Net4.5 or above mandatory for Plan Explorer", "")
  End Select

  Select Case True
    Case strOSVersion > "6.1"
      Call SetParam("SetupPS2",              strSetupPS2,              "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPS2 <> ""
      ' Nothing
    Case strSQLVersion >= "SQL2008"
      Call SetParam("SetupPS2",              strSetupPS2,              "YES", "PS2 is mandatory for " & strSQLVersion, "")
  End Select

  Select Case True
    Case strOSVersion > "6.0"
      Call SetParam("SetupPS1",              strSetupPS1,              "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPS2 = "YES"
      Call SetParam("SetupPS1",              strSetupPS1,              "No",  "PS1 not required when PS2 being installed", "")
    Case strSQLVersion <= "SQL2008"
      Call SetParam("SetupPS1",              strSetupPS1,              "YES", "PS1 is mandatory for " & strSQLVersion, "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupReportViewer",     strSetupReportViewer,     "N/A", "", strListSQLTools)
    Case Else
      If strSetupRMLTools = "YES" Then
        Call SetParam("SetupReportViewer",   strSetupReportViewer,     "YES", "Report Viewer mandatory for RML Tools", "")
      End If
      If strSetupSQLNexus = "YES" Then
        Call SetParam("SetupReportViewer",   strSetupReportViewer,     "YES", "Report Viewer mandatory for SQL Nexus", "")
      End If
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLRSCluster = "YES"
      Call SetParam("SetupRSAdmin",          strSetupRSAdmin,          "YES", "RS Administration Configuration mandatory for SSRS Cluster", "")
      Call SetParam("SetupRSExec",           strSetupRSExec,           "YES", "RS Report Execution Account mandatory for SSRS Cluster", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupStreamInsight = "YES"
      Call SetParam("SetupNet4",             strSetupNet4,             "YES", ".Net4.0 or above mandatory for StreamInsight", "")
      Call SetParam("SetupSQLCE",            strSetupSQLCE,            "YES", "SQL Compact Edition mandatory for StreamInsight", "")
    Case strSetupSQLCE <> ""
      ' Nothing
    Case Else
      strSetupSQLCE = "NO"
  End Select

  Select Case True    
    Case strSQLVersion > "SQL2005"
      Call SetParam("SetupVS2005SP1",        strSetupVS2005SP1,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupVS2005SP1",        strSetupVS2005SP1,        "N/A", "", strListSQLTools)
    Case strSetupBIDS <> "YES"
      Call SetParam("SetupVS2005SP1",        strSetupVS2005SP1,        "N/A", "Visual Studio 2005 SP1 can not be installed when BIDS is not installed", "")
    Case strSetupVS2005SP1 <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupVS2005SP1",        strSetupVS2005SP1,        "YES", "Visual Studio 2005 SP1 recommended when BIDS is installed", "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "N/A", "", strListSQLVersion)
    Case strSQLVersion > "SQL2014"
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "N/A", "", strListSQLTools)
    Case strSetupSSMS <> "YES"
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "N/A", "", strListSQLVersion)
    Case strSSMSexe <> ""
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "N/A", "", strListSQLVersion)
    Case strSetupVS2010SP1 <> ""
      ' Nothing
    Case Else
      Call SetParam("SetupVS2010SP1",        strSetupVS2010SP1,        "YES", "Visual Studio 2010 SP1 recommended for " & strSQLVersion, "")
  End Select

End Sub


Sub SetupDataDep2()
  Call SetProcessId("0BMD", "Setup Parameter Data for Dependency Level 2")  

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSSISDB = "YES"
      Call SetParam("SetCLREnabled",         strSetCLREnabled,         "1",  "CLR Required for SSIS Catalog DB", "")
  End Select

  Select Case True
    Case strOSVersion <> "6.0"
      Call SetParam("SetupKB932232",         strSetupKB932232,         "N/A", "", strListOSVersion)
    Case strSQLVersion > "SQL2005"
      Call SetParam("SetupKB932232",         strSetupKB932232,         "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupVS2005SP1 <> "YES"
      Call SetParam("SetupKB932232",         strSetupKB932232,         "NO",  "KB932232 can not be installed when Visual Studio 2005 SP1 is not installed", "")
    Case Else
      Call SetParam("SetupKB932232",         strSetupKB932232,         "YES", "KB932232 recommended when Visual Studio 2005 SP1 is installed", "")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2012"
      Call SetParam("SetupKB2549864",        strSetupKB2549864,        "N/A", "", strListSQLVersion)
    Case strSQLVersion > "SQL2014"
      Call SetParam("SetupKB2549864",        strSetupKB2549864,        "N/A", "", strListSQLVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      Call SetParam("SetupKB2549864",        strSetupKB2549864,        "N/A", "", strListSQLTools)
    Case strSetupVS2010SP1 <> "YES" 
      Call SetParam("SetupKB2549864",        strSetupKB2549864,        "NO",  "KB2549864 can not be installed when Visual Studio 2010 SP1 is not installed", "")
    Case strSetupKB2549864 = ""
      Call SetParam("SetupKB2549864",        strSetupKB2549864,        "YES", "KB2549864 recommended if Visual Studio 2010 SP1 installed", "")
   End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPS2 = "N/A"
      ' Nothing
    Case strSetupMBCA = "YES"
      Call SetParam("SetupPS2",              strSetupPS2,              "YES", "Powershell V2 mandatory for Microsoft Baseline Configuration Analyzer", "")
    Case strSQLVersion <= "SQL2012"
      Call SetParam("SetupPS2",              strSetupPS2,              "YES", "Powershell V2 mandatory for " & strSQLVersion, "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupReportViewer = "YES"
      Call SetParam("SetupNet3",             strSetupNet3,             "YES", ".Net3 mandatory for Report Viewer", "")
    Case strSetupReportViewer = ""
      strSetupReportViewer = "NO"
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupRSAlias <> "YES"
      ' Nothing
    Case strRSAlias <> ""
      ' Nothing
    Case strSetupAlwaysOn = "YES"
      strRSAlias    = strGroupAO
      Call SetBuildMessage(strMsgInfo,  "/RSAlias: set to " & strGroupAO)
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/RSAlias: parameter must be specified for /SetupRSAlias:Yes")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupSQLNexus = "YES"
      Call SetParam("SetupNet4",             strSetupNet4,             "YES", ".Net4 mandatory for SQL Nexus", "")
  End Select

End Sub


Sub SetupDataDep3()
  Call SetProcessId("0BME", "Setup Parameter Data for Dependency Level 3")  

  Select Case True
    Case strOSVersion >= "6.1"
      Call SetParam("SetupMSI45",            strSetupMSI45,            "N/A", "", strListOSVersion)
    Case strType = "REBUILD"
      ' Nothing
    Case strSQLVersion >= "SQL2008"
      Call SetParam("SetupMSI45",            strSetupMSI45,            "YES", "Installer v4.5 mandatory for " & strSQLVersion, "")
    Case (strSetupPS2 = "YES") And (strOSVersion >= "6.0")
      Call SetParam("SetupMSI45",            strSetupMSI45,            "YES", "Installer v4.5 mandatory for Powershell", "")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strSetupPS2 = "YES"
      Call SetParam("SetupNet3",              strSetupNet3,            "YES", ".Net3 mandatory for Powershell", "")
  End Select

End Sub


Sub GetMiscData()
  Call SetProcessId("0BN", "Get Miscellaneous data for Buildfile")
  Dim strTF, strExchServer, strLogin, strPassword, strStatusKB933789

  intIdx            = Int((intProcNum / 1.5) + 1)
  Select Case True
    Case strSQLMaxDop = ""
      strSQLMaxDop     = intIdx
    Case strSQLMaxDop = 0
      strSQLMaxDop     = intIdx
  End Select
  If strSQLMaxDop > 8 Then
    strSQLMaxDop       = 8
  End If
  Call ParamListAdd("ListNonDefault", "SQLMaxDop")

  If strSQLTempdbFileCount = "" Then
    strSQLTempdbFileCount = strSQLMaxDop
  End If

  Select Case True
    Case colArgs.Exists("AllowUpgradeForRSSharePointMode")
      strAllowUpgradeForRSSharePointMode = "YES"
    Case Else
      strAllowUpgradeForRSSharePointMode = "NO"
  End Select

  strPath           = "SOFTWARE\Microsoft\Updates\Windows Server 2003\SP3\KB933789\"
  objWMIReg.GetStringValue strHKLM,strPath,"Description",strStatusKB933789
  Select Case True
    Case strOSVersion >= "6.0"
      strCheckRegPerm = "OK"
    Case Instr(Ucase(strOSName), " XP") > 0
      strCheckRegPerm = "OK"
    Case strStatusKB933789 > ""
      strCheckRegPerm = "OK"
    Case Else
      strCheckRegPerm = ""
  End Select

  Select Case True
    Case colArgs.Exists("CLUSTERPASSIVE")
      strClusterPassive = "YES"
    Case Else
      strClusterPassive = "NO"
  End Select

  Select Case True
    Case strOSVersion < "6.0"
      strDefaultUser = "Default User"
    Case Else
      strDefaultUser = "Default"
  End Select 
  strPath           = strProfDir & "\" & strDefaultUser 
  strDfltDoc        = Replace(objShell.RegRead(strUserReg & "Personal"),   "%USERPROFILE%", strPath)
  strDfltProf       = Replace(objShell.RegRead(strUserReg & "Start Menu"), "%USERPROFILE%", strPath)
  strDfltRoot       = Replace(objShell.RegRead(strUserReg & "AppData"),    "%USERPROFILE%", strPath)

  Select Case True
    Case strClusterHost <> "YES"
      ' Nothing
    Case strLabPrefix <> ""
      ' Nothing
    Case Else
      strLabPrefix  = strClusterName
  End Select

  strExchServer     = GetAccountAttr(strSQLAccount, strUserDNSDomain, "msExchHomeServerName")
  strDebugMsg1      = "Mail Server: " & strMailServer
  strDebugMsg2      = "Exch Server: " & strExchServer
  Select Case True
    Case strMailServer <> ""
      ' Nothing
    Case strExchServer = ""
      ' Nothing
    Case Else
      strMailServer     = Mid(strExchServer, InstrRev(strExchServer, "=") + 1)
      strMailServerType = "E"
  End Select

  strIISRoot        = strVolSys & ":\inetpub\wwwroot"

  Select Case True
    Case (strSQLVersion = "SQL2005") And (strWOWX86 = "TRUE")
      strRegSSIS      = "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSDTS\ServiceConfigFile\"
      strRegSSISSetup = "HKLM\SOFTWARE\Wow6432Node\Microsoft\MSDTS\Setup\DTSPath\"
    Case strSQLVersion = "SQL2005" 
      strRegSSIS      = "HKLM\SOFTWARE\Microsoft\MSDTS\ServiceConfigFile\"
      strRegSSISSetup = "HKLM\SOFTWARE\Microsoft\MSDTS\Setup\DTSPath\"
    Case Else
      strRegSSIS      = strHKLMSQL & strSQLVersionNum & "\SSIS\ServiceConfigFile\"
      strRegSSISSetup = strHKLMSQL & strSQLVersionNum & "\SSIS\Setup\DTSPath\"
  End Select

  Select Case True
    Case strRSExecAccount = ""
      ' Nothing
    Case Instr(strRSExecAccount, "\") = 0 
      strRSExecAccount = strDomain & "\" & strRSExecAccount
  End Select
  strRSEmail        = GetParam(Null,                  "RSEmail",            "",                    strRSExecAccount)
  intIdx            = Instr(strRSEmail, "\")
  Select Case True
    Case strRSEmail = ""
      ' Nothing
    Case strUserDNSDomain = ""
      strRSEmail    = ""
    Case intIdx > 0
      strRSEmail    = Mid(strRSEmail, intIdx + 1) & "@" & strUserDNSDomain
    Case Instr(strRSEmail, "@") = 0
      strRSEmail    = strRSEmail & "@" & strUserDNSDomain
  End Select
  Call SetBuildfileValue("RSEmail", strRSEmail)

  Select Case True
    Case strRsShareAccount = ""
      ' Nothing
    Case Instr(strRsShareAccount, "\") = 0 
      strRsShareAccount = strDomain & "\" & strRsShareAccount
  End Select

  Select Case True
    Case strOUPath = ""
      ' Nothing
    Case GetOUAttr(strOUPath, strUserDNSDomain, "distinguishedName") = ""
      Call SetBuildMessage(strMsgInfo,  "OU cannot be found: " & strOUPath)
    Case Else
      strOUCName    =  GetOUAttr(strOUPath, strUserDNSDomain, "distinguishedName")
      Call ParamListAdd("ListNonDefault", "OUPath")
  End Select

  Call GetCmdSQL()

  strSQLEmail       = GetParam(Null,                  "SQLEmail",           "",                    strDBAEmail)
  If strSQLEmail = "" Then
    strSQLEmail     = strSQLAccount
  End If
  intIdx            = Instr(strSQLEmail, "\")
  Select Case True
    Case intIdx > 0
      strSQLEmail   = Mid(strSQLEmail, intIdx + 1) & "@" & strUserDNSDomain
    Case Instr(strSQLEmail, "@") = 0
      strSQLEmail   = strSQLEmail & "@" & strUserDNSDomain
  End Select

  strSQLSupportMsi  = strFileArc & "\Setup\SqlSupport.msi"

  Select Case True
    Case strSQLMinMemory = ""
      strSQLMinMemory  = "0"
    Case Int(strSQLMinMemory) > strSQLMaxMemory
      strSQLMinMemory  = strSQLMaxMemory
      Call ParamListAdd("ListNonDefault", "SQLMinMemory")
  End Select

  Select Case True
    Case strSQLVersion = "SQL2005"
      strSQLSvcStartupType = GetParam(colGlobal,      "SQLAutoStart",       "",                    "1")
      strAGTSvcStartupType = GetParam(colGlobal,      "AGTAutoStart",       "",                    "1")
      strASSvcStartupType  = GetParam(colGlobal,      "ASAutoStart",        "",                    "0")
      strIsSvcStartupType  = GetParam(colGlobal,      "ISAutoStart",        "IsSvcStartupType",    "1")
      strRSSvcStartupType  = GetParam(colGlobal,      "RSAutoStart",        "",                    "0")
      strSqlBrowserStartup = GetParam(colGlobal,      "SQLBrowserAutoStart","",                    "0")
      strWriterSvcStartupType = GetParam(colGlobal,   "SQLWriterAutoStart", "",                    "0")
    Case Else
      strSQLSvcStartupType = UCase(GetParam(colGlobal,"SQLSvcStartupType",  "",                    "Automatic"))
      strAGTSvcStartupType = UCase(GetParam(colGlobal,"AGTSvcStartupType",  "",                    "Automatic"))
      strASSvcStartupType  = Ucase(GetParam(colGlobal,"ASSvcStartupType",   "",                    "Manual"))
      strIsSvcStartupType  = Ucase(GetParam(colGlobal,"IsSvcStartupType",   "",                    "Automatic"))
      strIsMasterStartupType = Ucase(GetParam(colGlobal,"ISMasterSvcStartupType", "",              "Automatic"))
      strIsWorkerStartupType = Ucase(GetParam(colGlobal,"ISWorkerSvcStartupType", "",              "Automatic"))
      strRSSvcStartupType  = UCase(GetParam(colGlobal,"RSSvcStartupType",   "",                    "Automatic"))
      strSqlBrowserStartup = UCase(GetParam(colGlobal,"BrowserSvcStartupType","",                  "Manual"))
      strTelSvcStartup     = UCase(GetParam(colGlobal,"TelSvcSvcStartupType", "",                  "Manual"))
      strWriterSvcStartupType = Ucase(GetParam(colGlobal, "WriterSvcStartupMode", "",              "Manual"))
      If strSQLVersion >= "SQL2016" Then
        strPBEngSvcStartup = UCase(GetParam(colGlobal, "PBEngSvcStartupType","",                   "Automatic"))
        strPBDMSSvcStartup = UCase(GetParam(colGlobal, "PBDMSSvcStartupType","",                   "Automatic"))
      End If
  End Select
  strCtlrStartupType    = UCase(GetParam(colGlobal, "CtlrStartupType",  "",                    "Manual"))
  strCltStartupType     = UCase(GetParam(colGlobal, "CltStartupType",   "",                    "Manual"))

  Select Case true
    Case strClusterAction = ""
      ' Nothing
    Case strSQLVersion <= "SQL2005"
      strSqlBrowserStartup = "1"
    Case Else
      strSqlBrowserStartup = UCase("Automatic")
  End Select

  Select Case True
    Case CInt(strNumLogins) > 99
      Call SetParam("NumLogins",             strNumLogins,             "99",  "Maximum of 99 Logins allowed", "")
  End Select

  For intIdx = 1 To strNumLogins
    strLogin        = GetParam(Null,                  "WinLogin" & Right("0" & CStr(intIdx), 2),     "",  "")
    If strLogin <> "" Then
      Call SetBuildfileValue("WinLogin" & Right("0" & CStr(intIdx), 2), strLogin)
    End If
    strLogin        = GetParam(Null,                  "UserLogin" & Right("0" & CStr(intIdx), 2),    "",  "")
    strPassword     = GetParam(Null,                  "UserPassword" & Right("0" & CStr(intIdx), 2), "",  "")
    If strLogin <> "" Then
      Call SetBuildfileValue("UserLogin" & Right("0" & CStr(intIdx), 2),    strLogin)
      Call SetBuildfileValue("UserPassword" & Right("0" & CStr(intIdx), 2), strPassword)
    End If
  Next

  Select Case True
    Case strNumTF = ""
      strNumTF      = 0
    Case CInt(strNumTF) > 99
      Call SetParam("NumTF",                 strNumTF,                 "99",  "Maximum of 99 Trace Flags allowed", "")
  End Select

  For intIdx = 1 To strNumTF
    strTF           = GetParam(colGlobal,             "TF" & Right("0" & CStr(intIdx), 2),"TF" & CStr(intIdx),  "")
    If strTF <> "" Then
      Call SetBuildfileValue("TF" & Right("0" & CStr(intIdx), 2), strTF)
    End If
  Next

  Select Case True
    Case colArgs.Exists("SKUUpgrade")
      strSKUUpgrade = "YES"
    Case Else
      strSKUUpgrade = "NO"
  End Select

  If Left(strUserConfigurationvbs, 16) = ".\Build Scripts\" Then
    strUserConfigurationvbs = strPathFBScripts & Mid(strUserConfigurationvbs, 17)
  End If
  If Left(strUserPreparationvbs, 16) = ".\Build Scripts\" Then
    strUserPreparationvbs = strPathFBScripts & Mid(strUserPreparationvbs, 17)
  End If

End Sub


Sub SetBuildfileData()
  Call SetProcessId("0C", "Set values for Buildfile")

  Call SetSQLMediaValues()

  Call SetBuildfileValue("DirASDLL",                strDirASDLL)
  Call SetBuildfileValue("DirServInst",             strDirServInst)
  Call SetBuildfileValue("DirSQL",                  strDirSQL)
  Call SetBuildfileValue("DirSQLBootstrap",         strDirSQLBootstrap)
  Call SetBuildfileValue("DirSys",                  strDirSys)
  Call SetBuildfileValue("DirSysData",              strDirSysData)
  Call SetBuildfileValue("DirProg",                 strDirProg)
  Call SetBuildfileValue("DirProgX86",              strDirProgX86)
  Call SetBuildfileValue("DirProgSys",              strDirProgSys)
  Call SetBuildfileValue("DirProgSysX86",           strDirProgSysX86)
  Call SetBuildfileValue("DiscoverFile",            strDiscoverFile)
  Call SetBuildfileValue("DiscoverFolder",          strDiscoverFolder)
  Call SetBuildfileValue("DriveList",               strDriveList)
  Call SetBuildfileValue("LabBackup",               strLabBackup)
  Call SetBuildfileValue("LabBackupAS",             strLabBackupAS)
  Call SetBuildfileValue("LabBPE",                  strLabBPE)
  Call SetBuildfileValue("LabData",                 strLabData)
  Call SetBuildfileValue("LabDataAS",               strLabDataAS)
  Call SetBuildfileValue("LabDataFS",               strLabDataFS)
  Call SetBuildfileValue("LabDataFT",               strLabDataFT)
  Call SetBuildfileValue("LabDBA",                  strLabDBA)
  Call SetBuildfileValue("LabDTC",                  strLabDTC)
  Call SetBuildfileValue("LabLog",                  strLabLog)
  Call SetBuildfileValue("LabLogAS",                strLabLogAS)
  Call SetBuildfileValue("LabLogTemp",              strLabLogTemp)
  Call SetBuildfileValue("LabPrefix",               strLabPrefix)
  Call SetBuildfileValue("LabProg",                 strLabProg)
  Call SetBuildfileValue("LabSysDB",                strLabSysDB)
  Call SetBuildfileValue("LabSystem",               strLabSystem)
  Call SetBuildfileValue("LabTemp",                 strLabTemp)
  Call SetBuildfileValue("LabTempAS",               strLabTempAS)
  Call SetBuildfileValue("LabTempWin",              strLabTempWin)
  Call SetBuildfileValue("OUPath",                  strOUPath)
  Call SetBuildfileValue("OUCName",                 strOUCName)
  Call SetBuildfileValue("PathAddComp",             strPathAddComp)
  Call SetBuildfileValue("PathAddCompOrig",         strPathAddCompOrig)
  Call SetBuildfileValue("PathAutoConfig",          strPathAutoConfig)
  Call SetBuildfileValue("PathAutoConfigOrig",      strPathAutoConfigOrig)
  Call SetBuildfileValue("PathBOL",                 strPathBOL)
  Call SetBuildfileValue("PathFB",                  strPathFB)
  Call SetBuildfileValue("PathFBScripts",           strPathFBScripts)
  Call SetBuildfileValue("PathCScript",             strPathCScript)
  Call SetBuildfileValue("PathPS",                  strPathPS)
  Call SetBuildfileValue("PathSQLSP",               strPathSQLSP)
  Call SetBuildfileValue("PathSQLSPOrig",           strPathSQLSPOrig)
  Call SetBuildfileValue("PathSSIS",                strPathSSIS)
  Call SetBuildfileValue("PathSSMS",                strPathSSMS)
  Call SetBuildfileValue("PathSSMSX86",             strPathSSMSX86)
  Call SetBuildfileValue("PathSys",                 strPathSys)
  Call SetBuildfileValue("PathVS",                  strPathVS)
  Call SetBuildfileValue("SQLProgDir",              strSQLProgDir)
  Call SetBuildfileValue("WinDir",                  strDirSys)

  intTimer          = Timer()

  Call SetBuildfileValue("AgentJobHistory",         strAgentJobHistory)
  Call SetBuildfileValue("AgentMaxHistory",         strAgentMaxHistory)
  Call SetBuildfileValue("AllowUpgradeForRSSharePointMode", strAllowUpgradeForRSSharePointMode)
  Call SetBuildfileValue("AllUserDTop",             strAllUserDTop)
  Call SetBuildfileValue("AllUserProf",             strAllUserProf)
  Call SetBuildfileValue("Alphabet",                strAlphabet)
  Call SetBuildfileValue("AnyKey",                  strAnyKey)
  Call SetBuildfileValue("ASProviderMSOlap",        strASProviderMSOlap)
  Call SetBuildfileValue("ASServerMode",            strAsServerMode)
  Call SetBuildfileValue("BPEFile",                 strBPEFile) 
  Call SetBuildfileValue("CatalogServer",           strCatalogServer)
  Call SetBuildfileValue("CatalogServerName",       strCatalogServerName)
  Call SetBuildfileValue("CatalogInstance",         strCatalogInstance)
  Call SetBuildfileValue("CheckRegPerm",            strCheckRegPerm)
  Call SetBuildfileValue("CollationAS",             strCollationAS)
  Call SetBuildfileValue("CollationSQL",            strCollationSQL)
  Call SetBuildfileValue("CompatFlags",             strCompatFlags)
  Call SetBuildfileValue("ConfirmIPDependencyChange",    strConfirmIPDependencyChange) 
  Call SetBuildfileValue("CSVRoot",                 strCSVRoot)
  Call SetBuildfileValue("DBA_DB",                  strDBA_DB)
  Call SetBuildfileValue("DBAEmail",                strDBAEmail)
  Call SetBuildfileValue("DBMailProfile",           strDBMailProfile)
  Call SetBuildfileValue("DBOwnerAccount",          strDBOwnerAccount)
  Call SetBuildfileValue("DefaultUser",             strDefaultUser)
  Call SetBuildfileValue("DfltDoc" ,                strDfltDoc)
  Call SetBuildfileValue("DfltProf",                strDfltProf)
  Call SetBuildfileValue("DfltRoot",                strDfltRoot)
  Call SetBuildfileValue("DisableNetworkProtocols", strDisableNetworkProtocols)
  Call SetBuildfileValue("Domain",                  strDomain)
  Call SetBuildfileValue("DomainSID",               strDomainSID)
  Call SetBuildfileValue("Debug",                   strDebug)
  Call SetBuildfileValue("EnableRANU",              strEnableRANU)
  Call SetBuildfileValue("Enu",                     strEnu)
  Call SetBuildfileValue("ErrorReporting",          strErrorReporting)
  Call SetBuildfileValue("ExpVersion",              strExpVersion) 
  Call SetBuildfileValue("FBCmd",                   strFBCmd)
  Call SetBuildfileValue("FBParm",                  strFBParm)
  Call SetBuildfileValue("Features",                strFeatures)
  Call SetBuildfileValue("FilePerm",                strFilePerm)
  Call SetBuildfileValue("FirewallStatus",          strFirewallStatus)
  Call SetBuildfileValue("FSLevel",                 strFSLevel)
  Call SetBuildfileValue("FSInstLevel",             strFSInstLevel)
  Call SetBuildfileValue("FSShareName",             strFSShareName)
  Call SetBuildfileValue("GroupAO",                 strGroupAO)
  Call SetBuildfileValue("GroupDBA",                strGroupDBA)
  Call SetBuildfileValue("GroupDBANonSA",           strGroupDBANonSA)
  Call SetBuildfileValue("GroupDistComUsers",       strGroupDistComUsers)
  Call SetBuildfileValue("GroupIISIUsers",          strGroupIISIUsers)
  Call SetBuildfileValue("GroupMSA",                strGroupMSA)
  Call SetBuildfileValue("GroupPerfLogUsers",       strGroupPerfLogUsers)
  Call SetBuildfileValue("GroupPerfMonUsers",       strGroupPerfMonUsers)
  Call SetBuildfileValue("GroupRDUsers",            strGroupRDUsers)
  Call SetBuildfileValue("HKLMFB",                  strHKLMFB)
  Call SetBuildfileValue("HTTP",                    strHTTP)
  Call SetBuildfileValue("InstRegAS",               strInstRegAS)
  Call SetBuildfileValue("InstRegRS",               strInstRegRS)
  Call SetBuildfileValue("InstRegSQL",              strInstRegSQL)
  Call SetBuildfileValue("MailServer",              strMailServer)
  Call SetBuildfileValue("MailServerType",          strMailServerType)
  Call SetBuildfileValue("MainInstance",            strMainInstance)
  Call SetBuildfileValue("ManagementDW",            strManagementDW)
  Call SetBuildfileValue("ManagementServer",        strManagementServer)
  Call SetBuildfileValue("ManagementServerRes",     strManagementServerRes)
  Call SetBuildfileValue("ManagementServerName",    strManagementServerName)
  Call SetBuildfileValue("ManagementInstance",      strManagementInstance)
  Call SetBuildfileValue("SQLMaxDop",               strSQLMaxDop)
  Call SetBuildfileValue("NumTF",                   strNumTF)
  Call SetBuildfileValue("MDSDB",                   strMDSDB)
  Call SetBuildfileValue("MDSPort",                 strMDSPort)
  Call SetBuildfileValue("MDSSite",                 strMDSSite)
  Call SetBuildfileValue("MembersDBA",              strMembersDBA)
  Call SetBuildfileValue("Mode",                    strMode)
  Call SetBuildfileValue("MSSupplied",              strMSSupplied)
  Call SetBuildfileValue("NumErrorLogs",            strNumErrorLogs)
  Call SetBuildfileValue("NumLogins",               strNumLogins)
  Call SetBuildfileValue("Options",                 strOptions)
  Call SetBuildfileValue("OSLanguage",              strOSLanguage)
  Call SetBuildfileValue("OSName",                  strOSName)
  Call SetBuildfileValue("OSLevel",                 strOSLevel)
  Call SetBuildfileValue("OSType",                  strOSType)
  Call SetBuildfileValue("OSVersion",               strOSVersion)
  Call SetBuildfileValue("PID",                     strPID)
  Call SetBuildfileValue("ProcArc",                 strProcArc)
  Call SetBuildfileValue("ProcNum",                 intProcNum)
  Call SetBuildfileValue("ProfileName",             strProfileName)
  Call SetBuildfileValue("ProgCacls",               strProgCacls)
  Call SetBuildfileValue("ProgNTRights",            strProgNTRights)
  Call SetBuildfileValue("ProgSetSPN",              strProgSetSPN)
  Call SetBuildfileValue("ProgReg",                 strProgReg)
  Call SetBuildfileValue("PSInstall",               strPsInstall)
  Call SetBuildfileValue("RebootStatus",            strRebootStatus)
  Call SetBuildfileValue("RegSSIS",                 strRegSSIS)
  Call SetBuildfileValue("RegSSISSetup",            strRegSSISSetup)
  Call SetBuildfileValue("ResponseNo",              strResponseNo)
  Call SetBuildfileValue("ResponseYes",             strResponseYes)
  Call SetBuildfileValue("ReportOnly",              strReportOnly)
  Call SetBuildfileValue("RSInstallMode",           strRSInstallMode)
  Call SetBuildfileValue("RSShpInstallMode",        strRSShpInstallMode)
  Call SetBuildfileValue("RSSQLLocal",              strRSSQLLocal)
  Call SetBuildfileValue("RSVersion",               strRSVersion)

  intTimer          = Timer() - intTimer          ' Get timing data for about 100 items

  Call SetBuildfileValue("Action",                  strAction)
  Call SetBuildfileValue("ActionAO",                strActionAO)
  Call SetBuildfileValue("ActionDAG",               strActionDAG)
  Call SetBuildfileValue("ActionDTC",               strActionDTC)
  Call SetBuildfileValue("ActionSQLDB",             strActionSQLDB)
  Call SetBuildfileValue("ActionSQLAS",             strActionSQLAS)
  Call SetBuildfileValue("ActionSQLIS",             strActionSQLIS)
  Call SetBuildfileValue("ActionSQLRS",             strActionSQLRS)
  Call SetBuildfileValue("ActionSQLTools",          strActionSQLTools)
  Call SetBuildfileValue("ActionClusInst",          strActionClusInst)
  Call SetBuildfileValue("AGDagName",               strAGDagName)
  Call SetBuildfileValue("AGDagNodes",              strAGDagNodes)
  Call SetBuildfileValue("AGName",                  strAGName)
  Call SetBuildfileValue("AOAliasOwner",            strAOAliasOwner)
  Call SetBuildfileValue("AutoLogonCount",          strAutoLogonCount)
  Call SetBuildfileValue("AVCmd",                   strAVCmd)
  Call SetBuildfileValue("CLSIdDTExec",             strCLSIdDTExec)
  Call SetBuildfileValue("CLSIdNetCon",             strCLSIdNetCon)
  Call SetBuildfileValue("CLSIdRunBroker",          strCLSIdRunBroker)
  Call SetBuildfileValue("CLSIdSQL",                strCLSIdSQL)
  Call SetBuildfileValue("CLSIdSQLSetup",           strCLSIdSQLSetup)
  Call SetBuildfileValue("CLSIdSSIS",               strCLSIdSSIS)
  Call SetBuildfileValue("CLSIdVS",                 strCLSIdVS)
  Call SetBuildfileValue("ClusGroups",              strClusGroups)
  Call SetBuildfileValue("ClusStorage",             strClusStorage)
  Call SetBuildfileValue("ClusSubnet",              strClusSubnet)
  Call SetBuildfileValue("ClusterAction",           strClusterAction)
  Call SetBuildfileValue("ClusterGroupAO",          strClusterGroupAO)
  Call SetBuildfileValue("ClusterGroupAS",          strClusterGroupAS)
  Call SetBuildfileValue("ClusterGroupDTC",         strClusterGroupDTC)
  Call SetBuildfileValue("ClusterGroupFS",          strClusterGroupFS)
  Call SetBuildfileValue("ClusterGroupRS",          strClusterGroupRS)
  Call SetBuildfileValue("ClusterGroupSQL",         strClusterGroupSQL)
  Call SetBuildfileValue("ClusterHost",             strClusterHost)
  Call SetBuildfileValue("ClusterNameAS",           strClusterNameAS)
  Call SetBuildfileValue("ClusterNameDTC",          strClusterNameDTC)
  Call SetBuildfileValue("ClusterNameIS",           strClusterNameIS)
  Call SetBuildfileValue("ClusterNamePE",           strClusterNamePE)
  Call SetBuildfileValue("ClusterNamePM",           strClusterNamePM)
  Call SetBuildfileValue("ClusterNameRS",           strClusterNameRS)
  Call SetBuildfileValue("ClusterNameSQL",          strClusterNameSQL)
  Call SetBuildfileValue("ClusterNetworkAS",        strClusterNetworkAS)
  Call SetBuildfileValue("ClusterNetworkDTC",       strClusterNetworkDTC)
  Call SetBuildfileValue("ClusterNetworkSQL",       strClusterNetworkSQL)
  Call SetBuildfileValue("ClusterNode",             strClusterNode)
  Call SetBuildfileValue("ClusIPAddress",           strClusIPAddress)
  Call SetBuildfileValue("ClusIPVersion",           strClusIPVersion)
  Call SetBuildfileValue("ClusIPV4Address",         strClusIPV4Address)
  Call SetBuildfileValue("ClusIPV4Mask",            strClusIPV4Mask)
  Call SetBuildfileValue("ClusIPV4Network",         strClusIPV4Network)
  Call SetBuildfileValue("ClusIPV6Address",         strClusIPV6Address)
  Call SetBuildfileValue("ClusIPV6Mask",            strClusIPV6Mask)
  Call SetBuildfileValue("ClusIPV6Network",         strClusIPV6Network)
  Call SetBuildfileValue("ClusterPassive",          strClusterPassive)
  Call SetBuildfileValue("DNSIPIM",                 strDNSIPIM)
  Call SetBuildfileValue("DNSNameIM",               strDNSNameIM)
  Call SetBuildfileValue("DTCClusterRes",           strDTCClusterRes)
  Call SetBuildfileValue("DTCMultiInstance",        strDTCMultiInstance)
  Call SetBuildfileValue("EditionEnt",              strEditionEnt)
  Call SetBuildfileValue("EdType",                  strEdType)
  Call SetBuildfileValue("EncryptAO",               strEncryptAO)
  Call SetBuildfileValue("FailoverClusterRollOwnership", strFailoverClusterRollOwnership)
  Call SetBuildfileValue("FarmAccount",             strFarmAccount)
  Call SetBuildfileValue("FarmPassword",            strFarmPassword)
  Call SetBuildfileValue("FarmAdminIPort",          strFarmAdminIPort)
  Call SetBuildfileValue("FineBuildStatus",         strFineBuildStatus)
  Call SetBuildfileValue("FTUpgradeOption",         strFTUpgradeOption)
  Call SetBuildfileValue("IsInstallDBA",            strIsInstallDBA)
  Call SetBuildfileValue("InstADHelper",            strInstADHelper)
  Call SetBuildfileValue("InstAgent",               strInstAgent)
  Call SetBuildfileValue("InstAnal",                strInstAnal)
  Call SetBuildfileValue("InstAO",                  strInstAO)
  Call SetBuildfileValue("InstAS",                  strInstAS)
  Call SetBuildfileValue("InstASCon",               strInstASCon)
  Call SetBuildfileValue("InstASSQL",               strInstASSQL)
  Call SetBuildfileValue("InstFT",                  strInstFT)
  Call SetBuildfileValue("InstIS",                  strInstIS)
  Call SetBuildfileValue("InstISMaster",            strInstISMaster)
  Call SetBuildfileValue("InstISWorker",            strInstISWorker)
  Call SetBuildfileValue("InstLog",                 strInstLog)
  Call SetBuildfileValue("InstMR",                  strInstMR)
  Call SetBuildfileValue("InstNode",                strInstNode)
  Call SetBuildfileValue("InstNodeAS",              strInstNodeAS)
  Call SetBuildfileValue("InstNodeIS",              strInstNodeIS)
  Call SetBuildfileValue("InstPE",                  strInstPE)
  Call SetBuildfileValue("InstPM",                  strInstPM)
  Call SetBuildfileValue("InstRS",                  strInstRS)
  Call SetBuildfileValue("InstRSDir",               strInstRSDir)
  Call SetBuildfileValue("InstRSSQL",               strInstRSSQL)
  Call SetBuildfileValue("InstRSHost",              strInstRSHost)
  Call SetBuildfileValue("InstRSURL",               strInstRSURL)
  Call SetBuildfileValue("InstRSWMI",               strInstRSWMI)
  Call SetBuildfileValue("InstSQL",                 strInstSQL)
  Call SetBuildfileValue("InstStream",              strInstStream)
  Call SetBuildfileValue("InstTel",                 strInstTel)
  Call SetBuildfileValue("IISRoot",                 strIISRoot)
  Call SetBuildfileValue("StartJobPassword",        strStartJobPassword)
  Call SetBuildfileValue("MountRoot",               strMountRoot)
  Call SetBuildfileValue("NativeOS",                strNativeOS)
  Call SetBuildfileValue("NetNameSource",           strNetNameSource)
  Call SetBuildfileValue("NetworkGUID",             strNetworkGUID)
  Call SetBuildfileValue("PreferredOwner",          strPreferredOwner)
  Call SetBuildfileValue("RebootLoop",              strRebootLoop)
  Call SetBuildfileValue("ReportViewerVersion",     strReportViewerVersion)
  Call SetBuildfileValue("ResSuffixAS",             strResSuffixAS)
  Call SetBuildfileValue("ResSuffixDB",             strResSuffixDB)
  Call SetBuildfileValue("Role",                    strRole)
  Call SetBuildfileValue("RoleDBANonSA",            strRoleDBANonSA)
  Call SetBuildfileValue("SecDBA",                  strSecDBA)
  Call SetBuildfileValue("SecMain",                 strSecMain)
  Call SetBuildfileValue("SecTemp",                 strSecTemp)
  Call SetBuildfileValue("SecurityMode",            strSecurityMode)
  Call SetBuildfileValue("ServInst",                strServInst)
  Call SetBuildfileValue("Server",                  strServer)
  Call SetBuildfileValue("ServerAO",                strServerAO)
  Call SetBuildfileValue("ServerGroups",            strServerGroups)
  Call SetBuildfileValue("ServerIP",                strServerIP)
  Call SetBuildfileValue("ServerMB",                strServerMB)
  Call SetBuildfileValue("ServName",                strServName)
  Call SetBuildfileValue("SetLowMemLimit",          strSetLowMemLimit)
  Call SetBuildfileValue("SetHardMemLimit",         strSetHardMemLimit)
  Call SetBuildfileValue("SetTotalMemLimit",        strSetTotalMemLimit)
  Call SetBuildfileValue("SetVertiMemLimit",        strSetVertiMemLimit)
  Call SetBuildfileValue("SetCLREnabled",           strSetCLREnabled)
  Call SetBuildfileValue("SetCostThreshold",        strSetCostThreshold)
  Call SetBuildfileValue("SetHeaderLength",         strSetHeaderLength)
  Call SetBuildfileValue("SetMemOptHybridBP",       strSetMemOptHybridBP)
  Call SetBuildfileValue("SetMemOptTempdb",         strSetMemOptTempdb)
  Call SetBuildfileValue("SQLMaxMemory",            strSQLMaxMemory)
  Call SetBuildfileValue("SQLMinMemory",            strSQLMinMemory)
  Call SetBuildfileValue("SetOptimizeForAdHocWorkloads", strSetOptimizeForAdHocWorkloads)
  Call SetBuildfileValue("SetRemoteAdminConnections",    strSetRemoteAdminConnections)
  Call SetBuildfileValue("SetRemoteProcTrans",      strSetRemoteProcTrans)
  Call SetBuildfileValue("SetROLAPDimensionProcessingEffort", strSetROLAPDimensionProcessingEffort)
  Call SetBuildfileValue("SetWorkingSetMaximum",    strSetWorkingSetMaximum)
  Call SetBuildfileValue("SetxpCmdshell",           strSetxpCmdshell)
  Call SetBuildfileValue("SIDDistComUsers",         strSIDDistComUsers)
  Call SetBuildfileValue("SIDIISIUsers",            strSIDIISIUsers)
  Call SetBuildfileValue("SKUUpgrade",              strSKUUpgrade)
  Call SetBuildfileValue("SpeedTest",               intSpeedTest) 
  Call SetBuildfileValue("SPLevel",                 strSPLevel) 
  Call SetBuildfileValue("SPCULevel",               strSPCULevel) 
  Call SetBuildfileValue("SQLAgentStart",           strSQLAgentStart)
  Call SetBuildfileValue("SQLExe",                  strSQLExe)
  Call SetBuildfileValue("SQLLanguage",             strSQLLanguage)
  Call SetBuildfileValue("SQLLogReinit",            strSQLLogReinit)
  Call SetBuildfileValue("SQLRecoveryComplete",     strSQLRecoveryComplete)
  Call SetBuildfileValue("SQLSupportMsi",           strSQLSupportMsi)
  Call SetBuildfileValue("SQLVersionNet",           strSQLVersionNet)
  Call SetBuildfileValue("SQLVersionWMI",           strSQLVersionWMI)
  Call SetBuildfileValue("SQLRSStart",              strSQLRSStart)
  Call SetBuildfileValue("SQLTempdbFileCount",      strSQLTempdbFileCount)
  Call SetBuildfileValue("SQLList",                 strSQLList)
  Call SetBuildfileValue("StatusAssumed",           strStatusAssumed)
  Call SetBuildfileValue("StatusBypassed",          strStatusBypassed)
  Call SetBuildfileValue("StatusComplete",          strStatusComplete)
  Call SetBuildfileValue("StatusFail",              strStatusFail)
  Call SetBuildfileValue("StatusManual",            strStatusManual)
  Call SetBuildfileValue("StatusPreConfig",         strStatusPreConfig)
  Call SetBuildfileValue("StatusProgress",          strStatusProgress)
  Call SetBuildfileValue("StatusKB2919355",         strStatusKB2919355)
  Call SetBuildfileValue("StatusRobocopy",          strStatusRobocopy)
  Call SetBuildfileValue("StatusXcopy",             strStatusXcopy)
  Call SetBuildfileValue("StreamInsightPID",        strStreamInsightPID)
  Call SetBuildfileValue("StopAt",                  strStopAt)

  Call SetBuildfileValue("AdminPassword",           strAdminPassword)
  Call SetBuildfileValue("AgtSvcStartupType",       strAgtSvcStartupType)
  Call SetBuildfileValue("AsSvcStartupType",        strAsSvcStartupType)
  Call SetBuildfileValue("AuditLevel",              strAuditLevel)
  Call SetBuildfileValue("AuditVersion",            strSQLVersion)
  Call SetBuildfileValue("AuditEdition",            strEdition) 
  Call SetBuildfileValue("BackupStart",             strBackupStart)
  Call SetBuildfileValue("BackupRetain",            strBackupRetain)
  Call SetBuildfileValue("BackupDiffRetain",        strBackupDiffRetain)
  Call SetBuildfileValue("BackupLogFreq",           strBackupLogFreq)
  Call SetBuildfileValue("BackupLogRetain",         strBackupLogRetain)
  Call SetBuildfileValue("BuildFileTime",           intTimer)  
  Call SetBuildfileValue("CmdShellAccount",         strCmdShellAccount)
  Call SetBuildfileValue("CmdShellPassword",        strCmdShellPassword)
  Call SetBuildfileValue("SqlBrowserStartup",       strSqlBrowserStartup)
  Call SetBuildfileValue("SqlWriterStartupType",    strWriterSvcStartupType)
  Call SetBuildfileValue("DistDatabase",            strDistDatabase)
  Call SetBuildfileValue("DistPassword",            strDistPassword)
  Call SetBuildfileValue("DQPassword",              strDQPassword)
  Call SetBuildfileValue("CtlrPassword",            strCtlrPassword)
  Call SetBuildfileValue("CtlrStartupType",         strCtlrStartupType)
  Call SetBuildfileValue("CltStartupType",          strCltStartupType)
  Call SetBuildfileValue("HistoryRetain",           strHistoryRetain)
  Call SetBuildfileValue("IsSvcStartupType",        strIsSvcStartupType)
  Call SetBuildfileValue("IsMasterStartupType",     strIsMasterStartupType)
  Call SetBuildfileValue("IsMasterPort",            strIsMasterPort)
  Call SetBuildfileValue("IsMasterThumbprint",      strIsMasterThumbprint)
  Call SetBuildfileValue("IsWorkerStartupType",     strIsWorkerStartupType)
  Call SetBuildfileValue("IsWorkerMaster",          strIsWorkerMaster)
  Call SetBuildfileValue("IsWorkerCert",            strIsWorkerCert)
  Call SetBuildfileValue("JobCategory",             strJobCategory)
  Call SetBuildfileValue("LocalDomain",             strLocalDomain)
  Call SetBuildfileValue("NPEnabled",               strNPEnabled) 
  Call SetBuildfileValue("NTAuthAccount",           strNTAuthAccount)
  Call SetBuildfileValue("NTAuthOSName",            strNTAuthOSName)
  Call SetBuildfileValue("NTService",               strNTService)
  Call SetBuildfileValue("Passphrase",              strPassphrase)
  Call SetBuildfileValue("PBEngSvcStartup",         strPBEngSvcStartup)
  Call SetBuildfileValue("PBDMSSvcStartup",         strPBDMSSvcStartup)
  Call SetBuildfileValue("PBPortRange",             strPBPortRange)
  Call SetBuildfileValue("PBScaleout",              strPBScaleout)
  Call SetBuildfileValue("PowerBIPID",              strPowerBIPID)
  Call SetBuildfileValue("RegasmExe",               strRegasmExe)
  Call SetBuildfileValue("RSAlias",                 strRSAlias)
  Call SetBuildfileValue("RSDBAccount",             strRSDBAccount)
  Call SetBuildfileValue("RSDBPassword",            strRSDBPassword)
  Call SetBuildfileValue("RSDBName",                strRSDBName)
  Call SetBuildfileValue("RSEmail",                 strRSEmail)
  Call SetBuildfileValue("RsFxVersion",             strRsFxVersion)
  Call SetBuildfileValue("RSName",                  strRSName)
  Call SetBuildfileValue("RSFullURL",               strRSFullURL)
  Call SetBuildfileValue("RsSvcStartupType",        strRsSvcStartupType)
  Call SetBuildfileValue("RSURLSuffix",             strRSURLSuffix)
  Call SetBuildfileValue("RSVersionNum",            strRSVersionNum)
  Call SetBuildfileValue("saName",                  strsaName)
  Call SetBuildfileValue("saPwd",                   strsaPwd)
  Call SetBuildfileValue("SQLAdminAccounts",        strSQLAdminAccounts)
  Call SetBuildfileValue("ClusterAOFound",          strClusterAOFound)
  Call SetBuildfileValue("ClusterASFound",          strClusterASFound)
  Call SetBuildfileValue("ClusterDTCFound",         strClusterDTCFound)
  Call SetBuildfileValue("ClusterSQLFound",         strClusterSQLFound)
  Call SetBuildfileValue("SQLJavaDir",              strSQLJavaDir)
  Call SetBuildfileValue("SQLOperator",             strSQLOperator)
  Call SetBuildfileValue("SQLEmail",                strSQLEmail)
  Call SetBuildfileValue("SQLSharedMR",             strSQLSharedMR)
  Call SetBuildfileValue("SqlSvcStartupType",       strSqlSvcStartupType)
  Call SetBuildfileValue("SQMReporting",            strSQMReporting)
  Call SetBuildfileValue("SSASAdminAccounts",       strSSASAdminAccounts)
  Call SetBuildfileValue("SSISDB",                  strSSISDB)
  Call SetBuildfileValue("SSISPassword",            strSSISPassword)
  Call SetBuildfileValue("SSISRetention",           strSSISRetention)
  Call SetBuildfileValue("SSMSexe",                 strSSMSexe)
  Call SetBuildfileValue("TallyCount",              strTallyCount)  
  Call SetBuildfileValue("TCPEnabled",              strTCPEnabled) 
  Call SetBuildfileValue("TCPPort",                 strTCPPort) 
  Call SetBuildfileValue("TCPPortAO",               strTCPPortAO)
  Call SetBuildfileValue("TCPPortAS",               strTCPPortAS)
  Call SetBuildfileValue("TCPPortDAC",              strTCPPortDAC)
  Call SetBuildfileValue("TCPPortDebug",            strTCPPortDebug)
  Call SetBuildfileValue("TCPPortDTC",              strTCPPortDTC)
  Call SetBuildfileValue("TCPPortRS",               strTCPPortRS)
  Call SetBuildfileValue("tempdbFile",              strtempdbFile)
  Call SetBuildfileValue("tempdbLogFile",           strtempdbLogFile)
  Call SetBuildfileValue("Type",                    strType) 
  Call SetBuildfileValue("TypeNode",                strXMLNode)
  Call SetBuildfileValue("UpdateSource",            strUpdateSource)
  Call SetBuildfileValue("UseFreeSSMS",             strUseFreeSSMS)
  Call SetBuildfileValue("UserAccount",             strUserAccount)
  Call SetBuildfileValue("UserConfiguration",       strUserConfiguration)
  Call SetBuildfileValue("UserConfigurationvbs",    strUserConfigurationvbs)
  Call SetBuildfileValue("UserDTop",                strUserDTop)
  Call SetBuildfileValue("UserProf",                strUserProf)
  Call SetBuildfileValue("UserPreparation",         strUserPreparation)
  Call SetBuildfileValue("UserPreparationvbs",      strUserPreparationvbs)
  Call SetBuildfileValue("UseSysDB",                strUseSysDB) 
  Call SetBuildfileValue("VersionNet3",             strVersionNet3) 
  Call SetBuildfileValue("VersionNet4",             strVersionNet4)  
  Call SetBuildfileValue("VolProgX86",              strVolProg)
  Call SetBuildfileValue("VolProgX86Source",        GetBuildfileValue("VolProgSource"))
  Call SetBuildfileValue("VSVersionNum",            strvsVersionNum)  
 
  Call SetBuildfileValue("SetupABE",                strSetupABE)
  Call SetBuildfileValue("SetupAlwaysOn",           strSetupAlwaysOn)
  Call SetBuildfileValue("SetupAOAlias",            strSetupAOAlias)
  Call SetBuildfileValue("SetupAODB",               strSetupAODB)
  Call SetBuildfileValue("SetupAOProcs",            strSetupAOProcs)
  Call SetBuildfileValue("SetupAnalytics",          strSetupAnalytics)
  Call SetBuildfileValue("SetupAPCluster",          strSetupAPCluster)
  Call SetBuildfileValue("SetupAutoConfig",         strSetupAutoConfig)
  Call SetBuildfileValue("SetupBIDS",               strSetupBIDS)
  Call SetBuildfileValue("SetupBIDSHelper",         strSetupBIDSHelper)
  Call SetBuildfileValue("SetupBOL",                strSetupBOL)
  Call SetBuildfileValue("SetupBPAnalyzer",         strSetupBPAnalyzer)
  Call SetBuildfileValue("SetupBPE",                strSetupBPE)
  Call SetBuildfileValue("SetupCacheManager",       strSetupCacheManager)
  Call SetBuildfileValue("SetupClusterShares",      strSetupClusterShares) 
  Call SetBuildfileValue("SetupCMD",                strSetupCMD) 
  Call SetBuildfileValue("SetupCmdshell",           strSetupCmdshell)
  Call SetBuildfileValue("SetupCompliance",         strSetupCompliance)
  Call SetBuildfileValue("SetupDB2OLE",             strSetupDB2OLE)
  Call SetBuildfileValue("SetupDBAManagement",      strSetupDBAManagement)
  Call SetBuildfileValue("SetupDBMail",             strSetupDBMail)
  Call SetBuildfileValue("SetupDBOpts",             strSetupDBOpts)
  Call SetBuildfileValue("SetupDCom",               strSetupDCom)
  Call SetBuildfileValue("SetupDimensionSCD",       strSetupDimensionSCD)
  Call SetBuildfileValue("SetupDisableSA",          strSetupDisableSA)
  Call SetBuildfileValue("SetupDistributor",        strSetupDistributor)
  Call SetBuildfileValue("SetupDQ",                 strSetupDQ)
  Call SetBuildfileValue("SetupDQC",                strSetupDQC)
  Call SetBuildfileValue("SetupDRUCtlr",            strSetupDRUCtlr) 
  Call SetBuildfileValue("SetupDRUClt",             strSetupDRUClt)
  Call SetBuildfileValue("SetupDTCCID",             strSetupDTCCID) 
  Call SetBuildfileValue("SetupDTCCluster",         strSetupDTCCluster) 
  Call SetBuildfileValue("SetupDTCNetAccess",       strSetupDTCNetAccess)
  Call SetBuildfileValue("SetupDTCNetAccessStatus", strSetupDTCNetAccessStatus)
  Call SetBuildfileValue("SetupDTSDesigner",        strSetupDTSDesigner)
  Call SetBuildfileValue("SetupDTSBackup",          strSetupDTSBackup)
  Call SetBuildfileValue("SetupFirewall",           strSetupFirewall) 
  Call SetBuildfileValue("SetupFT",                 strSetupFT)
  Call SetBuildfileValue("SetupGenMaint",           strSetupGenMaint) 
  Call SetBuildfileValue("SetupGovernor",           strSetupGovernor) 
  Call SetBuildfileValue("SetupIIS",                strSetupIIS)
  Call SetBuildfileValue("SetupIntViewer",          strSetupIntViewer)
  Call SetBuildfileValue("SetupISMaster",           strSetupISMaster)
  Call SetBuildfileValue("SetupISMasterCluster",    strSetupISMasterCluster)
  Call SetBuildfileValue("SetupISWorker",           strSetupISWorker)
  Call SetBuildfileValue("SetupJavaDBC",            strSetupJavaDBC)
  Call SetBuildfileValue("SetupStartJob",           strSetupStartJob)
  Call SetBuildfileValue("SetupJRE",                strSetupJRE)
  Call SetBuildfileValue("SetupKB925336",           strSetupKB925336)
  Call SetBuildfileValue("SetupKB932232",           strSetupKB932232)
  Call SetBuildfileValue("SetupKB933789",           strSetupKB933789)
  Call SetBuildfileValue("SetupKB937444",           strSetupKB937444)
  Call SetBuildfileValue("SetupKB954961",           strSetupKB954961)
  Call SetBuildfileValue("SetupKB956250",           strSetupKB956250)
  Call SetBuildfileValue("SetupKB2549864",          strSetupKB2549864)
  Call SetBuildfileValue("SetupKB2781514",          strSetupKB2781514)
  Call SetBuildfileValue("SetupKB2862966",          strSetupKB2862966)
  Call SetBuildfileValue("SetupKB2919355",          strSetupKB2919355)
  Call SetBuildfileValue("SetupKB2919442",          strSetupKB2919442)
  Call SetBuildfileValue("SetupKB3090973",          strSetupKB3090973)
  Call SetBuildfileValue("SetupKB4019990",          strSetupKB4019990)
  Call SetBuildfileValue("SetupManagementDW",       strSetupManagementDW)
  Call SetBuildfileValue("SetupMBCA",               strSetupMBCA)
  Call SetBuildfileValue("SetupMDS",                strSetupMDS)
  Call SetBuildfileValue("SetupMDSC",               strSetupMDSC)
  Call SetBuildfileValue("SetupMDXStudio",          strSetupMDXStudio)
  Call SetBuildfileValue("SetupMenus",              strSetupMenus)
  Call SetBuildfileValue("SetupMyDocs",             strSetupMyDocs) 
  Call SetBuildfileValue("SetupMSI45",              strSetupMSI45) 
  Call SetBuildfileValue("SetupMSMPI",              strSetupMSMPI)
  Call SetBuildfileValue("SetupNet3",               strSetupNet3)
  Call SetBuildfileValue("SetupNet4",               strSetupNet4)
  Call SetBuildfileValue("SetupNet4x",              strSetupNet4x)
  Call SetBuildfileValue("SetupNetBind",            strSetupNetBind)
  Call SetBuildfileValue("SetupNetName",            strSetupNetName)
  Call SetBuildfileValue("SetupNetTrust",           strSetupNetTrust)
  Call SetBuildfileValue("SetupNetwork",            strSetupNetwork) 
  Call SetBuildfileValue("SetupNoDefrag",           strSetupNoDefrag)
  Call SetBuildfileValue("SetupNoDriveIndex",       strSetupNoDriveIndex)
  Call SetBuildfileValue("SetupNoSSL3",             strSetupNoSSL3)
  Call SetBuildfileValue("SetupNoTCPNetBios",       strSetupNoTCPNetBios)
  Call SetBuildfileValue("SetupNoTCPOffload",       strSetupNoTCPOffload)
  Call SetBuildfileValue("SetupNoWinGlobal",        strSetupNoWinGlobal)
  Call SetBuildfileValue("SetupNonSAAccounts",      strSetupNonSAAccounts)
  Call SetBuildfileValue("SetupOLAP",               strSetupOLAP)
  Call SetBuildfileValue("SetupOLAPAPI",            strSetupOLAPAPI)
  Call SetBuildfileValue("SetupOldAccounts",        strSetupOldAccounts) 
  Call SetBuildfileValue("SetupParam",              strSetupParam) 
  Call SetBuildfileValue("SetupPBM",                strSetupPBM)
  Call SetBuildfileValue("SetupPDFReader",          strSetupPDFReader)
  Call SetBuildfileValue("SetupPerfDash",           strSetupPerfDash)
  Call SetBuildfileValue("SetupPolyBase",           strSetupPolyBase)
  Call SetBuildfileValue("SetupPolyBaseCluster",    strSetupPolyBaseCluster)
  Call SetBuildfileValue("SetupPlanExplorer",       strSetupPlanExplorer) 
  Call SetBuildfileValue("SetupPlanExpAddin",       strSetupPlanExpAddin) 
  Call SetBuildfileValue("SetupPowerCfg",           strSetupPowerCfg)
  Call SetBuildfileValue("SetupProcExp",            strSetupProcExp)
  Call SetBuildfileValue("SetupProcMon",            strSetupProcMon)
  Call SetBuildfileValue("SetupPS1",                strSetupPS1)
  Call SetBuildfileValue("SetupPS2",                strSetupPS2)
  Call SetBuildfileValue("SetupPowerBI",            strSetupPowerBI)
  Call SetBuildfileValue("SetupPowerBIDesktop",     strSetupPowerBIDesktop)
  Call SetBuildfileValue("SetupPSRemote",           strSetupPsRemote)
  Call SetBuildfileValue("SetupPython",             strSetupPython)
  Call SetBuildfileValue("SetupRawReader",          strSetupRawReader)
  Call SetBuildfileValue("SetupReportViewer",       strSetupReportViewer)
  Call SetBuildfileValue("SetupRMLTools",           strSetupRMLTools)
  Call SetBuildfileValue("SetupRptTaskPad",         strSetupRptTaskPad)
  Call SetBuildfileValue("SetupRServer",            strSetupRServer)
  Call SetBuildfileValue("SetupRSAdmin",            strSetupRSAdmin)
  Call SetBuildfileValue("SetupRSAlias",            strSetupRSAlias)
  Call SetBuildfileValue("SetupRSDB",               strSetupRSDB)
  Call SetBuildfileValue("SetupRSAT",               strSetupRSAT)
  Call SetBuildfileValue("SetupRSExec",             strSetupRSExec)
  Call SetBuildfileValue("SetupRSIndexes",          strSetupRSIndexes)
  Call SetBuildfileValue("SetupRSKeepAlive",        strSetupRSKeepAlive)
  Call SetBuildfileValue("SetupRSLinkGen",          strSetupRSLinkGen) 
  Call SetBuildfileValue("SetupRSScripter",         strSetupRSScripter)
  Call SetBuildfileValue("SetupRSShare",            strSetupRSShare)
  Call SetBuildfileValue("SetupSAAccounts",         strSetupSAAccounts)
  Call SetBuildfileValue("SetupSAPassword",         strSetupSAPassword)
  Call SetBuildfileValue("SetupSamples",            strSetupSamples)
  Call SetBuildfileValue("SetupSemantics",          strSetupSemantics)
  Call SetBuildfileValue("SetupServices",           strSetupServices)
  Call SetBuildfileValue("SetupServiceRights",      strSetupServiceRights)
  Call SetBuildfileValue("SetupShares",             strSetupShares) 
  Call SetBuildfileValue("SetupSlipstream",         strSetupSlipstream)
  Call SetBuildfileValue("SetupSnapshot",           strSetupSnapshot)
  Call SetBuildfileValue("SetupKerberos",           strSetupKerberos)
  Call SetBuildfileValue("SetupSQLAgent",           strSetupSQLAgent)
  Call SetBuildfileValue("SetupSQLAS",              strSetupSQLAS) 
  Call SetBuildfileValue("SetupSQLASCluster",       strSetupSQLASCluster) 
  Call SetBuildfileValue("SetupSQLBC",              strSetupSQLBC)
  Call SetBuildfileValue("SetupSQLCE",              strSetupSQLCE)
  Call SetBuildfileValue("SetupSQLDB",              strSetupSQLDB)
  Call SetBuildfileValue("SetupSQLDBCluster",       strSetupSQLDBCluster)
  Call SetBuildfileValue("SetupSQLDBAG",            strSetupSQLDBAG)
  Call SetBuildfileValue("SetupSQLDBFS",            strSetupSQLDBFS)
  Call SetBuildfileValue("SetupSQLDBFT",            strSetupSQLDBFT)
  Call SetBuildfileValue("SetupSQLDBRepl",          strSetupSQLDBRepl)
  Call SetBuildfileValue("SetupSQLDebug",           strSetupSQLDebug)
  Call SetBuildfileValue("SetupSQLInst",            strSetupSQLInst)
  Call SetBuildfileValue("SetupSQLIS",              strSetupSQLIS)
  Call SetBuildfileValue("SetupSQLMail",            strSetupSQLMail)
  Call SetBuildfileValue("SetupSQLNexus",           strSetupSQLNexus)
  Call SetBuildfileValue("SetupSQLNS",              strSetupSQLNS)
  Call SetBuildfileValue("SetupSQLPowershell",      strSetupSQLPowershell)
  Call SetBuildfileValue("SetupSQLRS",              strSetupSQLRS)
  Call SetBuildfileValue("SetupSQLRSCluster",       strSetupSQLRSCluster)
  Call SetBuildfileValue("SetupSQLServer",          strSetupSQLServer)
  Call SetBuildfileValue("SetupSQLTools",           strSetupSQLTools)
  Call SetBuildfileValue("SetupSP",                 strSetupSP) 
  Call SetBuildfileValue("SetupSPCU",               strSetupSPCU)
  Call SetBuildfileValue("SetupSPCUSNAC",           strSetupSPCUSNAC)
  Call SetBuildfileValue("SetupSSDTBI",             strSetupSSDTBI)
  Call SetBuildfileValue("SetupSSISCluster",        strSetupSSISCluster)
  Call SetBuildfileValue("SetupSSISDB",             strSetupSSISDB) 
  Call SetBuildfileValue("SetupSSL",                strSetupSSL)
  Call SetBuildfileValue("SetupSSMS",               strSetupSSMS)
  Call SetBuildfileValue("SetupStdAccounts",        strSetupStdAccounts)
  Call SetBuildfileValue("SetupStreamInsight",      strSetupStreamInsight)
  Call SetBuildfileValue("SetupStretch",            strSetupStretch)
  Call SetBuildfileValue("SetupSysDB",              strSetupSysDB)
  Call SetBuildfileValue("SetupSysIndex",           strSetupSysIndex)
  Call SetBuildfileValue("SetupSysManagement",      strSetupSysManagement)
  Call SetBuildfileValue("SetupSystemViews",        strSetupSystemViews)
  Call SetBuildfileValue("SetupTelemetry",          strSetupTelemetry)
  Call SetBuildfileValue("SetupTempDb",             strSetupTempDb)
  Call SetBuildfileValue("SetupTempWin",            strSetupTempWin)
  Call SetBuildfileValue("SetupTLS12",              strSetupTLS12)
  Call SetBuildfileValue("SetupTrouble",            strSetupTrouble)
  Call SetBuildfileValue("SetupVC2010",             strSetupVC2010)
  Call SetBuildfileValue("SetupVS",                 strSetupVS)
  Call SetBuildfileValue("SetupVS2005SP1",          strSetupVS2005SP1)
  Call SetBuildfileValue("SetupVS2010SP1",          strSetupVS2010SP1)
  Call SetBuildfileValue("SetupWinAudit",           strSetupWinAudit)
  Call SetBuildfileValue("SetupWindows",            strSetupWindows) 
  Call SetBuildfileValue("SetupXEvents",            strSetupXEvents)
  Call SetBuildfileValue("SetupXMLNotepad",         strSetupXMLNotepad)
  Call SetBuildfileValue("SetupZoomIt",             strSetupZoomIt) 

End Sub


Sub SetSQLMediaValues()
  Call SetProcessId("0CA", "Set Buildfile values for SQL Media folders")
  Dim strFBPathLocal, strFBPathLocalPrev, strPathSQLMediaPrev

  strFBPathLocal      = GetBuildfileValue("FBPathLocal")
  strFBPathLocalPrev  = GetBuildfileValue("FBPathLocalPrev")
  strPathSQLMediaPrev = GetBuildFileValue("PathSQLMedia")

  Select Case True
    Case strPathSQLMediaPrev = ""
      Call SetBuildfileValue("PathSQLMedia",     strPathSQLMedia)
      Call SetBuildfileValue("PathSQLMediaBase", strPathSQLMedia)
      Call SetBuildfileValue("PathSQLMediaOrig", strPathSQLMediaOrig)
      Call SetBuildfileValue("SQLMediaArc",      strSQLMediaArc)
      Call SetBuildfileValue("PCUSource",        strPCUSource)
      Call SetBuildfileValue("CUSource",         strCUSource)
      Call SetBuildfileValue("GroupDBAAlt",      strGroupDBA)
      Call SetBuildfileValue("GroupDBANonSAAlt", strGroupDBANonSA)
    Case strFBPathLocalPrev = ""
      ' Nothing
    Case strFBPathLocal <> strFBPathLocalPrev
      Call ResetMediaPath("PathSQLMedia",        strFBPathLocal, strFBPathLocalPrev)
      Call ResetMediaPath("PathSQLMediaBase",    strFBPathLocal, strFBPathLocalPrev)
      Call ResetMediaPath("PathSQLMediaOrig",    strFBPathLocal, strFBPathLocalPrev)
      Call ResetMediaPath("SQLMediaArc",         strFBPathLocal, strFBPathLocalPrev)
      Call ResetMediaPath("PCUSource",           strFBPathLocal, strFBPathLocalPrev)
      Call ResetMediaPath("CUSource",            strFBPathLocal, strFBPathLocalPrev)
  End Select

End Sub


Sub ResetMediaPath(strPathVar, strFBPathLocal, strFBPathLocalPrev)
  Call DebugLog("ResetMediaPath: " & strPathVar)
  Dim strPath

  strPath           = GetBuildfileValue(strPathVar)
  If Left(strPath, Len(strFBPathLocalPrev)) = strFBPathLocalPrev Then
    strPath         = strFBPathLocal & Mid(strPath, Len(strFBPathLocalPrev) + 1)
    Call SetBuildfileValue(strPathVar, strPath)
  End If

End Sub


Sub CheckUtils()
  Call SetProcessId("0D", "Check Utility Programs")
  Dim strStatusTemp

  strStatusTemp     = SetupUtil(strProgCacls)
  strStatusTemp     = SetupUtil(strProgNtrights)
  strStatusTemp     = SetupUtil("REG")
  strStatusTemp     = SetupUtil(strProgSetSPN)
  strStatusRobocopy = SetupUtil("ROBOCOPY")
  strStatusXcopy    = SetupUtil("XCOPY")

End Sub


Function SetupUtil(strUtil)
  Call DebugLog("SetupUtil: Check " & strUtil & ".EXE exists in Windows folder")
  Dim strUtilStatus

  strPathNew        = strPathFBScripts & strUtil & ".EXE"
  strCmd            = strUtil & " /?"
  strUtilStatus     = ""
  Call Util_RunExec(strCmd, "EOF", "", -1)
  Select Case True
    Case intErrsave = 0
      strUtilStatus = "OK"
    Case intErrsave = 1
      strUtilStatus = "OK"
    Case intErrSave = 160 ' Deprecated command
      strUtilStatus = "OK"
    Case objFSO.FileExists(strPathSys & strUtil & ".EXE")
      strUtilStatus = "OK"
    Case objFSO.FileExists(strPathNew)
      Call DebugLog("Copy " & strPathNew & " to Windows folder")
      Set objFile   = objFSO.GetFile(strPathNew)
      strDebugMsg1  = "Source: " & objFile.Path
      strDebugMsg2  = "Target: " & strPathSys
      objFile.Copy strPathSys
      strUtilStatus = "OK"
  End Select

  SetupUtil         = strUtilStatus

End Function


Sub FineBuild_Validate()
  Call SetProcessId("0E", "Validation processing for FineBuild for " & strSQLVersion)

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strType = "CLIENT"
      Call Validate_Client()
    Case strType = "CONFIG"
      Call Validate_Config()
    Case strType = "DISCOVER"
      Call Validate_Discover()
    Case strType = "WORKSTATION"
      Call Validate_Workstation()
    Case strType = "FIX"
      Call Validate_Fix()
    Case Else
      Call Validate_Full()
  End Select

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call Validate_Cluster()
  End Select

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case Else
      Call Validate_Common()
  End Select

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case Else
      Call Output_Lists()
  End Select

  Call Validate_License()

  Select Case True
    Case err.Number <> 0
      Call SetBuildMessage(strMsgError, "FineBuild cancelled due to Internal Error")
    Case strValidate = "NO"
      ' Nothing
    Case GetBuildfileValue("ErrorConfig") = "YES"
      Call SetBuildMessage(strMsgError, "FineBuild cancelled due to Validation Errors")
  End Select

End Sub


Sub Validate_Full()
  Call SetProcessId("0EA", "Validation processing for build type FULL")
  Dim strAccountType

  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstSQL
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strService

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case (strService = "") Or (IsNull(strService))
      ' Nothing
    Case strProcessId >= "2B"
      ' Nothing
    Case strType = "FIX"
      ' Nothing
    Case strType = "UPGRADE"
      ' Nothing
    Case strReportOnly = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " already exists")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case strType <> "UPGRADE"
      ' Nothing
    Case strReportOnly = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case (strProcessId < "2C") Or (strProcessId >= "7")
      ' Nothing
    Case strReportOnly = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

  strAccountType    = GetBuildfileValue("SqlAccountType")
  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case InStr("FULL FULLPROG", strType) = 0
      ' Nothing
    Case strEdition = "WORKGROUP"
      ' Nothing
    Case (strSQLVersion <= "SQL2005") And (strAccountType = "L")
      Call SetBuildMessage(strMsgErrorConfig, "/SQLACCOUNT: parameter must be a domain account")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Left(strSqlAccount, Len(strNTAuth)) = strNTAuth
      ' Nothing
    Case Left(strSqlAccount, Len(strNTService)) = strNTService
      ' Nothing
    Case strAccountType = "M"
      Select Case True
        Case strSQLVersion >= "SQL2016"
          ' Nothing
        Case Right(strSqlAccount, 1) = "$" 
          ' Nothing
        Case Else
          Call SetBuildMessage(strMsgErrorConfig, "/SQLSVCACCOUNT: parameter must end with '$'")
      End Select
    Case strSqlPassword > ""
      ' Nothing
    Case strSQLVersion <= "SQL2005" 
      Call SetBuildMessage(strMsgErrorConfig, "/SQLPASSWORD: parameter must be supplied")
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/SQLSVCPASSWORD: parameter must be supplied")
  End Select

  Select Case True
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case (strAccountType = "L") Or (strAccountType = "S")
      Call SetBuildMessage(strMsgErrorConfig, "/SQLSVCACCOUNT: parameter must be a domain account for /SetupAlwaysOn:YES")
  End Select

  Select case True
    Case strSetupCmdshell <> "YES"
      ' Nothing
    Case (strCmdshellAccount > "") And (strCmdshellPassword > "")
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/CmdshellAccount: and /CmdshellPassword: parameters must be supplied")
  End Select

  Select Case True
    Case strType = "REBUILD"
      ' Nothing
    Case strsaPwd = ""
      Call SetBuildMessage(strMsgErrorConfig, "/saPwd: parameter must be supplied")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strActionSQLDB = strActionClusInst
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case strMainInstance = "YES"
      ' Nothing
    Case strTCPPort <> "1433"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/TCPPort: with value other than 1433 must be supplied for a named instance install")
  End Select

End Sub


Sub Validate_Client()
  Call SetProcessId("0EB", "Validation processing for build type CLIENT")

  If Instr(strOSType, "CORE") > 0 Then
    Call SetBuildMessage(strMsgErrorConfig, "CLIENT install not supported on Core OS")
  End If

End Sub


Sub Validate_Config()
  Call SetProcessId("0EC", "Validation processing for build type CONFIG")

  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstSQL
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strService

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case strReportOnly = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

End Sub


Sub Validate_Discover()
  Call SetProcessId("0ED", "Validation processing for build type DISCOVER")

  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstSQL
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strService

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case strReportOnly = "YES"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

End Sub


Sub Validate_Fix()
  Call SetProcessId("0EE", "Validation processing for build type FIX")

' No validation at present

End Sub


Sub Validate_Workstation()
  Call SetProcessId("0EF", "Validation processing for build type WORKSTATION")
  Dim strAccountType

  strPath           = "SYSTEM\CurrentControlSet\Services\" & strInstSQL
  objWMIReg.GetStringValue strHKLM,strPath,"DisplayName",strService

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case (strService = "") Or (IsNull(strService))
      ' Nothing
    Case strProcessId >= "2B"
      ' Nothing
    Case strType = "FIX"
      ' Nothing
    Case strType = "UPGRADE"
      ' Nothing
    Case strReportOnly = "YES" 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " already exists")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case strType <> "UPGRADE"
      ' Nothing
    Case strReportOnly = "YES" 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strService > ""
      ' Nothing
    Case (strProcessId < "2C") Or (strProcessId >= "7")
      ' Nothing
    Case strReportOnly = "YES" 
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "FineBuild cancelled - Requested SQL instance " & strInstSQL & " does not exist")
  End Select

  Select Case True
    Case strSetupManagementDW <> "YES"
      ' Nothing
    Case (strMDWAccount = "") And (strMSSupplied = "Y")
      Call SetBuildMessage(strMsgErrorConfig, "/MDWACCOUNT: and /MDWPASSWORD: parameters must be supplied")
    Case strMDWAccount = ""
      ' Nothing
    Case Instr(strMDWAccount, "\") = 0
      Call SetBuildMessage(strMsgErrorConfig, "/MDWACCOUNT: parameter must be a domain account")
    Case strMDWAccount = strAgtAccount
      ' Nothing
    Case strMDWPassword <> ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/MDWPASSWORD: parameter must be supplied")
  End Select

  If strsaPwd = "" Then
    Call SetBuildMessage(strMsgErrorConfig, "/saPwd: parameter must be supplied")
  End If

  strAccountType    = GetBuildfileValue("SqlAccountType")
  Select Case True
    Case strSetupAlwaysOn <> "YES"
      ' Nothing
    Case (strAccountType = "L") Or (strAccountType = "S")
      Call SetBuildMessage(strMsgErrorConfig, "/SQLSVCACCOUNT: parameter must be a domain account for /SetupAlwaysOn:YES")
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case strMainInstance = "YES"
      ' Nothing
    Case strTCPPort <> "1433"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/TCPPort: with value other than 1433 must be supplied for a named instance install")
  End Select

End Sub


Sub Validate_Cluster()
  Call SetProcessId("0EG", "Validation processing for Cluster installs")

  Select Case True
    Case strClusterTCP = "IPV4"
      ' Nothing
    Case (strClusterTCP = "IPV6") And (strSQLVersion <= "SQL2005")
      Call SetBuildMessage(strMsgErrorConfig, "/ClusterTCP:IPV6 not valid for " & strSQLVersion)
    Case strClusterTCP = "IPV6"
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Invalid /ClusterTCP: value " & strClusterTCP)
  End Select

  Select Case True
    Case strClusterAction = "ADDNODE"
      ' Nothing
    Case strOSVersion < "6.0"                ' Installing on W2003 or below
      ' Nothing
    Case Instr(strOSType, "CORE") > 0        ' Installing on Core OS
      ' Nothing
    Case strClusterReport = ""
      Call SetBuildMessage(strMsgErrorConfig, "Cluster Validation Report can not be found.  Validate the Cluster to produce the report")
    Case CheckReport(strClusterPath, "Testing has completed successfully and the configuration is suitable for clustering")
      ' Nothing
    Case CheckReport(strClusterPath, "Testing has completed successfully. The configuration appears to be suitable for clustering")
      ' Nothing
    Case (strOSVersion >= "6.3A") And CheckReport(strClusterPath, "Testing has completed for the tests you selected")
      ' Nothing
    Case Else
      strDebugMsg1  = "Cluster Report: " & strClusterPath
      Call SetBuildMessage(strMsgErrorConfig, "Cluster must have clean Validation Report for all tests")
  End Select

  Select Case True
    Case strSQLVersion > "SQL2005"
      ' Nothing
    Case strAdminPassword <> ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/AdminPassword: must be supplied and contain the password for the account running SQL FineBuild")
  End Select

  Select Case True
    Case strSetupSQLASCluster <> "YES" 
      ' Nothing
    Case Len(strClusterNameAS) > 15
      Call SetBuildMessage(strMsgErrorConfig, "SQL AS Cluster name """ & strClusterNameAS & """ is longer than 15 characters")
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case (strProcessId < "2B") And (CheckClusterExists(strClusterNameAS) <> "")
      Call SetBuildMessage(strMsgErrorConfig, "SSAS Cluster name """ & strClusterNameAS & """ already exists")
    Case strReportOnly = "YES"
      ' Nothing
    Case (strProcessId > "2CA") And (strProcessId < "7") And (CheckClusterExists(strClusterNameAS) = "")
      Call SetBuildMessage(strMsgErrorConfig, "SSAS Cluster name """ & strClusterNameAS & """ does not exist")
  End Select

  Select Case True
    Case strSetupSQLDBCluster <> "YES" 
      ' Nothing
    Case Len(strClusterNameSQL) > 15
      Call SetBuildMessage(strMsgErrorConfig, "SQL DB Cluster name """ & strClusterNameSQL & """ is longer than 15 characters")
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case (strProcessId > "2ACZ") And (strProcessId < "7") And (CheckClusterExists(strClusterNameSQL) = "") And (strClusterGroupDTC = strClusterGroupSQL)
      Call SetBuildMessage(strMsgErrorConfig, "SQL DB Cluster name """ & strClusterNameSQL & """ does not exist")
    Case (strProcessId < "2B") And (CheckClusterExists(strClusterNameSQL) <> "") And (strClusterGroupDTC <> strClusterGroupSQL)
      Call SetBuildMessage(strMsgErrorConfig, "SQL DB Cluster name """ & strClusterNameSQL & """ already exists")
    Case strReportOnly = "YES"
      ' Nothing
    Case (strProcessId > "2CA") And (strProcessId < "7") And (CheckClusterExists(strClusterNameSQL) = "")
      Call SetBuildMessage(strMsgErrorConfig, "SQL DB Cluster name """ & strClusterNameSQL & """ does not exist")
  End Select

  Select Case True
    Case strSetupSSISCluster <> "YES" 
      ' Nothing
    Case Len(strClusterNameIS) > 15
      Call SetBuildMessage(strMsgErrorConfig, "SQL IS Cluster name """ & strClusterNameIS & """ is longer than 15 characters")
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case (strProcessId < "2B") And (CheckChildClusterExists(strClusterNameIS) <> "")
      Call SetBuildMessage(strMsgErrorConfig, "SSIS Cluster name """ & strClusterNameIS & """ already exists")
    Case strReportOnly = "YES"
      ' Nothing
    Case (strProcessId > "2CCZ") And (strProcessId < "7") And (CheckChildClusterExists(strClusterNameIS) = "")
      Call SetBuildMessage(strMsgErrorConfig, "SSIS Cluster name """ & strClusterNameIS & """ does not exist")
  End Select

  Select Case True
    Case strSetupSQLRSCluster <> "YES" 
      ' Nothing
    Case Len(strClusterNameRS) > 15
      Call SetBuildMessage(strMsgErrorConfig, "RS Cluster name """ & strClusterNameRS & """ is longer than 15 characters")
    Case strActionSQLRS = "ADDNODE"
      ' Nothing
    Case (strProcessId < "4RAQ") And (CheckClusterExists(strClusterNameRS) <> "")
      Call SetBuildMessage(strMsgErrorConfig, "RS Cluster name """ & strClusterNameRS & """ already exists")
    Case strReportOnly = "YES"
      ' Nothing
    Case (strProcessId > "4RAZ") And (strProcessId < "7") And (CheckClusterExists(strClusterNameRS) = "")
      Call SetBuildMessage(strMsgErrorConfig, "RS Cluster name """ & strClusterNameRS & """ does not exist")
  End Select

  Select Case True
    Case strActionDTC = "ADDNODE"
      ' Nothing
    Case strSetupDTCCluster <> "YES" 
      ' Nothing
    Case Len(strClusterNameDTC) > 15
      Call SetBuildMessage(strMsgErrorConfig, "MSDTC Cluster name """ & strClusterNameDTC & """ is longer than 15 characters")
    Case GetBuildfileValue("VolDTCSource") = "C"
      Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDTC:" & strVolDTC & ", MSDTC can not be installed to a CSV")
    Case GetBuildfileValue("VolDTCSource") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDTC:" & strVolDTC & ", MSDTC can not be installed to a network share")
    Case GetBuildfileValue("VolDTCType") <> "C"
      Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDTC:" & strVolDTC & ", MSDTC must be installed to a cluster volume")
  End Select

  Select Case True
    Case strSQLVersion >= "SQL2014"
      ' Nothing
    Case strCSVFound = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " does not support use of CSV storage in a SQL Cluster")
  End Select

End Sub


Function CheckReport(strReport, strText)
  Call DebugLog("CheckReport: " & strReport)
  Dim bFound

  bFound            = False
  strCmd            = "%COMSPEC% /D /C FIND /C """ & strText & """ """ & strReport & """"
  Call Util_RunExec(strCmd, "", "", 1)
  If intErrSave = 0 Then
    bFound          = True
  End If

  CheckReport       = bFound

End Function


Function CheckClusterExists(strClusterName)
  Call DebugLog("CheckClusterExists:" & strClusterName)
  Dim strClusAddr

  CheckClusterExists = ""
  strClusAddr        = GetAddress(strClusterName, "", "N")

  If strClusAddr <> "" Then
    CheckClusterExists = "Y"
  End If

End Function


Function CheckChildClusterExists(strClusterName)
  Call DebugLog("CheckChildClusterExists:" & strClusterName)
  Dim colResources
  Dim objResource

  CheckChildClusterExists = ""
  Set colResources = GetClusterResources()
  For Each objResource In colResources
    Select Case True
      Case objResource.Name = strClusterName
        CheckChildClusterExists = "Y"
        Exit Function
    End Select
  Next

End Function


Sub Validate_Common()
  Call SetProcessId("0EH", "Validation processing for all build types")

  Select Case True
    Case Instr("X86 AMD64", strProcArc) > 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Unsupported processor architecture " & strProcArc)
  End Select

  Select Case True
    Case Instr(strSQLList, strSQLVersion) > 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Unsupported SQL Server version: " & strSQLVersion)
  End Select

  Select Case True
    Case Instr(" CONFIG DISCOVER FIX REBUILD ", strType) > 0
      ' Nothing
    Case strPathSQLMedia = ""
      Call SetBuildMessage(strMsgErrorConfig, "Folder for /PATHSQLMEDIA: for Edition " & strEdition & " can not be found: " & strPathSQLMediaOrig)
    Case (strSQLVersion <> "SQL2008") And (strSQLVersion <> "SQL2008R2")
      ' Nothing
    Case strSetupSlipstream <> "DONE"
      ' Nothing
    Case Instr(strPathSQLMedia, " ") = 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "SQL Media folder name must not contain spaces")
  End Select

  If (strOSVersion < "6.1") And (Instr(strOSType, "CORE") > 0) Then
    Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
  End If
 
  Select Case True
    Case strSQLVersion <> "SQL2005"
      ' Nothing
    Case Left(strOSVersion, 1) < "5" ' NT4 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
    Case (strOSVersion = "5.0") And (strOSLevel < "Service Pack 4") ' W2000
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 4 is needed.")
    Case (strOSVersion = "5.1") And (strOSLevel < "Service Pack 2") ' XP
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 2 is needed.")
    Case (strOSVersion < "6.0") And (strOSLevel < "Service Pack 1") ' W2003
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 1 is needed.")
    Case strOSVersion >= "6.2" ' Windows 2012
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " can not be installed on this operating system")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2008"
      ' Nothing
    Case strOSVersion < "5.1" ' W2000 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
    Case (strOSVersion < "6.0") And (strOSLevel < "Service Pack 2") ' W2003 and below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 2 is needed.")
    Case strOSVersion >= "6.3A" ' Windows 2016, Windows 10
      Call SetBuildMessage(strMsgWarning, strSQLVersion & " install not supported on this operating system")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2008R2"
      ' Nothing
    Case strOSVersion < "5.1" ' W2000 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
    Case (strOSVersion = "5.1") And (strOSLevel < "Service Pack 3") ' XP
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 3 is needed.")
    Case (strOSVersion < "6.1") And (strOSLevel < "Service Pack 2") ' W2008 RTM, Vista, W2003
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 2 is needed.")
    Case strOSVersion >= "6.3A" ' Windows 2016, Windows 10
      Call SetBuildMessage(strMsgWarning, strSQLVersion & " install not supported on this operating system")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2012"
      ' Nothing
    Case Left(strOSVersion, 1) <= "5" ' W2003 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
    Case (strOSVersion < "6.1") And (strOSLevel < "Service Pack 2") ' W2008 RTM or Vista
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 2 is needed.")
    Case (strOSVersion = "6.1") And (strOSLevel < "Service Pack 1") ' W2008 R2 RTM 
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 1 is needed.")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2014"
      ' Nothing
    Case Left(strOSVersion, 1) <= "5" ' W2003 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
    Case (strOSVersion < "6.1") And (strOSLevel < "Service Pack 2") ' W2008 RTM or Vista
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 2 is needed.")
    Case (strOSVersion = "6.1") And (strOSLevel < "Service Pack 1") ' W2008 R2 RTM 
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system.  Service Pack 1 is needed.")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2016"
      ' Nothing
    Case (strOSVersion <= "6.1") ' W2008 R2 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
  End Select

  Select Case True
    Case strSQLVersion <> "SQL2017"
      ' Nothing
    Case (strOSVersion <= "6.1") ' W2008 R2 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2019"
      ' Nothing
    Case (strOSVersion <= "6.3") ' W2012 R2 or below
      Call SetBuildMessage(strMsgErrorConfig, strSQLVersion & " install not supported on this operating system")
  End Select

  Select Case True
    Case Instr(" CON PRN AUX NUL COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9 LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 LPT8 and LPT9 ", " " & strInstance & " ") = 0
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, strInstance & " is a Windows reserved word")
      ' See https://blogs.msdn.com/b/sqlserverfaq/archive/2011/11/02/sql-server-service-pack-installation-may-fail-if-your-instance-name-is-a-windows-reserved-word.aspx
  End Select

  If strUserName = strServer Then
    Call SetBuildMessage(strMsgErrorConfig, "User name and Server name must not be the same")
  End If

  Select Case True
    Case strSQLProgDir = ""
      Call SetBuildMessage(strMsgErrorConfig, "/SQLProgDir: parameter must be supplied")
    Case Instr(strSQLProgDir, "\") > 0
      Call SetBuildMessage(strMsgErrorConfig, "/SQLProgDir: value must not contain a '\': " & strSQLProgDir)
  End Select

  Select Case True
    Case strProcessId > "1TZ"
      ' Nothing
    Case strFirewallstatus <> "1"
      Call SetBuildMessage(strMsgWarning, "Firewall is OFF.  Best practice recommends that Firewall is set to ON")
  End Select

  Select Case True
    Case Instr(" INSTALL UPGRADE ADDNODE INSTALLFAILOVERCLUSTER " , " " & strAction & " ") > 0
      ' Nothing
    Case Instr(" PREPAREIMAGE COMPLETEIMAGE REPAIR REBUILDDATABASE UNINSTALL PREPAREFAILOVERCLUSTER COMPLETEFAILOVERCLUSTER REMOVENODE " , " " & strAction & " ") > 0
      Call SetBuildMessage(strMsgErrorConfig, "Value for /Action: is not supported " & strAction)
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /Action: " & strAction)
  End Select

  Select Case True
    Case strSetupAnalytics <> "YES"
      ' Nothing
    Case strExtSvcAccount = ""
      Call SetBuildMessage(strMsgErrorConfig, "/EXTSVCACCOUNT: parameter must be supplied")
    Case strSetupSQLDBCluster <> "YES"
      ' Nothing 
    Case GetBuildfileValue("ExtSvcAccountType") = "L"
      Call SetBuildMessage(strMsgErrorConfig, "/EXTSVCACCOUNT: parameter must be a domain account")
    Case GetBuildfileValue("ExtSvcAccountType") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "/EXTSVCACCOUNT: parameter must be a domain account")
  End Select

  Select Case True
    Case strSetupDRUCtlr <> "YES"
      ' Nothing
    Case strSQLVersion <= "SQL2016"
      ' Nothing
    Case strCtlrAccount = ""
      Call SetBuildMessage(strMsgErrorConfig, "/CTLRSVCACCOUNT: parameter must be supplied")
    Case GetBuildfileValue("CtlrSvcAccountType") = "L"
      Call SetBuildMessage(strMsgErrorConfig, "/CTLRSVCACCOUNT: parameter must be a domain account")
    Case GetBuildfileValue("CtlrSvcAccountType") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "/CTLRSVCACCOUNT: parameter must be a domain account")
  End Select

  Select Case True
    Case strSetupDRUClt <> "YES"
      ' Nothing
    Case strSQLVersion <= "SQL2016"
      ' Nothing
    Case strCltAccount = ""
      Call SetBuildMessage(strMsgErrorConfig, "/CLTSVCACCOUNT: parameter must be supplied")
    Case GetBuildfileValue("CltSvcAccountType") = "L"
      Call SetBuildMessage(strMsgErrorConfig, "/CLTSVCACCOUNT: parameter must be a domain account")
    Case GetBuildfileValue("CltSvcAccountType") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "/CLTSVCACCOUNT: parameter must be a domain account")
  End Select

  Select Case True
    Case strSetupDQ <> "YES"
      ' Nothing
    Case Len(strDQPassword) < 8  
      Call SetBuildMessage(strMsgErrorConfig,  "/DQPassword: value must be at least 8 characters")
  End Select

  Select Case True
    Case strSetupISMaster <> "YES"
      ' Nothing
    Case strIsMasterAccount = ""
      Call SetBuildMessage(strMsgErrorConfig, "/ISMasterSvcAccount: parameter must be supplied")
    Case GetBuildfileValue("IsMasterSvcAccountType") = "L"
      Call SetBuildMessage(strMsgErrorConfig, "/ISMasterSvcAccount: parameter must be a domain account")
    Case GetBuildfileValue("IsMasterSvcAccountType") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "/ISMasterSvcAccount: parameter must be a domain account")
  End Select

  Select Case True
    Case strSetupISWorker <> "YES"
      ' Nothing
    Case strIsWorkerAccount = ""
      Call SetBuildMessage(strMsgErrorConfig, "/ISWorkerSvcAccount: parameter must be supplied")
    Case Else
      ' Nothing
  End Select

  Select Case True
    Case strSetupManagementDW <> "YES"
      ' Nothing
    Case (strMDWAccount = "") And (strMSSupplied = "Y")
      Call SetBuildMessage(strMsgErrorConfig, "/MDWACCOUNT: and /MDWPASSWORD: parameters must be supplied")
    Case strMDWAccount = ""
      ' Nothing
    Case Instr(strMDWAccount, "\") = 0
      Call SetBuildMessage(strMsgErrorConfig, "/MDWACCOUNT: parameter must be a domain account")
    Case strMDWAccount = strAgtAccount
      ' Nothing
    Case strMDWPassword <> ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgErrorConfig, "/MDWPASSWORD: parameter must be supplied")
  End Select

  Select Case True
    Case strSQLVersion < "SQL2016"
      ' Nothing
    Case strSetupPolyBase <> "YES"
      ' Nothing
    Case strPBEngSvcAccount = ""
      Call SetParam("SetupPolyBase",         strSetupPolyBase,         "NO",     "PolyBase Service Accounts not supplied", "")
    Case strPBEngSvcAccount <> strPBDMSSvcAccount
      Call SetBuildMessage(strMsgErrorConfig, "/PBEngSvcAccount: and /PBDMSSvcAccount: must be the same")
    Case GetBuildfileValue("PBEngSvcAccount" & "Type") = "L"
      Call SetBuildMessage(strMsgErrorConfig, "/PBEngSvcAccount: parameter must be a domain account")
    Case GetBuildfileValue("PBEngSvcAccount" & "Type") = "S"
      Call SetBuildMessage(strMsgErrorConfig, "/PBEngSvcAccount: parameter must be a domain account")
  End Select
  If strSetupPolyBase = "YES" Then
    Call SetParam("TCPEnabled",              strTCPEnabled,            "1",     "TCP is mandatory for PolyBase", "")
  End If

  Select Case True
    Case strSetupRSExec <> "YES"
      ' Nothing
    Case Else
      If GetBuildfileValue("RsExecAccount") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/RsExecAccount: parameter must be supplied")
      End If
      If GetBuildfileValue("RsExecPassword") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/RsExecPassword: parameter must be supplied")
      End If
      If GetBuildfileValue("RSEmail") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/RSEmail: parameter must be supplied")
      End If
  End Select

  Select Case True
    Case strSetupRSShare <> "YES"
      ' Nothing
    Case Else
      If GetBuildfileValue("RsShareAccount") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/RsShareAccount: parameter must be supplied")
      End If
      If GetBuildfileValue("RsSharePassword") = "" Then
        Call SetBuildMessage(strMsgErrorConfig, "/RsSharePassword: parameter must be supplied")
      End If
  End Select

  Select Case True
    Case strSetupCompliance <> "YES"
      ' Nothing
    Case strGroupDBA = ""
      Call SetBuildMessage(strMsgErrorConfig, "/GroupDBA: must be specified for /SetupCompliance:YES")
    Case strGroupDBANonSA = ""
      Call SetBuildMessage(strMsgErrorConfig, "/GroupDBANonSA: must be specified for /SetupCompliance:YES")
    Case strGroupDBANonSA = strGroupDBA
      Call SetBuildMessage(strMsgErrorConfig, "/GroupDBANonSA must be different to /GroupDBA for /SetupCompliance:YES")
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      If GetBuildfileValue("VolDataASSource") = "C" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDataAS:" & strVolDataAS & ", Analysis Services can not be installed to a CSV")
      End If
      If GetBuildfileValue("VolLogASSource") = "C" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolLogAS:" & strVolLogAS & ", Analysis Services can not be installed to a CSV")
      End If
      If GetBuildfileValue("VolTempASSource") = "C" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolTempAS:" & strVolTempAS & ", Analysis Services can not be installed to a CSV")
      End If
      If GetBuildfileValue("VolDataASSource") = "S" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDataAS:" & strVolDataAS & ", Analysis Services can not be installed to a network share")
      End If
      If GetBuildfileValue("VolLogASSource") = "S" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolLogAS:" & strVolLogAS & ", Analysis Services can not be installed to a network share")
      End If
      If GetBuildfileValue("VolTempASSource") = "S" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolTempAS:" & strVolTempAS & ", Analysis Services can not be installed to a network share")
      End If
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strFSLevel = "0"
      ' Nothing
    Case Else
      If GetBuildfileValue("VolDataFSSource") = "S" Then
        Call SetBuildMessage(strMsgErrorConfig, "Invalid value for /VolDataFS:, Filestream can not be installed to a network share")
      End If
  End Select

End Sub


Sub Output_Lists()
  Call SetProcessId("0EI", "Output Parameter lists")

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListType = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to " & strType & " build")
      Call FormatList("ListType", strListType)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListOSVersion = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to OS " & strOSName)
      Call FormatList("ListOSVersion", strListOSVersion)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListCore = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A on Server Core OS")
      Call FormatList("ListCore", strListCore)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListAddNode = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /Action:ADDNODE")
      Call FormatList("ListAddNode", strListAddNode)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSQLVersion = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to SQL Version " & strSQLVersion)
      Call FormatList("ListSQLVersion", strListSQLVersion)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListEdition = ""
      ' Nothing
    Case strEdType <> ""
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to Edition " & strEdition & " (" & strEdType & ")")
      Call FormatList("ListEdition", strListEdition)
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to Edition " & strEdition)
      Call FormatList("ListEdition", strListEdition)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListCompliance = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to YES due to /SetupCompliance: value " & strSetupCompliance)
      Call FormatList("ListCompliance", strListCompliance)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSQLTools = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /SetupSQLTools:" & strSetupSQLTools)
      Call FormatList("ListSQLTools", strListSQLTools)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListCluster = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to not a SQL Cluster install")
      Call FormatList("ListCluster", strListCluster)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSQLDB = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /SetupSQLDB:" & strSetupSQLDB)
      Call FormatList("ListSQLDB", strListSQLDB)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSSAS = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /SetupSQLAS:" & strSetupSQLAS)
      Call FormatList("ListSSAS", strListSSAS)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSQLRS = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /SetupSQLRS:" & strSetupSQLRS)
      Call FormatList("ListSQLRS", strListSQLRS)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListSSIS = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to N/A due to /SetupSQLIS:" & strSetupSQLIS)
      Call FormatList("ListSSIS", strListSSIS)
  End Select

  Select Case True
    Case (strProcessId > "1") And (strProcessId < "7")
      ' Nothing
    Case strListMain = ""
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgInfo, "The following Parameters set to NO due to /MainInstance:" & strMainInstance)
      Call FormatList("ListMain", strListMain)
  End Select

End Sub


Sub FormatList(strListType, strList)
  Call DebugLog("FormatList: " & strListType)
  Dim strWorkList, strLine

  strWorkList       = RTrim(LTrim(strList))
  While LTrim(strWorkList) > ""
    strLine         = RTrim(Left(strWorkList, 76))
    Select Case True
      Case Len(strLine) = 76
        strLine     = Left(strLine, InstrRev(strLine, " "))
        strWorkList = LTrim(Mid(strWorkList & " ", Len(strLine)))
      Case Else
        strWorkList = " "
    End Select
    Call SetBuildMessage(strMsgInfo, "  " & strLine)
  WEnd

End Sub


Sub Validate_License()
  Call SetProcessId("0EJ", "Validate License details")
  Dim strMessage

  Call FBLog(" ")
  Select Case True
    Case colArgs.Exists("IAcceptLicenseTerms") 
      strMessage    = "/IAcceptLicenseTerms is present"
      Call FBLog(strMessage)
      Call SetBuildfileValue("LicenseMsg1", strMessage)
      strMessage    = "  This means that you accept the license terms of FineBuild and the licence terms of all products that FineBuild installs"
      Call FBLog(strMessage)
      Call SetBuildfileValue("LicenseMsg2", strMessage)
      Call SetBuildfileValue("IAcceptLicenseTerms", "YES")
    Case Else
      strMessage    = strMsgErrorConfig & ": /IAcceptLicenseTerms is not found"
      Call FBLog(strMessage)
      Call SetBuildfileValue("LicenseMsg1", strMessage)
      strMessage    = strMsgErrorConfig & ":  /IAcceptLicenseTerms must be given to show you accept the license terms of FineBuild and of all products that FineBuild installs"
      Call FBLog(strMessage)
      Call SetBuildfileValue("LicenseMsg2", strMessage)
      Call SetBuildfileValue("IAcceptLicenseTerms", "NO")
  End Select
  Call FBLog(" ")

End Sub


Function GetGroupAction(strGroupName)
  Call DebugLog("GetGroupAction:" & strGroupName)

  Select Case True
    Case strGroupName = strClusterGroupSQL
      GetGroupAction = strActionSQLDB
    Case strGroupName = strClusterGroupAS
      GetGroupAction = strActionSQLAS
    Case strGroupName = strClusterGroupDTC
      GetGroupAction = strActionDTC
    Case Else
      GetGroupAction = strAction
  End Select

End Function


Function Include(strFile)
  Dim objFile
  Dim strFilePath, strFileText

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      err.Raise 8, "", "ERROR: This process must be run by SQLFineBuild.bat"
    Case Else
      strFilePath       = strPathFB & "Build Scripts\" & strFile
      Set objFile       = objFSO.OpenTextFile(strFilePath)
      strFileText       = objFile.ReadAll()
      objFile.Close 
      ExecuteGlobal strFileText
  End Select

End Function


End Class