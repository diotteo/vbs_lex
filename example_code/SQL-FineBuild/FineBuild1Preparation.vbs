''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild1Preparation.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'  Code to clear IndexingEnabled flag adapted from "Windows Server Cookbook" by Robbie Allen, ISBN 0-596-00633-0
'
'  Purpose:      Builds directory structure and shares for use in a standard
'                SQL Server build as defined in the FineBuild Reference document.
'
'  Author:       Ed Vassie, based on work for SQL 2000 by Mark Allison
'
'  Date:         December 2007
'
'  Change History
'  Version  Author        Date         Description
'  2.3      Ed Vassie     18 Jun 2010  Initial SQL Server R2 version
'  2.2.2    Ed Vassie     29 Oct 2009  Added extra drives to support clustering
'  2.2.1    Ed vassie     28 Jun 2009  Added support for Express Edition
'  2.2      Ed Vassie      8 Oct 2008  Major rewrite for FineBuild V2.0.0
'  2.1      Ed Vassie     20 Feb 2008  Bypass create of SQLAS files for Workgroup Edition
'                                      Add configure of local groups
'  2.0      Ed Vassie     02 Feb 2008  Initial version for FineBuild v1.0.0
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim arrProfFolders, arrUserSid
Dim colSysEnvVars, colUsrEnvVars, colVol
Dim intErrSave, intIdx
Dim objADOCmd, objADOConn, objDrive, objFile, objFolder, objFSO, objFW, objFWRules, objOS, objNetwork, objShell, objVol, objWMI, objWMIReg
Dim strAccountGUID, strAccountDelegate, strAction, strActionDTC, strActionSQLAS, strActionSQLDB, strAgtAccount, strAlphabet, strAnyKey, strAsAccount, strAVCmd
Dim strClusIPVersion, strClusIPV4Network, strClusIPV6Network, strClusStorage, strClusterHost, strClusterAction, strClusterName, strClusterNameAS, strClusterNameRS, strClusterNameSQL, strClusterNode, strClusterRoot, strCmd, strCmdshellAccount, strCSVRoot
Dim strDomain, strDomainSID, strDriveList, strDirDBA, strDirProg, strDirProgX86, strDirSQL, strDirSys, strDirSystemDataBackup, strDirSystemDataShared, strDirProgSys, strDirProgSysX86, strCtlrAccount, strCltAccount
Dim strEdition, strExtSvcAccount, strFirewallStatus, strFolderName, strFSLevel, strFSShareName, strGroupAdmin, strGroupAO, strGroupDBA, strGroupDBANonSA, strGroupDistComUsers, strGroupIISIUsers, strGroupMSA, strGroupPerfLogUsers, strGroupPerfMonUsers, strGroupRDUsers, strGroupUsers, strHKLMFB, strHKU, strHTTP
Dim strFTAccount, strIsAccount, strIsMasterPort, strLocalAdmin, strLocalDomain, strNTAuthAccount, strNTService, strRsAccount, strRSAlias, strSqlAccount
Dim strInstance, strInstAgent, strInstASSQL, strInstLog, strInstNode, strInstNodeAS, strInstNodeIS, strInstRS, strInstRSURL, strInstSQL, strIsMasterAccount, strIsWorkerAccount
Dim strKerberosFile, strMenuSSMS, strNetworkGUID, strOSName, strOSType, strOSVersion, strOUCName, strOUPath
Dim strPath, strPathFB, strPathFBScripts, strPathNew, strPathTemp, strPBPortRange, strPrepareFolderPath, strProcArc, strProfDir, strProgCacls, strProgNtrights, strProgSetSPN, strProgReg, strReboot
Dim strRSInstallMode, strSchedLevel, strSecDBA, strSecMain, strSecNull, strSecTemp, strServer, strServerSID, strServInst, strSetupAlwaysOn, strSetupNetBind, strSetupNetName, strSetupNoDefrag, strSetupNoDriveIndex, strSetupNoTCPNetBios, strSetupNoTCPOffload
Dim strSetupPowerBI, strSetupPowerCfg, strSetupPolyBase, strSetupKerberos, strSetupSQLAS, strSetupSQLASCluster, strSetupSQLDB, strSetupSQLDBCluster, strSetupSQLDBAG, strSetupSQLDBFS, strSetupSQLDBFT, strSetupSQLIS, strSetupSQLRS, strSetupSQLRSCluster, strSetupSQLTools
Dim strSetupWinAudit, strSetupBPE, strSetupCmdshell, strSetupISMaster, strSetupTempWin, strSetupNoWinGlobal, strSetupDRUClt, strDTCClusterRes, strDTCMultiInstance, strSetupDTCCluster, strSetupDTCClusterStatus, strSetupDTCNetAccess, strSetupFirewall, strSetupSP, strSetupSSISCluster
Dim strSetupShares, strSIDDistComUsers, strSIDIISIUsers, strSPLevel, strSQLVersion, strSQLVersionNum, strSQLVersionNet, strSQLVersionWMI, strSqlBrowserAccount
Dim strTCPPort, strTCPPortDTC, strTCPPortRS, strTCPPortISMaster, strType, strUserAccount, strUserName, strUserDNSDomain
Dim strLabBackup, strLabBackupAS, strLabBPE, strLabData, strLabDataAS, strLabDataFS, strLabDataFT, strLabDTC, strLabLog, strLabLogAS, strLabLogTemp, strLabPrefix, strLabProg, strLabSysDB, strLabSystem, strLabTemp, strLabTempAS, strLabTempWin, strLabDBA
Dim strSpace
Dim strVersionFB, strVol, strVolType, strVolUsed, strVolBackup, strVolBackupAS, strVolBPE, strVolData, strVolDataAS, strVolDataFS, strVolDataFT, strVolDTC, strVolLog, strVolLogAS, strVolSysDB, strVolLogTemp, strVolTemp, strVolTempAS, strVolTempWin, strVolProg, strVolSys, strVolDBA
Dim strWaitLong, strWaitShort

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "1TZ"
      ' Nothing
    Case Else
      Call PreparationTasks()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1UZ" ' 1UA to 1UZ reserved for User Preparation processing
      ' Nothing
    Case Else
      Call UserPreparation()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination
  
  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "1TZ"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("SQL Server install preparation complete", 2, "Preparation processing" ,64)
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
      Call FBLog(" SQL Server install preparation failed")
  End Select

  Wscript.Quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation procesing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBUtils.vbs"
  Include "FBManageAccount.vbs"
  Include "FBManageCluster.vbs"
  Include "FBManageInstall.vbs"
  Call SetProcessIdCode("FB1P")

  Set objADOConn    = CreateObject("ADODB.Connection")
  Set objADOCmd     = CreateObject("ADODB.Command")
  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objFW         = CreateObject("HNetCfg.FwPolicy2")
  Set objFWRules    = objFW.Rules
  Set objNetwork    = CreateObject("Wscript.Network")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colSysEnvVars = objShell.Environment("System")
  Set colUsrEnvVars = objShell.Environment("User")

  objADOConn.Provider            = "ADsDSOObject"
  objADOConn.Open "ADs Provider"
  Set objADOCmd.ActiveConnection = objADOConn

  strHKLMFB         = GetBuildfileValue("HKLMFB")
  strHKU            = &H80000003
  strSpace          = Space(20)
  strAction         = GetBuildfileValue("Action")
  strActionDTC      = GetBuildfileValue("ActionDTC")
  strActionSQLAS    = GetBuildfileValue("ActionSQLAS")
  strActionSQLDB    = GetBuildfileValue("ActionSQLDB")
  strAgtAccount     = GetBuildfileValue("AgtAccount")
  strAlphabet       = GetBuildfileValue("Alphabet")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strAsAccount      = GetBuildfileValue("AsAccount")
  strAVCmd          = GetBuildfileValue("AVCmd")
  strSqlBrowserAccount = GetBuildfileValue("SqlBrowserAccount")
  strClusterHost    = GetBuildfileValue("ClusterHost")
  strClusterAction  = GetBuildfileValue("ClusterAction")
  strClusterName    = GetBuildfileValue("ClusterName")
  strClusterNameAS  = GetBuildfileValue("ClusterNameAS")
  strClusterNameRS  = GetBuildfileValue("ClusterNameRS")
  strClusterNameSQL = GetBuildfileValue("ClusterNameSQL")
  strClusterNode    = GetBuildfileValue("ClusterNode")
  strClusterRoot    = GetBuildfileValue("ClusterRoot")
  strClusIPVersion  = GetBuildfileValue("ClusIPVersion")
  strClusIPV4Network  = GetBuildfileValue("ClusIPV4Network")
  strClusIPV6Network  = GetBuildfileValue("ClusIPV6Network")
  strClusStorage    = GetBuildfileValue("ClusStorage")
  strCmdshellAccount  = GetBuildfileValue("CmdshellAccount")
  strCSVRoot        = GetBuildfileValue("CSVRoot")
  strDirDBA         = GetBuildfileValue("DirDBA")
  strDirProg        = GetBuildfileValue("DirProg")
  strDirProgX86     = GetBuildfileValue("DirProgX86")
  strDirProgSys     = GetBuildfileValue("DirProgSys")
  strDirProgSysX86  = GetBuildfileValue("DirProgSysX86")
  strDirSQL         = GetBuildfileValue("DirSQL")
  strDirSys         = GetBuildfileValue("DirSys")
  strDirSystemDataBackup = GetBuildfileValue("DirSystemDataBackup")
  strDirSystemDataShared = GetBuildfileValue("DirSystemDataShared")
  strDomain         = GetBuildfileValue("Domain")
  strDomainSID      = GetBuildfileValue("DomainSID")
  strDriveList      = GetBuildfileValue("DriveList")
  strCtlrAccount    = GetBuildfileValue("CtlrAccount")
  strCltAccount     = GetBuildfileValue("CltAccount")
  strPath           = Mid(strHKLMFB, 6)
  objWMIReg.GetStringValue strHKLM,strPath,"DTCClusterRes",strDTCClusterRes
  strDTCMultiInstance = GetBuildfileValue("DTCMultiInstance")
  strEdition        = GetBuildfileValue("AuditEdition")
  strExtSvcAccount  = GetBuildfileValue("ExtSvcAccount")
  strFirewallStatus = GetBuildfileValue("FirewallStatus")
  strFSShareName    = GetBuildfileValue("FSShareName")
  strFSLevel        = GetBuildfileValue("FSLevel")
  strFTAccount      = GetBuildfileValue("FtAccount")
  strGroupAdmin     = GetBuildfileValue("GroupAdmin")
  strGroupAO        = GetBuildfileValue("GroupAO")
  strGroupDBA       = GetBuildfileValue("GroupDBA")
  strGroupDBANonSA  = GetBuildfileValue("GroupDBANonSA")
  strGroupDistComUsers = GetBuildfileValue("GroupDistComUsers")
  strGroupIISIUsers = GetBuildfileValue("GroupIISIUsers")
  strGroupMSA       = GetBuildfileValue("GroupMSA")
  strGroupPerfLogUsers = GetBuildfileValue("GroupPerfLogUsers")
  strGroupPerfMonUsers = GetBuildfileValue("GroupPerfMonUsers")
  strGroupRDUsers   = GetBuildfileValue("GroupRDUsers")
  strGroupUsers     = GetBuildfileValue("GroupUsers")
  strHTTP           = GetBuildfileValue("HTTP")
  strInstance       = GetBuildfileValue("Instance")
  strInstASSQL      = GetBuildfileValue("InstASSQL")
  strInstLog        = GetBuildfileValue("InstLog")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstNodeAS     = GetBuildfileValue("InstNodeAS")
  strInstNodeIS     = GetBuildfileValue("InstNodeIS")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strIsAccount      = GetBuildfileValue("IsAccount")
  strIsMasterAccount  = GetBuildfileValue("IsMasterAccount")
  strIsMasterPort   = GetBuildfileValue("IsMasterPort")
  strIsWorkerAccount  = GetBuildfileValue("IsWorkerAccount")
  strLabBackup      = GetBuildfileValue("LabBackup")
  strLabBackupAS    = GetBuildfileValue("LabBackupAS")
  strLabBPE         = GetBuildfileValue("LabBPE")
  strLabData        = GetBuildfileValue("LabData")
  strLabDataAS      = GetBuildfileValue("LabDataAS")
  strLabDataFS      = GetBuildfileValue("LabDataFS")
  strLabDataFT      = GetBuildfileValue("LabDataFT")
  strLabDBA         = GetBuildfileValue("LabDBA")
  strLabDTC         = GetBuildfileValue("LabDTC")
  strLabLog         = GetBuildfileValue("LabLog")
  strLabLogAS       = GetBuildfileValue("LabLogAS")
  strLabLogTemp     = GetBuildfileValue("LabLogTemp")
  strLabSysDB       = GetBuildfileValue("LabSysDB")
  strLabPrefix      = GetBuildfileValue("LabPrefix")
  strLabProg        = GetBuildfileValue("LabProg")
  strLabSystem      = GetBuildfileValue("LabSystem")
  strLabTemp        = GetBuildfileValue("LabTemp")
  strLabTempAS      = GetBuildfileValue("LabTempAS")
  strLabTempWin     = GetBuildfileValue("LabTempWin")
  strLocalAdmin     = GetBuildfileValue("LocalAdmin")
  strLocalDomain    = GetBuildfileValue("LocalDomain")
  strNetworkGUID    = GetBuildfileValue("NetworkGUID")
  strNTAuthAccount  = GetBuildfileValue("NTAuthAccount")
  strNTService      = GetBuildfileValue("NTService")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strOUCName        = GetBuildfileValue("OUCName")
  strOUPath         = GetBuildfileValue("OUPath")
  strPathFBScripts  = FormatFolder("PathFBScripts")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPBPortRange    = GetBuildfileValue("PBPortRange")
  strPrepareFolderPath = ""
  strProcArc        = GetBuildfileValue("ProcArc")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strProgNtrights   = GetBuildfileValue("ProgNTRights")
  strProgSetSPN     = GetBuildfileValue("ProgSetSPN")
  strProgReg        = GetBuildfileValue("ProgReg")
  strReboot         = GetBuildfileValue("RebootStatus")
  strRsAccount      = GetBuildfileValue("RsAccount")
  strRSAlias        = GetBuildfileValue("RSAlias")
  strRSInstallMode  = GetBuildfileValue("RSInstallMode")
  strSchedLevel     = GetBuildfileValue("SchedLevel")
  strSecDBA         = GetBuildfileValue("SecDBA")
  strSecMain        = GetBuildfileValue("SecMain")
  strSecNull        = ""
  strSecTemp        = GetBuildfileValue("SecTemp")
  strServerSID      = GetBuildfileValue("ServerSID")
  strServInst       = GetBuildfileValue("ServInst")
  strSetupAlwaysOn  = GetBuildfileValue("SetupAlwaysOn")
  strSetupBPE       = GetBuildfileValue("SetupBPE")
  strSetupCmdshell  = GetBuildfileValue("SetupCmdshell")
  strSetupDRUClt    = GetBuildfileValue("SetupDRUClt")
  strSetupDTCCluster   = GetBuildfileValue("SetupDTCCluster")
  strSetupDTCNetAccess = GetBuildfileValue("SetupDTCNetAccess")
  strSetupFirewall  = GetBuildfileValue("SetupFirewall")
  strSetupISMaster  = GetBuildfileValue("SetupISMaster")
  strSetupNetBind   = GetBuildfileValue("SetupNetBind")
  strSetupNetName   = GetBuildfileValue("SetupNetName")
  strSetupNoDefrag  = GetBuildfileValue("SetupNoDefrag")
  strSetupNoDriveIndex = GetBuildfileValue("SetupNoDriveIndex")
  strSetupNoTCPNetBios = GetBuildfileValue("SetupNoTCPNetBios")
  strSetupNoTCPOffload = GetBuildfileValue("SetupNoTCPOffload")
  strSetupNoWinGlobal  = GetBuildfileValue("SetupNoWinGlobal")
  strSetupPolyBase  = GetBuildfileValue("SetupPolyBase")
  strSetupPowerBI   = GetBuildfileValue("SetupPowerBI")
  strSetupPowerCfg  = GetBuildfileValue("SetupPowerCfg")
  strSetupShares    = GetBuildfileValue("SetupShares")
  strSetupKerberos  = GetBuildfileValue("SetupKerberos")
  strSetupSQLASCluster = GetBuildfileValue("SetupSQLASCluster")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLDBCluster = GetBuildfileValue("SetupSQLDBCluster")
  strSetupSQLDBAG   = GetBuildfileValue("SetupSQLDBAG")
  strSetupSQLDBFS   = GetBuildfileValue("SetupSQLDBFS")
  strSetupSQLDBFT   = GetBuildfileValue("SetupSQLDBFT")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLRSCluster = GetBuildfileValue("SetupSQLRSCluster")
  strSetupSP        = GetBuildfileValue("SetupSP")
  strSetupSSISCluster = GetBuildfileValue("SetupSSISCluster")
  strSetupTempWin   = GetBuildfileValue("SetupTempWin")
  strSetupWinAudit  = GetBuildfileValue("SetupWinAudit")
  strServer         = GetBuildfileValue("AuditServer")
  strSPLevel        = GetBuildfileValue("SPLevel")
  strSqlAccount     = GetBuildfileValue("SqlAccount")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNet  = GetBuildfileValue("SQLVersionNet")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strSQLVersionWMI  = GetBuildfileValue("SQLVersionWMI")
  strTCPPort        = GetBuildfileValue("TCPPort")
  strTCPPortDTC     = GetBuildfileValue("TCPPortDTC")
  strTCPPortRS      = GetBuildfileValue("TCPPortRS")
  strType           = GetBuildfileValue("Type")
  strUserAccount    = GetBuildfileValue("UserAccount")
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  strUserName       = GetBuildfileValue("AuditUser")
  strVersionFB      = objShell.ExpandEnvironmentStrings("%SQLFBVERSION%")
  strVolProg        = GetBuildfileValue("VolProg")
  strVolBackup      = GetBuildfileValue("VolBackup")
  strVolBackupAS    = GetBuildfileValue("VolBackupAS")
  strVolData        = GetBuildfileValue("VolData")
  strVolDataAS      = GetBuildfileValue("VolDataAS")
  strVolDataFS      = GetBuildfileValue("VolDataFS")
  strVolDataFT      = GetBuildfileValue("VolDataFT")
  strVolDBA         = GetBuildfileValue("VolDBA")
  strVolDTC         = GetBuildfileValue("VolDTC")
  strVolLog         = GetBuildfileValue("VolLog")
  strVolLogAS       = GetBuildfileValue("VolLogAS")
  strVolLogTemp     = GetBuildfileValue("VolLogTemp")
  strVolSys         = GetBuildfileValue("VolSys")
  strVolSysDB       = GetBuildfileValue("VolSysDB")
  strVolTemp        = GetBuildfileValue("VolTemp")
  strVolTempAS      = GetBuildfileValue("VolTempAS")
  strVolBPE         = GetBuildfileValue("VolBPE")
  strVolTempWin     = GetBuildfileValue("VolTempWin")
  strVolUsed        = ""
  strWaitLong       = GetBuildfileValue("WaitLong")
  strWaitShort      = GetBuildfileValue("WaitShort")
  Set arrProfFolders  = objFSO.GetFolder(strProfDir).SubFolders

End Sub


Sub PreparationTasks()
  Call SetProcessId("1", strSQLVersion & " Preparation processing (FineBuild1Preparation.vbs)")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case Else
      Call SetupFineBuild()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1BZ"
      ' Nothing
    Case Else
      Call SetupServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CZ"
      ' Nothing
    Case Else
      Call SetupWindows()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DZ"
      ' Nothing
    Case Else
      Call SetupNetwork()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EZ"
      ' Nothing
    Case Else
      Call SetupAccounts()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FZ"
      ' Nothing
    Case Else
      Call SetupVolumes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GZ"
      ' Nothing
    Case Else
      Call SetupFolders()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HZ"
      ' Nothing
    Case Else
      Call PostPreparation()
  End Select

  Call SetProcessId("1TZ", " Preparation processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupFineBuild()
  Call SetProcessId("1A", "Setup FineBuild")

  ' No actions at present

  Call SetProcessId("1AZ", " Setup FineBuild" & strStatusComplete)
  Call ProcessEnd("")

End Sub 


Sub SetupServer()
  Call SetProcessId("1B", "Setup Server")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1BA"
      ' Nothing
    Case Else
      Call SetupServerName()
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub 


Sub SetupServerName()
  Call SetProcessId("1BA", "Setup Server Name")
  Dim colHostname
  Dim objHostname
  Dim strHostname

  strHostname       = objShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Hostname")
  If StrComp(strHostname, UCase(strHostname), vbBinaryCompare) <> 0 Then
    Call DebugLog("Change Hostname to upper case")
    Set colHostname = objWMI.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each objHostname In colHostname
      Call DebugLog("Set server name to upper case " & objHostname.Name)
      Call objHostname.Rename(Ucase(objHostname.Name))
    Next
    Call SetBuildfileValue("RebootStatus", "Pending") 
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWindows()
  Call SetProcessId("1C", "Setup Windows")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CA"
      ' Nothing
    Case Else
      Call SetupServiceTimeout()
  End Select

  ' ProcessId 1CB available for reuse

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CC"
      ' Nothing
    Case strSetupPowerCfg <> "YES"
      ' Nothing
    Case Else
      Call SetupPowerCfg()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CD"
      ' Nothing
    Case strSetupNoDefrag <> "YES"
      ' Nothing
    Case Else
      Call SetupNoDefrag()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1CE"
      ' Nothing
    Case strSetupWinAudit <> "YES"
      ' Nothing
    Case strOSVersion < "6.0"
      Call SetBuildfileValue("SetupWinAuditStatus", strStatusManual)
    Case Else
      Call SetupWinAudit()
  End Select

  Call SetProcessId("1CZ", " Setup Windows" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupServiceTimeout()
  Call SetProcessId("1CA", "Setup Service Timeout")
  Dim intTime, intSpeedTest

  intSpeedTest      = GetBuildfileValue("SpeedTest")
  intTime           = GetBuildfileValue("BuildFileTime") 
  Select Case True 
    Case CDbl(intTime) <= CDbl(intSpeedTest)
      ' Nothing
    Case Else 
      intTime       = Cstr((Int(intTime) + 1) * 10000) ' Increase service startup time 1/10 second for every second of Buildfile time.
      Call SetServicePipeTimeOut(intTime, "Slow system detected, service start time allowance increased.")
  End Select

  Select Case True
    Case strSetupSQLDBCluster = "YES"
      Call SetServicePipeTimeOut("600000", "Service start time allowance increased to 10 minutes for SQL DB Cluster") 
    Case strSetupPolybase = "YES"
      Call SetServicePipeTimeOut("120000", "Service start time allowance increased to 2 minutes for PolyBase") 
    Case strSetupSQLRS = "YES"  
      Call SetServicePipeTimeOut("60000", "Service start time allowance increased to 1 minute for Reporting Services")
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetServicePipeTimeOut(intTime, strMsg)
  Call DebugLog("SetServicePipeTimeOut:")
  Dim intTimeout

  strPath           = "SYSTEM\CurrentControlSet\Control"
  objWMIReg.GetDwordValue strHKLM, strPath, "ServicesPipeTimeout", intTimeout
  If IsNull(intTimeout) Then
    intTimeout      = 30000
  End If

  If CLng(intTimeout) < CLng(intTime) Then
    strPath         = "HKLM\" & strPath & "\ServicesPipeTimeout"
    Call SetBuildMessage(strMsgInfo, strMsg)
    Call DebugLog("Adjusting " & strPath & " from " & Cstr(intTimeout) & " to " & CStr(intTime) & " milliseconds")
    Call Util_RegWrite(strPath, intTime, "REG_DWORD") 
    Call SetBuildfileValue("RebootStatus", "Pending")   ' Reboot needed so new Timeout can take effect 
  End If

End Sub


Sub SetupCertificateLog()
  Call DebugLog("SetupCertificateLog:")
' Described in KB2661254, SQL Self-Signs with 1024 bit Certificates and this change allows them to be accepted by Windows

  strPath           = strDirSys & "\Logs\CertLog"
  If Not objFSO.FolderExists(strPath) Then
    objFSO.CreateFolder(strPath)
    WScript.Sleep strWaitShort
  End If
  strCmd            = "CERTUTIL -SETREG chain\WeakSignatureLogDir """ & strPath & "\Under1024Key.Log"""
  Call Util_RunExec(strCmd, "", strResponseYes, 0)
  strCmd            = "CERTUTIL -SETREG chain\EnableWeakSignatureFlags 8"
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

End Sub


Sub SetupPowerCfg()
  Call SetProcessId("1CC", "Setup Windows Power Configuration")
  Dim strPowerScheme

  strPath           = "SOFTWARE\Policies\Microsoft\Power\PowerSettings"
  objWMIReg.GetStringValue strHKLM,strPath,"ActivePowerScheme",strPowerScheme

  Select Case True
    Case strPowerScheme = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
      Call SetBuildfileValue("SetupPowerCfgStatus", strStatusComplete)
    Case Else
      strPowerScheme = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
      strPath        = "HKLM\" & strPath & "\ActivePowerScheme"
      Call Util_RegWrite(strPath, strPowerScheme, "REG_SZ") 
      Call SetBuildfileValue("SetupPowerCfgStatus", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoDefrag()
  Call SetProcessId("1CD", "Setup No Disk Defragmentation")

  strCmd            = "SCHTASKS /Change /tn ""Microsoft/Windows/Defrag/ScheduledDefrag"" /DISABLE"
  Call Util_RunExec(strCmd, "", strResponseYes, -1)

  Call SetBuildfileValue("SetupNoDefragStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupWinAudit()
  Call SetProcessId("1CE", "Setup Windows Audit")

  Call Util_RunExec("AUDITPOL /set /Category:""Account Logon""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Account Management"" /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""DS Access""          /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Logon/Logoff""       /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Object Access""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Policy Change""      /success:enable",                   "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Privilege Use""      /success:enable  /failure:enable",  "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""Detailed Tracking""  /success:disable /failure:disable", "", "", 0) 
  Call Util_RunExec("AUDITPOL /set /Category:""System""             /success:enable",                   "", "", 0) 

  Call SetBuildfileValue("SetupWinAuditStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNetwork()
  Call SetProcessId("1D", "Setup SQL Server Network")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DA"
      ' Nothing
    Case strSetupFirewall <> "YES"
      ' Nothing
    Case Else
      Call SetupFireWall()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DB"
      ' Nothing
    Case strSetupNetName <> "YES"
      ' Nothing
    Case strClusterHost <> "YES"
      Call SetBuildfileValue("SetupNetNameStatus", strStatusBypassed)
    Case Else
      Call SetupNetName()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DC"
      ' Nothing
    Case strSetupNetBind <> "YES"
      ' Nothing
    Case strClusterHost <> "YES"
      Call SetBuildfileValue("SetupNetBindStatus", strStatusBypassed)
    Case Else
      Call SetupNetBind()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DD"
      ' Nothing
    Case Else
      Call SetupAdapter()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DE"
      ' Nothing
    Case strSetupNoTCPNetBios <> "YES"
      ' Nothing
    Case Else
      Call SetupNoTCPNetBios()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DF"
      ' Nothing
    Case strSetupNoTCPOffload <> "YES"
      ' Nothing
    Case Else
      Call SetupNoTCPOffload()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DG"
      ' Nothing
    Case GetBuildfileValue("SetupTLS12") <> "YES"
      ' Nothing
    Case Else
      Call SetupTLS12()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1DH"
      ' Nothing
    Case GetBuildfileValue("SetupNoSSL3") <> "YES"
      ' Nothing
    Case Else
      Call SetupNoSSL3()
  End Select

  Call SetProcessId("1DZ", " Setup SQL Server Network" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupFirewall()
  Call SetProcessId("1DA", "Setup Firewall")
  Dim strPort

  Select Case True
    Case strClusterAction = ""
      ' Nothing
    Case Else
      Call OpenPort("Failover Clusters (UDP-In)",     "3343", "UDP", "IN", "")
  End Select

  Call OpenPort("RPC Endpoint Mapper",                "RPC-EPMap", "TCP", "IN", strDirSys & "\System32\svchost.exe")

  If strSetupISMaster = "YES" Then
    Call OpenPort("IS Master",                        strIsMasterPort, "TCP", "IN", "")
  End If

  If strSetupSQLDB = "YES" Then
    Call OpenPort("SQL Server (" & strInstance & ")", strTCPPort,    "TCP", "IN", "")
    Call OpenPort("SQL DAC",                          GetBuildfileValue("TCPPortDAC"), "TCP", "IN", "")
    Call OpenPort("SQL Service Broker",               "4022", "TCP", "IN", "")
    Call OpenPort("SQL DB Mirroring",                 GetBuildfileValue("TCPPortAO"),  "TCP", "IN", "")
    If strSetupAlwaysOn = "YES" Then
      Call OpenPort("SQL AG Health Probe",            "59999", "TCP", "IN", "")
    End If
    If strInstance <> "MSSQLSERVER" Then
      Call OpenPort("SQL Browser",                    "1434",  "UDP", "IN", "")
    End If
  End If

  If GetBuildfileValue("SetupSQLDebug") = "YES" Then ' See https://docs.microsoft.com/en-us/sql/ssms/scripting/configure-firewall-rules-before-running-the-tsql-debugger
    strPort         = GetBuildfileValue("TCPPortDebug")
    Call OpenPort("SQL Debug (TCP-In)", "RPC", "TCP", "IN", strDirProg & "\MSSQL" & strSQLVersionWMI & "." & strInstance & "\MSSQL\Binn\sqlservr.exe")
    Call OpenPort("SQL Debug (UDP-In)", "500,4500", "UDP", "IN", strDirProg & "\MSSQL" & strSQLVersionWMI & "." & strInstance & "\MSSQL\Binn\sqlservr.exe")
    If strPort <> "" Then
      Call OpenPort("VS SQL Debug (TCP-In)", strPort, "TCP", "IN", "")
    End If
    Call SetBuildfileValue("SetupSQLDebugStatus", strStatusComplete)
  End If

  If strSetupSQLIS = "YES" Then
    Call OpenPort("SQL Integration Services",         "",      "TCP", "IN", strVolSys & ":\Program Files\Microsoft SQL Server\" & strSQLVersionNum & "\DTS\Binn\MsDtsSrvr.exe")
  End If

  If strSetupPolyBase = "YES" Then
    Call OpenPort("PolyBase",                         strPBPortRange, "TCP", "IN", "")
  End If

  If strSetupSQLAS = "YES" Then
    Call OpenPort("SQL Analysis Server",              GetBuildfileValue("TCPPortAS"),  "TCP", "IN", "")
    If strInstance <> "MSSQLSERVER" Then
      Call OpenPort("SQL Browser",                    "2382", "TCP", "IN", "")
    End If
  End If

  Select Case True
    Case strSetupSQLDBFS <> "YES"
      ' Nothing
    Case strFSLevel < "2"
      ' Nothing
    Case Else
      Call OpenPort("SQL Filestream",                 "139,145", "TCP", "IN", "")
  End Select

  If strSetupSQLRS = "YES" Then
    Call OpenPort("HTTP",                             "80",   "TCP", "IN", "")
  End If

  Select Case True
    Case strSQLVersion > "SQL2005"
      ' Nothing
    Case strClusterAction = "" 
      ' Nothing
    Case Else
      Call OpenPort("SQL Server Setup",               "",     "TCP", "IN", strVolSys & ":\Program Files\Microsoft SQL Server\" & strSQLVersionNum & "\Setup Bootstrap\setup.exe")
      Call OpenPort("SQL Server Setup",               "",     "UDP", "IN", strVolSys & ":\Program Files\Microsoft SQL Server\" & strSQLVersionNum & "\Setup Bootstrap\setup.exe")
  End Select

  Call DebugLog("Add Firewall Exception for DTC")
  Select Case True
    Case (strSetupDTCNetAccess <> "YES") And (strClusterHost <> "YES")
      ' Nothing
    Case Left(strOSVersion, 1) >= "6"
      Call OpenPort("Distributed Transaction Coordinator (RPC)",       "",   "", "", "")
      Call OpenPort("Distributed Transaction Coordinator (RPC-EPMAP)", "",   "", "", "")
      Call OpenPort("Distributed Transaction Coordinator (TCP-In)",    "",   "", "", "")
      Call OpenPort("Distributed Transaction Coordinator (TCP-Out)",   "",   "", "", "")
    Case Else
      strCmd        = "NETSH FIREWALL ADD ALLOWEDPROGRAM NAME=""MSDTC"" "
      strCmd        = strCmd & "PROGRAM=""" & strDirSys & "\system32\msdtc.exe"" "
      strCmd        = strCmd & "MODE=ENABLE SCOPE=ALL PROFILE=DOMAIN"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

  Call SetBuildfileValue("SetupFirewallStatus", strStatusComplete)  
  Call ProcessEnd(strStatusComplete)

End Sub


Sub OpenPort(strFWName, strFWPort, strFWType, strFWDir, strFWProgram)
  Call DebugLog("OpenPort: " & strFWName & " for " & strFWPort)
  Dim strFWStatus

  strFWStatus       = CheckFWName(strFWName)
  Select Case True
    Case Left(strOSVersion, 1) < "6"
      Call SetFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWStatus)
    Case Else
      Call SetAdvFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWStatus)
  End Select

End Sub


Function CheckFWName(strFWName)
  Call DebugLog("CheckFWName:")
  Dim objFWRule

  CheckFWName       = False
  For Each objFWRule In objFWRules
    If objFWRule.Name = strFWName Then
      CheckFWName   = True
    End If
  Next

End Function


Sub SetFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWStatus)
  Call DebugLog("SetFirewall:")

  Select Case True
    Case strFirewallStatus <> "1"
      ' Nothing
    Case strFWStatus = True
      ' TBC
    Case Else
      strCmd        = "NETSH FIREWALL ADD PORTOPENING NAME=""" & strFWName & """ "
      strCmd        = strCmd & "PROTOCOL=" & strFWType & " MODE=ENABLE SCOPE=ALL PROFILE=DOMAIN "
      If strFWPort <> "" Then
        strCmd      = strCmd & "PORT=" & Replace(strFWPort, " ", "") & " "
      End If
      If strFWProgram <> "" Then
        strCmd      = strCmd & "PROGRAM=""" & strFWProgram & """ "
      End If
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
  End Select

End Sub


Sub SetAdvFirewall(strFWName, strFWPort, strFWType, strFWDir, strFWProgram, strFWStatus)
  Call DebugLog("SetAdvFirewall:")

  Select Case True
    Case strFWStatus = True
      strCmd        = "NETSH ADVFIREWALL FIREWALL SET RULE NAME=""" & strFWName & """ "
      strCmd        = strCmd & "NEW PROFILE=DOMAIN ENABLE=YES "
    Case Else
      strCmd        = "NETSH ADVFIREWALL FIREWALL ADD RULE NAME=""" & strFWName & """ "
      strCmd        = strCmd & "ACTION=ALLOW PROFILE=DOMAIN "
  End Select

  If strFWType <> "" Then
    strCmd          = strCmd & "PROTOCOL=" & strFWType & " "
  End If
  If strFWDir <> "" Then
    strCmd          = strCmd & "DIR=" & strFWDir & " "
  End If
  If strFWPort <> "" Then
    strCmd          = strCmd & "LOCALPORT=" & Replace(strFWPort, " ", "") & " "
  End If
  If strFWProgram <> "" Then
    strCmd          = strCmd & "PROGRAM=""" & strFWProgram & """ "
  End If
 
 Call Util_RunExec(strCmd, "", strResponseYes, 0)

End Sub


Sub SetupNetName()
  Call SetProcessId("1DB", "Setup Network Adaptor Names")
  Dim arrInterfaces, arrNetworks
  Dim colInterface, colNetwork
  Dim strInterfaceName, strNetNameSource, strNetworkName, strPathAdapter, strPathNetworks, strPathInterface, strPathInterfaces
  Dim intIdx, intIdxNew

  strNetNameSource  = GetBuildfileValue("NetNameSource")
  strPathNetworks   = "HKLM\Cluster\Networks\"
  objWMIReg.EnumKey strHKLM, Mid(strPathNetworks, 6), arrNetworks
  strPathInterfaces = "HKLM\Cluster\NetworkInterfaces\"
  objWMIReg.EnumKey strHKLM, Mid(strPathInterfaces, 6), arrInterfaces
  
  For Each colNetwork In arrNetworks
    strPath         = strPathNetworks & colNetwork & "\Name"
    strNetworkName  = objShell.RegRead(strPath)
    Call DebugLog("Processing Network " & strNetworkName)
    For Each colInterface In arrInterfaces
      strPathInterface = strPathInterfaces & colInterface
      Select Case True
        Case objShell.RegRead(strPathInterface & "\Network") <> colNetwork
          ' Nothing
        Case objShell.RegRead(strPathInterface & "\Node") <> strClusterNode
          ' Nothing
        Case Else
          strPathAdapter   = objShell.RegRead(strPathInterface & "\AdapterId")
          strInterfaceName = objShell.RegRead("HKLM\System\CurrentControlSet\Control\Network\{" & strNetworkGUID &"}\{" & strPathAdapter & "}\Connection\Name")
          Select Case True
            Case strInterfaceName = strNetworkName
              ' Nothing
            Case (strNetNameSource = "CLUSTER") Or (strClusterAction = "ADDNODE")
              strCmd = "NETSH INTERFACE SET INTERFACE NAME=""" & strInterfaceName & """ NEWNAME=""" & strNetworkName & """ "
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              Call DebugLog(" Network Adapter '" & strInterfaceName & "' renamed to '" & strNetworkName & "'")
              Wscript.Sleep strWaitShort 
            Case Else
              strCmd = "CLUSTER " & strClusterName & " NETWORK """ & strNetworkName & """ /RENAME:""" & strInterfaceName & """ "
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              Call DebugLog(" Cluster Network '" & strNetworkName & "' renamed to '" & strInterfaceName & "'")
              If strClusIPV4Network = strNetworkName Then
                strClusIPV4Network = strInterfaceName
                Call SetBuildfileValue("ClusIPV4Network", strClusIPV4Network)
              End If
              If strClusIPV6Network = strNetworkName Then
                strClusIPV6Network = strInterfaceName
                Call SetBuildfileValue("ClusIPV6Network", strClusIPV6Network)
              End If
              Wscript.Sleep strWaitShort 
          End Select
      End Select
    Next 
  Next

  Call SetBuildfileValue("SetupNetNameStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNetBind()
  Call SetProcessId("1DC", "Setup Network Bindings")
  Dim arrInterfaces
  Dim colInterface
  Dim strAdapter, strName, strNetwork, strNetworkRole, strNode, strPathInterface, strPathInterfaces, strPathNetwork

  strPathInterfaces = "HKLM\Cluster\NetworkInterfaces\"
  objWMIReg.EnumKey strHKLM, Mid(strPathInterfaces, 6), arrInterfaces
  For Each colInterface In arrInterfaces
    strPathInterface = strPathInterfaces & colInterface
    strNetwork      = objShell.RegRead(strPathInterface & "\Network")
    strNode         = objShell.RegRead(strPathInterface & "\Node")
    strPathNetwork  = "HKLM\Cluster\Networks\" & strNetwork
    strNetworkRole  = objShell.RegRead(strPathNetwork & "\Role")
    Select Case True
      Case strNode <> strClusterNode
        ' Nothing
      Case CStr(strNetworkRole) < "2"
        ' Nothing
      Case Else
        strAdapter  = objShell.RegRead(strPathInterface & "\AdapterId")
        strName     = objShell.RegRead(strPathNetwork & "\Name")
        Call SetBindingOrder("IPv4", strAdapter, strName)
        Call SetBindingOrder("IPv6", strAdapter, strName)
    End Select
  Next 

  Call SetBuildfileValue("SetupNetBindStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetBindingOrder(strTCPVersion, strAdapter, strName)
  Call DebugLog("SetBindingOrder: " & strTCPVersion & " for " & strAdapter)
  Dim arrBindings
  Dim bFound
  Dim intBind, intIdx, intIdxNew
  Dim strAdapterBind

  Select Case True
    Case strTCPVersion = "IPv6"
      strPath       = "HKLM\SYSTEM\CurrentControlSet\Services\TCPIP6\Linkage\"
    Case Else
      strPath       = "HKLM\SYSTEM\CurrentControlSet\Services\Tcpip\Linkage\"
  End Select
  objWMIReg.GetMultiStringValue strHKLM, Mid(strPath, 6), "Bind", arrBindings
  If IsNull(arrBindings) Then
    Exit Sub
  End If

  intBind           = Ubound(arrBindings)
  bFound            = False
  strAdapterBind    = "\Device\{" & strAdapter & "}"
  For intIdx = 0 To intBind
    If arrBindings(intIdx) = strAdapterBind Then
      bFound        = True
    End If
  Next
  If Not bFound Then
    Exit Sub
  End If

  Call DebugLog("Checking Bindings")
  ReDim arrBindingsNew(intBind)
  arrBindingsNew(0) = strAdapterBind
  intIdxNew         = 1
  strDebugMsg2      = "IdxNew: " & cStr(intIdxNew)
  For intIdx = 0 To intBind
    strDebugMsg1    = "Idx: " & cStr(intIdx)
    Select Case True
      Case arrBindings(intIdx) = strAdapterBind
        If intIdx <> 0 Then
          Call SetBuildMessage(strMsgInfo,  "TCP " & strTCPVersion & " Network Bind Order corrected for: " & strName)
        End If
      Case Else
        arrBindingsNew(intIdxNew) = arrBindings(intIdx)
        intIdxNew    = intIdxNew + 1
        strDebugMsg2 = "IdxNew: " & cStr(intIdxNew)
    End Select    
  Next
  objWMIReg.SetMultiStringValue strHKLM, Mid(strPath, 6), "Bind", arrBindingsNew
 
End Sub


Sub SetupAdapter()
  Call SetProcessId("1DD", "Setup Network Adapter Parameters")
  Dim arrInterfaces
  Dim intIdx
  Dim strNameServer, strDomain, strPathV4

  strPathV4         =  "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\"
  objWMIReg.EnumKey strHKLM, strPathV4, arrInterfaces

  For intIdx = 0 To Ubound(arrInterfaces)
    objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "NameServer", strNameServer
    objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain",     strDomain
    If (strNameServer > "") And (Not (strDomain > "")) Then
      Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
      objWMIReg.SetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain", strUserDNSDomain
    End If
  Next

  If strOSVersion >= "6.0" Then
    strPath         =  "SYSTEM\CurrentControlSet\Services\TCPIP6\Parameters\Interfaces\"
    objWMIReg.EnumKey strHKLM, strPath, arrInterfaces

    For intIdx = 0 To Ubound(arrInterfaces)
      objWMIReg.GetStringValue strHKLM, strPath   & arrInterfaces(intIdx), "NameServer", strNameServer
      objWMIReg.GetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain",     strDomain
      If (strNameServer > "") And (Not (strDomain > "")) Then
        Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
        objWMIReg.SetStringValue strHKLM, strPathV4 & arrInterfaces(intIdx), "Domain", strUserDNSDomain
      End If
    Next
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoTCPNetBios()
  Call SetProcessId("1DE", "Setup No TCP NetBios access")
' Based on code published by Mark Harris http://lifeofageekadmin.com/disable-netbios-over-tcpip-with-vbscript
  Dim arrInterfaces
  Dim intIdx

  strPath           =  "SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces\"
  objWMIReg.EnumKey strHKLM, strPath, arrInterfaces

  For intIdx = 0 To Ubound(arrInterfaces)
    Call DebugLog("Processing Adapter " & arrInterfaces(intIdx))
    objWMIReg.SetDWORDValue  strHKLM, strPath & arrInterfaces(intIdx), "NetBIOSOptions", Hex(2)
  Next

  Call SetBuildfileValue("SetupNoTCPNetBiosStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupNoTCPOffload()
  Call SetProcessId("1DF", "Setup No TCP Offload")
' Process described in KB976640
  Dim arrAdapters
  Dim colAdapter
  Dim intFound
  Dim strOffload, strPathAdapters

  intFound          = 0
  Call DebugLog("Turn of TCP Offload in Network Adapters")
  strPathAdapters   = "System\CurrentControlSet\Control\Class\{" & strNetworkGUID & "}\"
  objWMIReg.EnumKey strHKLM, strPathAdapters, arrAdapters
  For Each colAdapter In arrAdapters
    Select Case True
      Case colAdapter = "Properties"
        ' Nothing
      Case Else
        strPath     = strPathAdapters & colAdapter
        intFound    = intFound + AdapterOffloadDisable(strPath)
    End Select
  Next

  Call DebugLog("Turn of TCP Offload in Windows")
  strPath           = "System\CurrentControlSet\Services\TCPIP\Parameters"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisableTaskOffload",strOffload
  Select Case True
    Case strOffload = 1
      ' Nothing
    Case Else
      intFound       = 1
      strPath        = "HKLM\" & strPath & "\DisableTaskOffload"
      Call Util_RegWrite(strPath, "1", "REG_DWORD") 
  End Select

  If intFound > 0 Then
    Call DebugLog(" TCP Offload Disabled")  
  End If

  Call SetBuildfileValue("SetupNoTCPOffloadStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Function AdapterOffloadDisable(strPathAdapter)
  Call DebugLog("AdapterOffloadDisable: " & strPathAdapter)
  Dim arrValueNames, arrValueTypes
  Dim intFound, intIdx
  Dim strValueName, strValueType

  intFound          = 0
  objWMIReg.EnumValues strHKLM, strPathAdapter, arrValueNames, arrValueTypes
  Select Case True
    Case Not IsArray(arrValueNames)
      ' Nothing
    Case Else
      For intIdx = 0 To UBound(arrValueNames)
        strValueName = arrValueNames(intIdx)
        strValueType = arrValueTypes(intIdx)
        Select Case True
          Case Left(strValueName, 1) <> "*"
            ' Nothing
          Case Instr(strValueName, "Offload") > 0
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*FlowControl"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV1IPv4"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV2IPv4"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*LsoV2IPv6"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
          Case strValueName = "*RSS"
            intFound = intFound + OptionOffloadDisable(strPathAdapter, strValueName, strValueType)
        End Select
      Next
  End Select

  AdapterOffloadDisable = intFound

End Function


Function OptionOffloadDisable(strPathAdapter, strOption, strType)
  Call DebugLog("OptionOffloadDisable: " & strOption)
  Dim intFound
  Dim strOffload, strRegType, strRegValue

  intFound          = 0
  Select Case True
    Case strType = 4
      strRegType    = "REG_DWORD"
      strRegValue   = 0
      objWMIReg.GetDWordValue strHKLM,strPathAdapter,strOption,strOffload
    Case Else
      strRegType    = "REG_SZ"
      strRegValue   = "0"
      objWMIReg.GetStringValue strHKLM,strPathAdapter,strOption,strOffload
  End Select

  Select Case True
    Case IsNull(strOffload)
      ' Nothing
    Case strOffload = strRegValue
      ' Nothing
    Case Else
      intFound      = 1
      strPath       = "HKLM\" & strPathAdapter & "\" & strOption
      Call Util_RegWrite(strPath, strRegValue, strRegType) 
  End Select

  OptionOffloadDisable = intFound

End Function


Sub SetupTLS12()
  Call SetProcessId("1DG", "Setup TLS 1.2 Support")
' More information given in KB 3135244
  Dim intProcess, intProtocol, intRegValue

  intProcess        = 0
  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Client"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "0", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "1", "REG_DWORD")
      intProcess    = 1
  End Select

  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS 1.2\Server"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "0", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "1", "REG_DWORD")
      intProcess    = 1
  End Select
  
  intProtocol       = 0
  strPath           = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
  objWMIReg.GetDWordValue strHKLM,strPath,"DefaultSecureProtocols",intRegValue
  Select Case True
    Case strOSVersion > "6.0"
      ' Nothing
    Case IsNull(intRegValue)
      intRegValue   = 0
      intProtocol   = 1
    Case intRegValue Or 2048 > 0
      ' Nothing
    Case Else
      intProtocol   = 1
  End Select

  If intProtocol = 1 Then ' See KB 3140245
   intProcess       = 1
   intRegValue      = intRegValue Or 2048
    Call Util_RegWrite("HKLM\" & strPath & "\DefaultSecureProtocols", CStr(intRegValue), "REG_DWORD")
    If strProcArc = "X64" Then
      strPath       = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
       Call Util_RegWrite("HKLM\" & strPath & "\DefaultSecureProtocols", CStr(intRegValue), "REG_DWORD")
    End If
  End If

  Call TLS12NetRegistry("v2.0.50727")
  Call TLS12NetRegistry("v4.0.30319")

  Select Case True
    Case intProcess = 0
      Call SetBuildfileValue("SetupTLS12Status", strStatusPreConfig)
    Case Else
      Call SetBuildfileValue("RebootStatus", "Pending")
      Call SetBuildfileValue("SetupTLS12Status", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub TLS12NetRegistry(strNetPath)
  Call DebugLog("TLS12NetRegistry: " & strNetPath)
' For details see: https://docs.microsoft.com/en-us/configmgr/core/plan-design/security/enable-tls-1-2-client

  strPath           = "HKLM\SOFTWARE\Microsoft\.NETFramework\" & strNetPath
  Call Util_RegWrite(strPath & "\SystemDefaultTlsVersions", "1", "REG_DWORD")
  Call Util_RegWrite(strPath & "\SchUseStrongCrypto",       "1", "REG_DWORD")
  If strProcArc = "X64" Then
    strPath     = "HKLM\SOFTWARE\Wow6432Node\Microsoft\.NETFramework\" & strNetPath
    Call Util_RegWrite(strPath & "\SystemDefaultTlsVersions", "1", "REG_DWORD")
    Call Util_RegWrite(strPath & "\SchUseStrongCrypto",       "1", "REG_DWORD")
  End If

End Sub


Sub SetupNoSSL3()
  Call SetProcessId("1DH", "Disable SSL3")
' More information given in KB 3009008
  Dim intProcess, intProtocol, intRegValue

  intProcess        = 0
  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Client"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "1", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "0", "REG_DWORD")
      intProcess    = 1
  End Select

  strPath           = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\SSL 3.0\Server"
  objWMIReg.GetDWordValue strHKLM,strPath,"DisabledByDefault",intRegValue
  Select Case True
    Case intRegValue = 1
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\DisabledByDefault", "1", "REG_DWORD")
      intProcess    = 1
  End Select

  objWMIReg.GetDWordValue strHKLM,strPath,"Enabled",intRegValue
  Select Case True
    Case intRegValue = 0
      ' Nothing
    Case Else
      Call Util_RegWrite("HKLM\" & strPath & "\Enabled",           "0", "REG_DWORD")
      intProcess    = 1
  End Select

  intProtocol       = 0
  strPath           = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
  objWMIReg.GetDWordValue strHKLM,strPath,"DefaultSecureProtocols",intRegValue
  Select Case True ' See KB 3140245
    Case strOSVersion > "6.0"
      ' Nothing
    Case IsNull(intRegValue)
      intRegValue   = 0
      intProtocol   = 1
    Case intRegValue XOr 32 = 0
      ' Nothing
    Case Else
      intProtocol   = 1
  End Select

  If intProtocol = 1 Then
    intProcess      = 1
    intRegValue     = intRegValue XOr 32
    Call Util_RegWrite("HKLM\" & strPath & "\DefaultSecureProtocols", CStr(intRegValue), "REG_DWORD")
    If strProcArc = "X64" Then
      strPath       = "SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\WinHttp"
       Call Util_RegWrite("HKLM\" & strPath & "\DefaultSecureProtocols", CStr(intRegValue), "REG_DWORD")
    End If
  End If

 Select Case True
    Case intProcess = 0
      Call SetBuildfileValue("SetupNoSSL3Status", strStatusPreConfig)
    Case Else
      Call SetBuildfileValue("RebootStatus", "Pending")
      Call SetBuildfileValue("SetupNoSSL3Status", strStatusComplete)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAccounts()
  Call SetProcessId("1E", "Setup Accounts")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EA"
      ' Nothing
    Case Else
      Call SetupLocalGroups()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EB"
      ' Nothing
    Case Else
      Call SetupGroupRights()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EC"
      ' Nothing
    Case Else
      Call SetupAccountRights()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1ED"
      ' Nothing
    Case strSetupKerberos <> "YES"
      ' Nothing
    Case Else
      Call SetupKerberos()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1EE"
      ' Nothing
    Case strSetupNoWinGlobal <> "YES"
      ' Nothing
    Case Else
      Call SetupNoWinGlobal()
  End Select

  Call SetProcessId("1EZ", " Setup Accounts" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupLocalGroups()
  Call SetProcessId("1EA", "Setup Local Groups")

  Call SetupStandardGroups()

  Call ProcessAccounts("AssignUserGroups", "")

  Call DebugLog("Process Computer accounts")
  If strClusterHost = "YES" Then
    strCmd          = "NET LOCALGROUP """ & strGroupUsers & """ """ & strDomain & "\" & strClusterName & "$" & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  If strGroupMSA <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupUsers & """ """ & strDomain & "\" & strGroupMSA & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupStandardGroups()
  Call SetProcessId("1EAA", "Setup Standard Groups")
  Dim objAccount
  Dim intServerLen

  intServerLen      = Len(strServer) + 1

  If strGroupDistComUsers = "" Then
    strGroupDistComUsers = "Distributed COM Users"
    strCmd             = "NET LOCALGROUP """ & strGroupDistComUsers & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
    strCmd             = "Win32_Group.Domain='" & strLocalDomain & "',Name='" & strGroupDistComUsers & "'"
    Set objAccount     = objWMI.Get(strCmd) 
    strSIDDistComUsers = objAccount.SID
    Call SetBuildfileValue("SIDDistComUsers",    strSIDDistComUsers)
    Call SetBuildfileValue("GroupDistComUsers",  strGroupDistComUsers)
  End If

  If strGroupIISIUsers = "" Then
    strGroupIISIUsers = "IIS_IUSRS"
    strCmd             = "NET LOCALGROUP """ & strGroupIISIUsers & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
    strCmd             = "Win32_Group.Domain='" & strLocalDomain & "',Name='" & strGroupIISIUsers & "'"
    Set objAccount     = objWMI.Get(strCmd) 
    strSIDIISIUsers = objAccount.SID
    Call SetBuildfileValue("SIDIISIUsers",    strSIDIISIUsers)
    Call SetBuildfileValue("GroupIISIUsers",  strGroupIISIUsers)
  End If

End Sub


Sub SetupGroupRights()
  Call SetProcessId("1EB", "Setup Group Rights")

  Call RunNTRights("""" & strGroupUsers & """ +r SeNetworkLogonRight")
  Call RunNTRights("""" & strGroupUsers & """ +r SeInteractiveLogonRight")
  Call RunNTRights("""" & strGroupUsers & """ +r SeChangeNotifyPrivilege")

  Call RunNTRights("""" & strGroupAdmin & """ +r SeInteractiveLogonRight")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeRemoteInteractiveLogonRight")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeRemoteShutdownPrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeManageVolumePrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeProfileSingleProcessPrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeSystemProfilePrivilege")
  Call RunNTRights("""" & strGroupAdmin & """ +r SeShutdownPrivilege")

  If strGroupRDUsers <> "" Then
    Call RunNTRights("""" & strGroupRDUsers & """ +r SeInteractiveLogonRight")
    Call RunNTRights("""" & strGroupRDUsers & """ +r SeRemoteInteractiveLogonRight")
  End If

  If (strSetupCmdshell = "YES") And (strCmdshellAccount <> "") Then
    Call RunNTRights("""" & strCmdshellAccount & """ +r SeBatchLogonRight")
  End If

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupAccountRights()
  Call SetProcessId("1EC", "Setup Account Rights")
  Dim arrShareList
  Dim strShareList
  Dim intIdx

  Call ProcessAccounts("AssignAccountRights", "")

  strShareList      = GetBuildfileValue("ShareList")
  Select Case True
    Case strSetupShares <> "YES"
      ' Nothing
    Case strShareList = ""
      ' Nothing
    Case Else
      arrShareList  = Split(strShareList, ",")
      For intIdx = 1 To Ubound(arrShareList)
        If strSetupSQLDB = "YES" Then
          Call SetupRemoteShareRights(arrShareList(intIdx), strSQLAccount, "SqlAccount")
        End If
        If strSetupSQLDBAG = "YES" Then
          Call SetupRemoteShareRights(arrShareList(intIdx), strAgtAccount, "AgtAccount")
        End If
      Next
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupRemoteShareRights(strShareName, strAccount, strAccountParm)
  Call DebugLog("SetupRemoteShareRights: " & strShareName & " for " & strSQLAccount)
  Dim arrACEs
  Dim objACE, objACEAccount, objSecDesc, objShareSec, objWMIRemote
  Dim strRemoteServer, strShare
  Dim intIdx, intRC

  intIdx            = Instr(strShareName, "\")
  strRemoteServer   = Left(strShareName, intIdx - 1)
  strShare          = Mid(strShareName, intIdx + 1)

  On Error Resume Next
  Set objWMIRemote  = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strRemoteServer & "\root\cimv2")
  Wscript.Sleep strWaitShort
  Set objShareSec   = objWMIRemote.Get("Win32_LogicalShareSecuritySetting.Name=""" & strShare & """")
  If IsEmpty(objShareSec) Then
    Err.Clear
    Exit Sub
  End If

  On Error GoTo 0
  intRC             = objShareSec.GetSecurityDescriptor(objSecDesc)
  arrACEs           = objSecDesc.DACL
  Set objACEAccount = GetShareDACL(strSQLAccount, "Full", strAccountParm)
  For Each objACE In arrACEs
    If objACEAccount.Trustee.Name = objACE.Trustee.Name Then
      objACEAccount.Trustee.Name = ""
    End If
  Next

  Select Case True
    Case objACEAccount.Trustee.Name = ""
      ' Nothing
    Case Else
      intIdx        = UBound(arrACEs) + 1
      ReDim Preserve arrACEs(intIdx)
      Set arrACEs(intIdx) = objACEAccount
  End Select

  objSecDesc.DACL   = arrACEs
  intRC             = objShareSec.SetSecurityDescriptor(objSecDesc)

  Set objWMIRemote  = Nothing
  Call SetBuildfileValue("SetupSharesStatus", strStatusProgress) 

End Sub


Sub SetupKerberos()
  Call SetProcessId("1ED", "Setup Kerberos")
  Dim objInstParm, objCmdFile
  Dim strCmdFile, strCmdPath

  Call GetKerberosFile(objCmdFile, strCmdPath, strCmdFile)

  Select Case True
    Case strOSVersion >= "6.0"
      Call SetupSPN(objCmdFile, "SETSPN -S ")
    Case Else
      Call SetupSPN(objCmdFile, "SETSPN -D ")
      Call SetupSPN(objCmdFile, "SETSPN -A ")
  End Select

  Call WriteKerberosCmd(objCmdFile, "ECHO.")
  If strSetupSQLAS = "YES" Then
    Call SetupDelegation(objCmdFile, "AS", GetSPNAccount(strASAccount,  strServer))
  End If
  If strSetupSQLDBAG = "YES" Then
    Call SetupDelegation(objCmdFile, "AG", GetSPNAccount(strAgtAccount, strServer))
  End If
  If strSetupSQLDB = "YES" Then
    Call SetupDelegation(objCmdFile, "DB", GetSPNAccount(strSQLAccount, strServer))
  End If
  If strSetupSQLRS = "YES" Then
    Call SetupDelegation(objCmdFile, "RS", GetSPNAccount(strRSAccount,  strServer))
  End If

  Call SetupMSAGroup(objCmdFile)

  Call SetupDNSAlias(objCmdFile)

  If strOUCName <> "" Then
    Call SetupOUCName(objCmdFile)
  End If
  
  Call WriteKerberosCmd(objCmdFile, "ECHO.")
  Call WriteKerberosCmd(objCmdFile, "ECHO End of Kerberos related commands")
  objCmdFile.WriteLine "IF '%1' NEQ 'FineBuild' PAUSE"
  objCmdFile.WriteLine "EXIT /B %MAXRC%"
  objCmdFile.Close

  Call SetXMLParm(objInstParm, "PathMain",     strCmdPath)
  Call SetXMLParm(objInstParm, "ParmXtra",     "FineBuild")
  Call SetXMLParm(objInstParm, "InstallError", strMsgIgnore)
  Call RunInstall("Kerberos",  strCmdFile,     objInstParm)

  If GetBuildfileValue("SetupKerberosStatus") <> strStatusComplete Then
    Call SetBuildMessage(strMsgWarning, "Unable to setup Kerberos. " &strCmdPath & "\" &  strCmdFile & " must be run by a Domain Administrator")
    Call SetBuildfileValue("SetupKerberosStatus", strStatusManual)
  End If

  Call ProcessEnd("")

End Sub


Sub GetKerberosFile(objCmdFile, strCmdPath, strCmdFile)
  Call DebugLog("GetKerberosFile:")

  strCmdFile        = strInstance
  If strSetupSQLAS = "YES" Then
    strCmdFile      = strCmdFile & "AS"
  End If
  If strSetupSQLDB = "YES" Then
    strCmdFile      = strCmdFile & "DB"
  End If
  If strSetupSQLRS = "YES" Then
    strCmdFile      = strCmdFile & "RS"
  End If
  strCmdFile        = "SetupKerberos" & strCmdFile & ".bat"

  strCmdPath        = GetBuildfileValue("PathAutoConfig")
  Select Case True
    Case strInstance = "MSSQLSERVER"
      strCmdPath    = strCmdPath & strServer & "\Documentation"
    Case Else
      strCmdPath    = strCmdPath & strServer & "$" & strInstance & "\Documentation"
  End Select
  Call SetupFolder(strCmdPath, strSecDBA)

  strKerberosFile   = strCmdPath & "\" & strCmdFile
  strDebugMsg1      = "Kerberos commands: " & strKerberosFile
  Set objCmdFile    = objFSO.OpenTextFile(strKerberosFile, 2, True)

  objCmdFile.WriteLine "@ECHO OFF"
  objCmdFile.WriteLine "ECHO Kerberos Commands for " & strServer & " created on " & CStr(Date()) & " by " & strUserName & " Using SQL FineBuild " & strVersionFB
  objCmdFile.WriteLine "ECHO Kerberos Command File: " & strKerberosFile
  objCmdFile.WriteLine "REM Only use parameter of 'FineBuild' if called directly from FineBuild"
  objCmdFile.WriteLine "IF '%1' ==  'FineBuild' SET PSErrorPref='SilentlyContinue'"
  objCmdFile.WriteLine "IF '%1' NEQ 'FineBuild' SET PSErrorPref='Continue'"
  objCmdFile.WriteLine "SET CMDRC=0"
  objCmdFile.WriteLine "SET MAXRC=0"

End Sub


Sub SetupSPN(objCmdFile, strSPNCmd)
  Call DebugLog("SetupSPN: " & strSPNCmd)
  Dim strUserDomain

  strUserDomain     = ""
  If strUserDNSDomain <> "" Then
    strUserDomain   = "." & strUserDNSDomain
  End If

  If strSetupSQLAS = "YES" Then
    Call WriteKerberosCmd(objCmdFile, "ECHO.")
    Call WriteKerberosCmd(objCmdFile, "ECHO Setup SPNs for SSAS Account " & strASAccount)
  End If
  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
     Case strSetupSQLASCluster = "YES"
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strClusterNameAS) & " " & GetSPNAccount(strASAccount, strClusterNameAS))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strClusterNameAS & strUserDomain) & " " & GetSPNAccount(strASAccount, strClusterNameAS))
    Case strInstASSQL = "MSSQLSERVER"
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer) & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer & strUserDomain) & " " & GetSPNAccount(strASAccount, strServer))
    Case Else
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer) & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer & ":" & strInstASSQL) & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer & strUserDomain) & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("AS", "MSOLAPSvc.3/" & strServer & strUserDomain & ":" & strInstASSQL) & " " & GetSPNAccount(strASAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & "MSOLAPDisco.3/" & strServer & " " & GetSPNAccount(strSqlBrowserAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & "MSOLAPDisco.3/" & strServer & strUserDomain & " " & GetSPNAccount(strSqlBrowserAccount, strServer))
  End Select

  If strSetupSQLDB = "YES" Then
    Call WriteKerberosCmd(objCmdFile, "ECHO.")
    Call WriteKerberosCmd(objCmdFile, "ECHO Setup SPNs for SQL DB Account " & strSQLAccount)
  End If
  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strClusterNameSQL & GetSPNInstance(strInstance)) & " " & GetSPNAccount(strSQLAccount, strClusterNameSQL))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strClusterNameSQL & strUserDomain) & " " & GetSPNAccount(strSQLAccount, strClusterNameSQL))
      If strSetupAlwaysOn = "YES" Then
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strGroupAO) & " " & GetSPNAccount(strSQLAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strGroupAO & ":" & strTCPPort) & " " & GetSPNAccount(strSQLAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strGroupAO & strUserDomain) & " " & GetSPNAccount(strSQLAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strGroupAO & strUserDomain & ":" & strTCPPort) & " " & GetSPNAccount(strSQLAccount, strServer))
      End If
    Case Else
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strServer & GetSPNInstance(strInstance)) & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strServer & ":" & strTCPPort) & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strServer & strUserDomain & GetSPNInstance(strInstance)) & " " & GetSPNAccount(strSQLAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("DB", "MSSQLSvc/" & strServer & strUserDomain & ":" & strTCPPort) & " " & GetSPNAccount(strSQLAccount, strServer))
  End Select

  If strSetupSQLRS = "YES" Then
    Call WriteKerberosCmd(objCmdFile, "ECHO.")
    Call WriteKerberosCmd(objCmdFile, "ECHO Setup SPNs for RS Account " & strRSAccount)
  End If
  intIdx            = Instr(strRSAccount, "\")
  Select Case True
    Case strSetupSQLRS <> "YES"
      ' Nothing
    Case Else
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strServer) & " " & GetSPNAccount(strRSAccount, strServer))
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strServer & strUserDomain) & " " & GetSPNAccount(strRSAccount, strServer))
      If strSetupSQLRSCluster = "YES" Then
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strClusterNameRS) & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strClusterNameRS & strUserDomain) & " " & GetSPNAccount(strRSAccount, strServer))
      End If
      If strRSAlias <> "" Then
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strRSAlias) & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strRSAlias & strUserDomain) & " " & GetSPNAccount(strRSAccount, strServer))
      End If
      If strSetupAlwaysOn = "YES" Then
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strGroupAO) & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strGroupAO & strUserDomain) & " " & GetSPNAccount(strRSAccount, strServer))
      End If
      If strSetupPowerBI = "YES" Then
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "MSOLAPSvc.3/" & strServer & ":PBIRS") & " " & GetSPNAccount(strRSAccount, strServer))
        Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "MSOLAPSvc.3/" & strServer & strUserDomain & ":PBIRS") & " " & GetSPNAccount(strRSAccount, strServer))
      End If
  End Select

  Select Case True
    Case strSetupSQLRS = "YES"
      ' Nothing
    Case strSetupISMaster = "YES"
      Call WriteKerberosCmd(objCmdFile, "ECHO.")
      Call WriteKerberosCmd(objCmdFile, "ECHO Setup SPNs for Server " & strServer)
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strServer) & " " & strServer)
      Call WriteKerberosCmd(objCmdFile, strSPNCmd & SaveSPN("RS", "HTTP/" & strServer & strUserDomain) & " " & strServer)
  End Select

End Sub


Function SaveSPN(strSetupType, strSPN)
  Call DebugLog("SaveSPN:")
  Dim strSPNList

  strSPNList        = GetBuildfileValue("SPNList" & strSetupType)
  Select Case True
    Case strSPNList = ""
      strSPNList    = strSPN
    Case Else
      strSPNList    = strSPNList & "," & strSPN
  End Select
  Call SetBuildfileValue("SPNList" & strSetupType, strSPNList)

  SaveSPN           = strSPN

End Function 


Sub SetupDelegation(objCmdFile, strSetupType, strAccount)
  Call DebugLog("SetupDelegation: " & strSetupType & ", " & strAccount)

  Call SetConstrained(objCmdFile, strSetupType, strAccount)

  Call SetDelegates(objCmdFile, strSetupType, strAccount)

End Sub


Sub SetConstrained(objCmdFile, strSetupType, strAccount)
  Call DebugLog("SetConstrained:")
  Dim strControl

  strAccountGUID     = GetAccountAttr(strAccount, strUserDNSDomain, "ObjectGUID")
  strAccountDelegate = GetAccountAttr(strAccount, strUserDNSDomain, "msDS-AllowedToDelegateTo")
  strControl         = GetAccountAttr(strAccount, strUserDNSDomain, "userAccountControl")
  Select Case True
    Case strAccountGUID = ""
      ' Nothing
    Case strControl = 17305600
      ' Nothing
    Case Else
      strControl    = strControl XOr 524288   ' Ensure Unconstrained Delegation disabled
      strControl    = strControl Or  16777216 ' Add Constrained Delegation
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Set-ADObject -Identity """ & strAccountGUID & """ -Replace @{'userAccountControl'=" & CStr(strControl) & "}"
      Call WriteKerberosCmd(objCmdFile, "ECHO Set Constrained Delegation for " & strSetupType & " account " & strAccount)
      Call WriteKerberosCmd(objCmdFile, strCmd)
  End Select

End Sub


Sub SetDelegates(objCmdFile, strSetupType, strAccount)
  Call DebugLog("SetDelegates:")
  Dim arrDelegateNew, arrDelegateOld
  Dim bFound
  Dim strDelegateList, strDelegateNew, strDelegateOld

  strDelegateList   = ""
  Select Case True
    Case strAccountGUID = ""
      ' Nothing
    Case strSetupType = "AS"
      strDelegateList = GetBuildfileValue("SPNListAS")
      strDelegateList = strDelegateList & "," & GetBuildfileValue("SPNListDB")
    Case strSetupType = "AG"
      strDelegateList = GetBuildfileValue("SPNListAS")
      strDelegateList = strDelegateList & "," & GetBuildfileValue("SPNListDB")
    Case strSetupType = "DB"
     strDelegateList = GetBuildfileValue("SPNListAS")
      strDelegateList = strDelegateList & "," & GetBuildfileValue("SPNListDB")
    Case strSetupType = "RS"
      strDelegateList = GetBuildfileValue("SPNListAS")
      strDelegateList = strDelegateList & "," & GetBuildfileValue("SPNListDB")
      strDelegateList = strDelegateList & "," & GetBuildfileValue("SPNListRS")
  End Select
  strDebugMsg1      = "Delegate List: " & strDelegateList

  Select Case True
    Case IsArray(strAccountDelegate)
      arrDelegateOld =  strAccountDelegate
    Case Else
      arrDelegateOld =  Array("")
  End Select
  arrDelegateNew    =  Split(strDelegateList, ",")

  strDelegateList   = ""
  For Each strDelegateNew In arrDelegateNew
    If strDelegateNew <> "" Then
      bFound        = False
      For Each strDelegateOld In arrDelegateOld
        If strDelegateNew = strDelegateOld Then
          bFound    = True
        End If
      Next
      Select Case True
        Case bFound = True
          ' Nothing
        Case strDelegateList = ""
          strDelegateList = "'" & strDelegateNew & "'"
        Case Else
          strDelegateList = strDelegateList & ", '" & strDelegateNew & "'"
      End Select
    End If
  Next

  Select Case True
    Case strAccountGUID = ""
      ' Nothing
    Case strDelegateList = ""
      ' Nothing
    Case Else
      Call WriteKerberosCmd(objCmdFile, "ECHO Set Delegate List for " & strSetupType & " account " & strAccount)
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Set-ADObject -Identity """ & strAccountGUID & """ -Add @{'msDS-AllowedToDelegateTo'=@(" & strDelegateList & ")}"
      Call WriteKerberosCmd(objCmdFile, strCmd)
  End Select

End Sub


Sub SetupMSAGroup(objCmdFile)
  Call DebugLog("SetupMSAGroup:")
  Dim arrItems
  Dim intItems, intUItems

  arrItems          = Split(Trim(GetBuildfileValue("ListMSA")), " ")
  intUItems         = UBound(arrItems)
  For intItems = 0 To intUItems
    Call ProcessMSAGroup(objCmdFile, GetBuildfileValue(arrItems(intItems)), strGroupMSA)
  Next 

  Call ProcessMSAGroup(objCmdFile, strServer, strGroupMSA)
  
  If strClusterName <> "" Then
    Call ProcessMSAGroup(objCmdFile, strClusterName, strGroupMSA)
  End If

  If strGroupMSA <> "" Then
    Call WriteKerberosCmd(objCmdFile, "ECHO.")
    Call WriteKerberosCmd(objCmdFile, "ECHO Ensure Delegation of Control is set up for group: " & strGroupMSA)
    Call WriteKerberosCmd(objCmdFile, "ECHO - for details see https://github.com/SQL-FineBuild/Common/wiki/Delegation-of-Control")
  End If

End Sub


Sub ProcessMSAGroup(objCmdFile, strItem, strGroup)
  Call DebugLog("ProcessMSAGroup: " & strItem)
  Dim intIdx
  Dim strAccount

  strAccount        = FormatAccount(strItem)
  intIdx            = Instr(strAccount, "\")
  If intIdx > 0 Then
    strAccount      = Mid(strAccount, intIdx + 1)
  End If
  Select Case True
    Case strItem = ""
      ' Nothing
    Case InStr(" " & GetAccountAttr(strItem, strUserDNSDomain, "memberOf") & " ", " " & strGroup & " ") > 0
      ' Nothing
    Case Else
      strCmd      = "NET GROUP """ & strGroup & """ """ & strAccount & """ /ADD /DOMAIN"
      Call WriteKerberosCmd(objCmdFile, strCmd)
      Call SetBuildfileValue("RebootStatus", "Pending")
  End Select

End Sub


Sub SetupDNSAlias(objCmdFile)
  Call DebugLog("SetupDNSAlias:")
  Dim strAGName, strAGDagName, strSetupAOAlias, strSetupRSAlias, strRSAlias
  
  strAGName         = GetBuildfileValue("AGName")
  strAGDagName      = GetBuildfileValue("AGDagName")
  strSetupAOAlias   = GetBuildfileValue("SetupAOAlias")
  strSetupRSAlias   = GetBuildfileValue("SetupRSAlias")
  strRSAlias        = GetBuildfileValue("RSAlias")

  Select Case True
    Case strSetupAOAlias <> "YES"
      ' Nothing
    Case strAGDagName = ""
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      ' Nothing
    Case GetAddress(strAGDagName, "Alias", "") <> ""
      ' Nothing
    Case Else
      Call WriteKerberosCmd(objCmdFile, "ECHO.")
      Call WriteKerberosCmd(objCmdFile, "ECHO Create DNS Alias for " & strAGDagName)
      Call SetBuildfileValue("SetupAOAliasStatus", strStatusProgress)
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strAGDagName & """ -HostNameAlias """ & strGroupAO & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """ -ComputerName """ & GetBuildfileValue("UserDNSServer") & """ "
      Call WriteKerberosCmd(objCmdFile, strCmd)
  End Select

  Select Case True
    Case strSetupAOAlias <> "YES"
      ' Nothing
    Case strSetupSQLDBCluster = "YES"
      ' Nothing
    Case GetBuildfileValue("ActionAO") = "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("AOAliasOwner") <> ""
      Call SetBuildfileValue("SetupAOAliasStatus", strStatusPreConfig)
    Case Else
      Call WriteKerberosCmd(objCmdFile, "ECHO.")
      Call WriteKerberosCmd(objCmdFile, "ECHO Create DNS Alias for " & strAGName)
      Call SetBuildfileValue("SetupAOAliasStatus", strStatusProgress)
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strAGName & """ -HostNameAlias """ & strServer & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """"
      Call WriteKerberosCmd(objCmdFile, strCmd)
  End Select

  Select Case True
    Case strSetupRSAlias <> "YES"
      ' Nothing
    Case GetBuildfileValue("ActionSQLRS") = "ADDNODE"
      ' Nothing
    Case GetAddress(strRSAlias, "Alias", "") <> ""
      ' Nothing
    Case Else
      Call SetupRSDNSAlias(objCmdFile, strRSAlias, strAGDagName, GetBuildfileValue("ClusterGroupRS"), strAGName)
  End Select

End Sub


Sub SetupOUCName(objCmdFile)
  Call DebugLog("SetupOUCName:")

  Call SetOUCName(objCmdFile, strServer)

  If strClusterHost = "YES" Then
    Call SetOUCName(objCmdFile, strClusterName)
  End If

End Sub



Sub SetOUCName(objCmdFile, strADObject)
  Call DebugLog("SetOUCName: " & strADObject)
  Dim strAccountGUID, strADCName

  strADCName        = "CN=" & strADObject & "," & strOUCName
  Select Case True
    Case GetAccountAttr(strADObject, strUserDNSDomain, "distinguishedName") = strADCName
      ' Nothing
    Case Else
      Call WriteKerberosCmd(objCmdFile, "ECHO Move " & strADObject & " to " & strOUPath)
      strAccountGUID = GetAccountAttr(strADObject, strUserDNSDomain, "ObjectGUID")
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Move-ADObject -Identity """ & strAccountGUID & """ -TargetPath """ & strOUCName & """ "
      Call WriteKerberosCmd(objCmdFile, strCmd)
  End Select

End Sub


Sub SetupRSDNSAlias(objCmdFile, strRSAlias, strAGDagName, strClusterGroupRS, strAGName)
  Call DebugLog("SetupRSDNSAlias: ")

  Call WriteKerberosCmd(objCmdFile, "ECHO.")
  Call WriteKerberosCmd(objCmdFile, "ECHO Create DNS Alias for " & strRSAlias)
  Call SetBuildfileValue("SetupRSAliasStatus", strStatusProgress)

  Select Case True
    Case strAGDagName <> ""
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strRSAlias & """ -HostNameAlias """ & strAGDagName      & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """"
    Case strClusterGroupRS <> ""
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strRSAlias & """ -HostNameAlias """ & strClusterGroupRS & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """"
    Case strAGName <> ""
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strRSAlias & """ -HostNameAlias """ & strAGName         & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """"
    Case Else
      strCmd        = "POWERSHELL $ErrorActionPreference = %PSErrorPref%;Add-DnsServerResourceRecordCName -Name """ & strRSAlias & """ -HostNameAlias """ & strServer         & "." & strUserDNSDomain & """ -ZoneName """ & strUserDNSDomain & """"
  End Select
  Call WriteKerberosCmd(objCmdFile, strCmd)

End Sub


Sub WriteKerberosCmd(objCmdFile, strKerberosCmd)
  Call DebugLog("WriteKerberosCmd: " & strKerberosCmd)

  Select Case True
    Case Left(strKerberosCmd, 4) = "REM "
      objCmdFile.WriteLine strKerberosCmd
    Case Left(strKerberosCmd, 4) = "ECHO"
      objCmdFile.WriteLine strKerberosCmd
    Case Else
      objCmdFile.WriteLine "ECHO " & strKerberosCmd
      objCmdFile.WriteLine strKerberosCmd
      objCmdFile.WriteLine "SET CMDRC=%ERRORLEVEL%"
      If Left(strKerberosCmd, 11) <> "POWERSHELL " Then
        objCmdFile.WriteLine "IF %CMDRC% == 1 SET CMDRC=0"
      End If
      objCmdFile.WriteLine "IF %CMDRC% LSS 0 SET /A CMDRC=0 - %CMDRC%"
      objCmdFile.WriteLine "IF %CMDRC% GTR %MAXRC% SET MAXRC=%CMDRC%"
  End Select

End Sub


Function GetSPNAccount(strAccount, strHost)
  Call DebugLog("GetSPNAccount: " & strAccount)
  Dim intIdx
  Dim strSPNAccount

  intIdx            = Instr(strAccount, "\")
  Select Case True
    Case intIdx = 0
      strSPNAccount = strHost
    Case Left(strAccount, intIdx - 1) <> strDomain
      strSPNAccount = strHost
    Case Else
      strSPNAccount = strAccount
  End Select

  GetSPNAccount     = UCase(strSPNAccount)

End Function


Function GetSPNInstance(strInstance)
  Call DebugLog("GetSPNInstance: " & strInstance)
  Dim strSPNInstance

  strSPNInstance    = ""
  If UCase(strInstance) <> "MSSQLSERVER" Then
    strSPNInstance  = ":" & strInstance
  End If

  GetSPNInstance    = UCase(strSPNInstance)

End Function


Sub SetupNoWinGlobal()
  Call SetProcessId("1EE", "Disble Windows Guest Access")
  Dim objUser
  Dim strAccount, strAccountSID
' Do not remove 'Authenticated Users', it is needed for Kerberos

  If strType <> "WORKSTATION" Then
    Call DebugLog("Disable Domain Non-Specific Access")
    Call RemoveUser(strGroupUsers, "S-1-1-0",  "L") ' Everyone
    Call RemoveUser(strGroupUsers, "S-1-5-4",  "L") ' NT AUTHORITY\INTERACTIVE
    Call RemoveUser(strGroupUsers, "S-1-5-7",  "L") ' NT AUTHORITY\Anonymous
    Call RemoveUser(strGroupUsers, "S-1-5-13", "L") ' NT AUTHORITY\Terminal Service Users
    If strDomainSID <> "" Then
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-501", "D") ' domain\Guest
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-513", "D") ' domain\Domain Users
      Call RemoveUser(strGroupUsers, "S-1-5-21-" & strDomainSID & "-514", "D") ' domain\Domain Guests
    End If
  End If

  Call DebugLog("Disable Local Guest Account")
  Call GetGuestAccount(strAccount, strAccountSID)
  If strAccount <> "" Then
    Call RemoveUser(strGroupUsers, strAccountSID, "L") '  Local Guest
    strDebugMsg1    = "Disabling local Guest account: " & strAccount
    Set objUser     = GetObject("WinNT://./" & strAccount)
    objUser.AccountDisabled = True
    objUser.SetInfo
  End If

  Call SetBuildfileValue("SetupNoWinGlobalStatus", strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub RemoveUser(strGroup, strSID, strType)
  Call DebugLog("RemoveUser: " & strGroup & " for " & strSID)
  Dim objAccount
  Dim strAccount

  Set objAccount    = objWMI.Get("Win32_SID.SID='" & strSid & "'") 
  strAccount        = objAccount.AccountName
  Select Case True
    Case strAccount = ""
      ' Nothing
    Case strType = "L"
      strCmd        = "NET LOCALGROUP """ & strGroup & """ """ & strAccount & """ /DELETE"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
    Case Else
      strCmd        = "NET LOCALGROUP """ & strGroup & """ """ & objAccount.ReferencedDomainName & "\" & strAccount & """ /DELETE"
      Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End Select

End Sub


Sub GetGuestAccount(strAccount, strAccountSID)
  Call DebugLog("GetGuestAccount:")
  Dim colUsers
  Dim objUser

  strAccount        = ""
  Set colUsers      = objWMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount=True") 
  For Each objUser In colUsers
    If Mid(objUser.SID, InstrRev(objUser.SID, "-") + 1) = "501" Then
      strAccount    = objUser.Name
      strAccountSID = objUser.SID
    End If
  Next

End Sub


Sub SetupVolumes()
  Call SetProcessId("1F", "Setup Volumes")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FA"
      ' Nothing
    Case Else
      Call SetupVolumeLabels()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1FB"
      ' Nothing
    Case Else
      Call SetupVolumeShares()
  End Select

  Call SetProcessId("1FZ", " Setup Volumes" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupVolumeLabels()
  Call SetProcessId("1FA", "Setup Volume Labels")

  Call DebugLog("Setup System volumes")
  Call SetupVolume("VolSys",           strVolSys,      strLabSystem)
  Call SetupVolume("VolProg",          strVolProg,     strLabProg)
  Call SetupVolume("VolDBA",           strVolDBA,      strLabDBA)
  If strSetupTempWin = "YES" Then
    Call SetupVolume("VolTempWin",     strVolTempWin,  strLabTempWin)
  End If

  Select Case True
    Case strSetupDTCCluster <> "YES"
      ' Nothing
    Case strActionDTC = "ADDNODE"
      ' Nothing
    Case (strOSVersion < "6.0") And (strDTCClusterRes > "")
      ' Nothing
    Case (strOSVersion >= "6.0") And (strDTCClusterRes > "") And (strDTCMultiInstance <> "YES")
      ' Nothing
    Case Else  
      Call DebugLog("Setup MSDTC volume")
      Call SetupVolume("VolDTC",       strVolDTC,      strLabDTC)
  End Select

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strActionSQLDB = "ADDNODE"
      If GetBuildfileValue("VolTempType") = "L" Then
        Call SetupVolume("VolTemp",      strVolTemp,     strLabTemp)
      End If
    Case Else
      Call DebugLog("Setup SQL volumes")
      Call SetupVolume("VolData",        strVolData,     strLabData)
      Call SetupVolume("VolLog",         strVolLog,      strLabLog)
      Call SetupVolume("VolSysDB",       strVolSysDB,    strLabSysDB)
      Call SetupVolume("VolBackup",      strVolBackup,   strLabBackup)
      Select Case True
        Case strSQLVersion >= "SQL2012"
          Call SetupVolume("VolTemp",    strVolTemp,     strLabTemp)
          Call SetupVolume("VolLogTemp", strVolLogTemp,  strLabLog)
        Case Else
          Call SetupVolume("VolTemp",    strVolTemp,     strLabTemp)
          Call SetupVolume("VolLogTemp", strVolLogTemp,  strLabLog)
        End Select
      If strSetupBPE = "YES" Then
        Call SetupVolume("VolBPE",       strVolBPE,      strLabBPE)
      End If
      If strSetupSQLDBFS = "YES" Then
        Call SetupVolume("VolDataFS",    strVolDataFS,   strLabDataFS)
      End If
      If strSetupSQLDBFT = "YES" Then
        Call SetupVolume("VolDataFT",    strVolDataFT,   strLabDataFT)
      End If
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strActionSQLAS = "ADDNODE"
      ' Nothing
    Case Else
      Call DebugLog("Setup SQL AS volumes")
      Call SetupVolume("VolDataAS",      strVolDataAS,   strLabDataAS)
      Call SetupVolume("VolLogAS",       strVolLogAS,    strLabLogAS)
      Call SetupVolume("VolBackupAS",    strVolBackupAS, strLabBackupAS)
      Call SetupVolume("VolTempAS",      strVolTempAS,   strLabTempAS)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupVolume(strVolParam, strVolList, strVolLabel)
  Call DebugLog("SetupVolume: " & strVolParam)
  Dim arrItems
  Dim strVol, strVolReq, strVolSource

  strVolReq         = GetBuildfileValue(strVolParam & "Req")
  strVolSource      = GetBuildfileValue(strVolParam & "Source")
  Select Case True
    Case strVolSource = "C"
      arrItems      = Split(Replace(strVolList, ",", " "))
      For intIdx = 0 To UBound(arrItems)
        strVol      = arrItems(intIdx)
        strVol      = UCase(Mid(strVol, Len(strCSVRoot) + 1))
        Call SetupThisCSV(strVolParam, strVol, strVolLabel)
      Next
    Case strVolSource = "D"
      For intIdx = 1 To Len(strVolList)
        strVol      = Mid(strVolList, intIdx, 1)
        Call SetupThisDrive(strVolParam, strVol, strVolLabel, strVolList, strVolReq)
      Next
  End Select

End Sub


Sub SetupThisCSV(strVolParam, strVol, strVolLabel)
  Call DebugLog("SetupThisCSV: " & strVol)
  Dim intIdx
  Dim strResName, strVolName

  intIdx            = Instr(strVol, "\")
  Select Case True
    Case intIdx > 0
      strVolName    = Left(strVol, intIdx - 1)
    Case Else
      strVolName    = strVol
  End Select
  strResName        = GetBuildfileValue("Vol_" & strVolName & "Res")

  Select Case True
    Case UCase(strResName) = UCase(strVolName)
      ' Nothing
    Case Else
      strCmd        = "CLUSTER """ & strClusterName & """ RESOURCE """ & strResName & """ /RENAME:""" & strVolName & """"
      Call Util_RunExec(strCmd, "", strResponseYes, 0)
      Call SetBuildfileValue("Vol_" & strVolName & "Res",    strVolName)
  End Select

End Sub


Sub SetupThisDrive(strVolParam, strVol, strVolLabel, strVolList, strVolReq)
  Call DebugLog("SetupThisDrive: " & strVol)
  Dim strNewLabel

  strNewLabel       = GetDriveLabel(strVolParam, strVol, strVolList, strVolLabel)

  Select Case True
    Case (Instr(strDriveList, strVol) = 0) And (strVol <> Left(strDriveList, Len(strVol)))
      ' No action, not a valid drive
    Case (Not objFSO.FolderExists(strVol & ":\")) And (strVol <> Left(strDriveList, Len(strVol)))
      Call DebugLog("Setup " & strVol & ": for " & strNewLabel & strStatusBypassed)
    Case Instr(strVolUsed, strVol) > 0
      ' Nothing
    Case Else
      Call DebugLog("Setup " & strVol & ": drive for " & strNewLabel)
      strVolUsed    = strVolUsed & " " & strVol
      If strClusterAction <> "" Then
        Call LabelThisClusterDrive(strVol, strNewLabel, strVolReq)
      End If
      Call LabelThisDrive(strVolParam,strVol, strNewLabel, strVolReq)
      Call CreateThisShare(strVol, strNewLabel)
  End Select

End Sub


Function GetDriveLabel(strVolParam, strVol, strVolList, strVolLabel)
  Call DebugLog("GetDriveLabel: " & strVolParam)
  Dim strVolNewLabel

  Select Case True
    Case Len(strVolList) = 1 
      strVolNewLabel = strVolLabel
    Case Else
      strVolNewLabel = strVolLabel & strVol
  End Select

  Select Case True
    Case strVol = strVolSys
      ' Nothing
    Case strVol = strVolProg
      ' Nothing
    Case strVol = strVolDTC
      ' Nothing
    Case strInstance = "MSSQLSERVER"
      ' Nothing
    Case strInstance = "SQLEXPRESS"
      ' Nothing
    Case Else
      strVolNewLabel = strVolNewLabel & "-" & strInstance
  End Select 

  If strLabPrefix <> "" Then
    strVolNewLabel  = Left(strLabPrefix & "-" & Replace(strVolNewLabel, " ", ""), 32)
    Call SetBuildfileValue("Lab" & Mid(strVolParam, 4), strVolNewLabel)
  End If

  strVolNewLabel    = Left(strVolNewLabel, 32)
  GetDriveLabel     = strVolNewLabel

End Function


Sub LabelThisClusterDrive(strVol, strVolLabel, strVolReq)
  Call DebugLog("LabelThisClusterDrive: " & strVol)
  Dim colClusGroups, colClusPartitions, colClusResources
  Dim objClusDisk, objClusGroup, objClusPartition, objClusResource
  Dim intCount, intFound

  strCmd            = "CLUSTER """ & strClusterName & """ GROUP """ & strClusStorage & """ /MOVETO:""" & strServer & """" 
  Call Util_RunExec(strCmd, "", strResponseYes, 0)

  intFound          = 0
  Set colClusGroups = GetClusterGroups()
  For Each objClusGroup In colClusGroups                   
    Set colClusResources = objClusGroup.Resources
    intCount        = 99
    For Each objClusResource In colClusResources
      If objClusResource.TypeName = "Physical Disk" Then
        Set objClusDisk           = objClusResource.Disk
        Set colClusPartitions     = objClusDisk.Partitions
        For Each objClusPartition In colClusPartitions
          Select Case True
            Case intFound <> 0
              ' Nothing
            Case strVolReq = "L" And Left(objClusPartition.DeviceName, 1) = strVol 
              ' Nothing 
            Case Left(objClusPartition.DeviceName, 1) = strVol 
              intCount = colClusResources.Count
              intFound = 1
              strCmd   = "CLUSTER """ & strClusterName & """ GROUP """ & objClusGroup.Name & """ /MOVETO:""" & strServer & """" 
              Call Util_RunExec(strCmd, "", strResponseYes, 0)
              strCmd   = "CLUSTER """ & strClusterName & """ RESOURCE """ & objClusResource.Name & """ /MOVE:""" & strClusStorage & """"
              Call Util_RunExec(strCmd, "", strResponseYes, 183)
              strCmd   = "CLUSTER """ & strClusterName & """ RESOURCE """ & objClusResource.Name & """ /RENAME:""" & strVolLabel & """"
              Call Util_RunExec(strCmd, "", strResponseYes, 0) 
          End Select
        Next
      End If
    Next

    Select Case True
      Case intCount > 1
        ' Nothing
      Case UCase(objClusGroup.Name) = UCase(strClusStorage)
        ' Nothing
      Case Else
        Call DebugLog("Delete empty group " & objClusGroup.Name)
        strCmd      = "CLUSTER """ & strClusterName & """ GROUP """ & objClusGroup.Name & """ /DELETE"
        Call Util_RunExec(strCmd, "", strResponseYes, 5010)
    End Select
  Next 

End Sub


Sub LabelThisDrive(strVolParam, strVol, strVolLabel, strVolReq)
  Call DebugLog("LabelThisDrive: " & strVol)
' Code to clear IndexingEnabled flag adapted from "Windows Server Cookbook" by Robbie Allen, ISBN 0-596-00633-0
  Dim strVolType

  If Not objFSO.FolderExists(strVol & ":\") Then
    Call SetBuildMessage(strMsgError, "Volume not found: " & strVol & ":\")
  End If

  Select Case True
    Case strOSVersion > "5.1"
      strCmd        = "SELECT * FROM Win32_Volume WHERE DriveLetter='" & strVol & ":'"
      Set colVol    = objWMI.ExecQuery(strCmd)
      For Each objVol In colVol
        If strSetupNoDriveIndex = "YES" Then
          objVol.IndexingEnabled = 0
        End If
        objVol.Label = strVolLabel
        objVol.Put_
      Next
    Case Else
      strCmd        = "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & strVol & ":'"
      Set colVol    = objWMI.ExecQuery(strCmd)
      For Each objVol In colVol
        objVol.VolumeName = strVolLabel
        objVol.Put_
      Next
  End Select

  If strSetupNoDriveIndex = "YES" Then
    Call DebugLog("Clearing index attribute from drive " & strVol)
    strCmd            = "ATTRIB +I " & strVol & ":\*.* /D /S"
    Call Util_RunCmdAsync(strCmd, 0)
    Call SetBuildfileValue("SetupNoDriveIndexStatus", strStatusComplete)
  End If

  Call SetBuildfileValue("Vol" & strVol & "Label", strVolLabel)

  strVolType        = GetBuildfileValue("Vol" & strVol & "Type")
  If strVolType = "" Then
    Call SetBuildfileValue("Vol" & strVol & "Type", "L")
  End If
  Select Case True
    Case strVolReq = "C" And Instr("CX", strVolType) = 0
      Call SetBuildMessage(strMsgError, strVolParam & ": " & strVol & ": must be a Cluster Drive")
    Case strVolReq = "L" And strVolType = "C"
      Call SetBuildMessage(strMsgError, strVolParam & ": " & strVol & ": must NOT be a Cluster Drive")
  End Select

End Sub


Sub SetupVolumeShares()
  Call SetProcessId("1FB", "Setup Volume Shares")

' KB245117 fix for share visibility
  Select Case True
    Case strSetupShares <> "YES"
      ' Nothing
    Case Left(strOSVersion, 1) >= "6"
      Call SetBuildfileValue("SetupSharesStatus", strStatusComplete) 
    Case Else
      strPath       = "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareServer"
      Call Util_RegWrite(strPath, 1, "REG_DWORD")
      strPath       = "HKLM\System\CurrentControlSet\Services\LanmanServer\Parameters\AutoShareWks"
      Call Util_RegWrite(strPath, 1, "REG_DWORD")
      Call SetBuildfileValue("RebootStatus", "Pending")
      Call SetBuildfileValue("SetupSharesStatus", strStatusComplete)     
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub CreateThisShare(strVol, strVolLabel)
  Call DebugLog("CreateThisShare: " & strVol)
  Dim strShareName

  strVolType        = GetBuildfileValue("Vol" & strVol & "Type")
  strShareName      = "(" & strVol & ") " & strVolLabel
  Select Case True
    Case strSetupShares <> "YES" 
      ' Nothing
    Case strVolType = "L"
      Call SetupLocalShare(strVol & ":\", strShareName)
      Call SetBuildfileValue("SetupSharesStatus", strStatusProgress)
  End Select

  Call SetBuildfileValue("Vol" & strVol & "Share", strShareName)

End Sub


Sub SetupLocalShare(strVol, strShareName)
  Call DebugLog("SetupLocalShare: " & strVol)
  Dim objACEAdmin, objACEUser, objSecDesc, objShare, objShareParm

  Set objSecDesc    = objWMI.Get("Win32_SecurityDescriptor").SpawnInstance_
  Set objShare      = objWMI.Get("Win32_Share")
  Set objACEAdmin   = GetShareDACL(strGroupAdmin, "Full",   "")
  Set objACEUser    = GetShareDACL(strGroupUsers, "Change", "")
  Set objShareParm  = objShare.Methods_("Create").InParameters.SpawnInstance_ 

  objSecDesc.DACL          = Array(objACEAdmin, objACEUser)
  objShareParm.Access      = objSecDesc
  objShareParm.Description = strShareName & " Share"
  objShareParm.Name = strShareName
  objShareParm.Path = strVol
  objShareParm.Type = 0
  objShare.ExecMethod_ "Create",  objShareParm

End Sub


Function GetShareDACL(strAccount, strAccess, strAccountParm)
  Call DebugLog("GetShareDACL: " & strAccount)
  Dim objACE, objTrustee

  Set objTrustee    = SetTrustee(strAccount, strAccountParm)
  Set objACE        = objWMI.Get("Win32_Ace").SpawnInstance_

  objACE.AceFlags   = 3
  objACE.AceType    = 0
  objACE.Trustee    = objTrustee
  Select Case True
    Case strAccess = "Full"
      objACE.AccessMask = 2032127
    Case Else
      objACE.AccessMask = 1245631 ' Change
  End Select

  Set GetShareDACL  = objAce

End Function


Function SetTrustee(strAccount, strAccountParm) 
  Call DebugLog("SetTrustee: " & strAccount & " for " & strAccountParm)
  Dim objRecordSet, objTrustee
  Dim strAttrObject, strDNSDomain, strLocal, strSID, strSIDBinary, strQueryDomain, strQueryAccount
  Dim intIdx

  strLocal          = ""
  intIdx            = InStr(strAccount, "\")
  Select Case True
    Case intIdx = 0
      strDNSDomain     = strServer
      strQueryDomain   = strServer
      strQueryAccount  = strAccount
    Case Left(strAccount, intIdx - 1) = strDomain
      strDNSDomain     = strUserDNSDomain
      strQueryDomain   = strDomain
      strQueryAccount  = Mid(strAccount, intIdx + 1)
    Case Else
      strDNSDomain     = strServer
      strQueryDomain   = strServer
      strQueryAccount  = Mid(strAccount, intIdx + 1)
      strLocal         = ",LocalAccount=True"
  End Select
  strDebugMsg1      = "QueryDomain=" & strQueryDomain & " QueryAccount=" & strQueryAccount

  Select Case True
    Case strAccountParm = ""
      strDebugMsg2  = "Group Account"
      strSID        = objWMI.Get("Win32_Group.Domain='" & strQueryDomain & "',Name='" & strQueryAccount & "'" & strLocal).SID
      strSIDBinary  = objWMI.Get("Win32_SID.SID='" & strSID &"'").BinaryRepresentation 
    Case GetBuildfileValue(strAccountParm & "Type") = "M"
      strDebugMsg2  = "MSA Account"
      strAttrObject = "objectClass=msDS-GroupManagedServiceAccount"
      objADOCmd.CommandText = "<LDAP://DC=" & Replace(strDNSDomain, ".", ",DC=") & ">;(&(" & strAttrObject & ")(CN=" & strQueryAccount & "));CN,objectSID"
      Set objRecordSet  = objADOCmd.Execute
      objRecordset.MoveFirst
      strSIDBinary = objRecordset.Fields(1).Value
    Case Else
      strDebugMsg2  = "User Account"
      strAttrObject = "objectClass=user"
      objADOCmd.CommandText = "<LDAP://DC=" & Replace(strDNSDomain, ".", ",DC=") & ">;(&(" & strAttrObject & ")(CN=" & strQueryAccount & "));CN,SID"
      Set objRecordSet  = objADOCmd.Execute
      objRecordset.MoveFirst
      strSIDBinary = objRecordset.Fields(1).Value
  End Select

  Set objTrustee    = objWMI.Get("Win32_Trustee").Spawninstance_ 
  objTrustee.Domain = strQueryDomain 
  objTrustee.Name   = strQueryAccount
  objTrustee.SID    = strSIDBinary
 
  Set SetTrustee    = objTrustee 

End Function


Sub SetupFolders()
  Call SetProcessId("1G", "Setup SQL Server folders")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GA"
      ' Nothing
    Case Else
      Call SetupSystemFolders()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GB"
      ' Nothing
    Case Else
      Call SetupStdVolumes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GC"
      ' Nothing
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLServer()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GD"
      ' Nothing
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case Else
      Call SetupSQLASVolumes()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GE"
      ' Nothing
    Case strSetupTempWin <> "YES"
      ' Nothing
    Case Else
      Call SetupTempVolume()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1GF"
      ' Nothing
    Case strSetupTempWin <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("1GF", "Set \Temp Folder Location", "SetTempLoc")
      Call SetBuildFileValue("SetupTempWinStatus", strStatusComplete)
  End Select

  Call SetProcessId("1GZ", " Setup SQL Server Folders" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupSystemFolders()
  Call SetProcessId("1GA", "Setup System folders")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Account1",   strUserAccount)
  Call SetXMLParm(objFolderParm, "Access",     strSecMain)
  Call PrepareFolder("Prog",     objFolderParm)

  Call SetXMLParm(objFolderParm, "Account1",   strUserAccount)
  Call SetXMLParm(objFolderParm, "Access",     strSecMain)
  Call PrepareFolder("ProgX86",  objFolderParm)

  Call SetPathPS()

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetPathPS()
  Call DebugLog("SetPathPS:")
  Dim strPathPS, strPathPSWin, strPathDirProg

  strPathPS         = GetBuildfileValue("PathPS")
  strPathPSWin      = colSysEnvVars("PSModulePath")
  strPathDirProg    = colSysEnvVars("ProgramFiles")

  Call SetupFolder(strPathPS, strSecMain)

  Select Case True
    Case strPathPSWin = ""
      strPathPSWin  = strPathPS
    Case Instr(UCase(strPathPSWin & ";"), UCase(Replace(strPathPS, strPathDirProg, "%PROGRAMFILES%") & ";")) > 0
      ' Nothing
    Case Instr(UCase(strPathPSWin & ";"), UCase(strPathPS & ";")) > 0
      ' Nothing
    Case Else
      strPathPSWin  = strPathPS & ";" & strPathPSWin
  End Select

  colSysEnvVars("PSModulePath") = strPathPSWin

End Sub


Sub SetupStdVolumes()
  Call SetProcessId("1GB", "Setup Standard Volumes")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm,   "Folder1",    "\Scripts")
  Call SetXMLParm(objFolderParm,   "Folder2",    "\Servers")
  If strSetupSQLTools = "YES" Then
    Call SetXMLParm(objFolderParm, "Folder3",  "\SQL Server Management Studio\Custom reports")
  End If
  Call PrepareFolder("DBA",        objFolderParm)

  If strSetupDRUClt = "YES" Then
    Call SetXMLParm(objFolderParm, "Account1", strSqlAccount)
    Call SetXMLParm(objFolderParm, "Account2", strAgtAccount)
    Call SetXMLParm(objFolderParm, "Folder1",  "\DRU.Work")
    Call SetXMLParm(objFolderParm, "Folder2",  "\DRU.Result")
    Call PrepareFolder("DRU",      objFolderParm)
  End If

  If GetBuildfileValue("SetupManagementDW") = "YES" Then
    Call SetXMLParm(objFolderParm, "Account1", GetBuildfileValue("MDWAccount"))
    Call SetXMLParm(objFolderParm, "Account2", strAgtAccount)
    Call PrepareFolder("MDW",      objFolderParm)
  End If

  Select Case True
    Case strSetupSQLIS <> "YES"
      ' Nothing
    Case GetBuildfileValue("ActionSQLIS") = "ADDNODE"
      ' Nothing
    Case Else
      Call SetXMLParm(objFolderParm, "Account1",  strIsAccount)
      Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
      Call SetXMLParm(objFolderParm, "Folder1",   "\Packages")
      Call PrepareFolder("DataIS",   objFolderParm)
  End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLServer()
  Call SetProcessId("1GC", "Setup SQL Server Volumes")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Data",     objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Log",      objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("SysDB",    objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call SetXMLParm(objFolderParm, "Folder1",   "\AdHocBackup")
  Call SetXMLParm(objFolderParm, "Folder2",   "\Reports")
  Call PrepareFolder("Backup",   objFolderParm)
  If strActionSQLDB <> "ADDNODE" Then
    Call PrepareFolderPath("Backup", strAction, strDirSystemDataBackup, strSecNull, "", "")
    Call PrepareFolderPath("Backup", strAction, strDirSystemDataShared, strSecNull, "", "")
  End If

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("Temp",     objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("LogTemp",  objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("BPE",      objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("DataFS",   objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLDB)
  Call SetXMLParm(objFolderParm, "Account1",  strSqlAccount)
  Call PrepareFolder("DataFT",   objFolderParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupSQLASVolumes()
  Call SetProcessId("1GD", "Setup AS Service Volumes")
  Dim objFolderParm

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLAS)
  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("DataAS",   objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLAS)
  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("LogAS",    objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLAS)
  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call SetXMLParm(objFolderParm, "Folder1",   "\Data")
  Call SetXMLParm(objFolderParm, "Folder2",   "\AdHocBackup")
  Call PrepareFolder("BackupAS", objFolderParm)

  Call SetXMLParm(objFolderParm, "Action",    strActionSQLAS)
  Call SetXMLParm(objFolderParm, "Account1",  strAsAccount)
  Call SetXMLParm(objFolderParm, "Account2",  strAgtAccount)
  Call PrepareFolder("TempAS",   objFolderParm)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupTempVolume()
  Call SetProcessId("1GE", "Setup Temp folder Volume")

  Call PrepareFolderPath("TempWin", strAction, strPathTemp, strSecTemp, "", "")

  colSysEnvVars("TEMP") = strPathTemp
  colSysEnvVars("TMP")  = strPathTemp

  colUsrEnvVars("TEMP") = strPathTemp
  colUsrEnvVars("TMP")  = strPathTemp

  Call SetBuildFileValue("SetupTempWinStatus", strStatusProgress)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub PrepareFolder(strFType, objFolderParm)
  Call DebugLog("PrepareFolder: " & strFType)
  Dim arrVolumes
  Dim strDirBase, strDirName
  Dim strFAction, strFAccount1, strFAccount2, strFAccess, strFFolder1, strFFolder2, strFFolder3
  Dim strVolume, strVolList, strVolSource
  Dim intIdx, intVol
' VolSource: C=CSV, D=Disk, M=Mount Point, N=Mapped Network Drive, S=Share
' VolType:   C=Clustered, L=Local, X=Either

  strFAction        = GetXMLParm(objFolderParm, "Action",   "")
  strFAccount1      = GetXMLParm(objFolderParm, "Account1", "")
  strFAccount2      = GetXMLParm(objFolderParm, "Account2", "")
  strFAccess        = GetXMLParm(objFolderParm, "Access",   strSecNull)
  strFFolder1       = GetXMLParm(objFolderParm, "Folder1",   "")
  strFFolder2       = GetXMLParm(objFolderParm, "Folder2",   "")
  strFFolder3       = GetXMLParm(objFolderParm, "Folder3",   "")

  strDirName        = GetBuildfileValue("Dir" & strFType)
  strDirBase        = GetBuildfileValue("Dir" & strFType & "Base")
  strVolList        = GetBuildfileValue("Vol" & strFType)
  strVolSource      = GetBuildfileValue("Vol" & strFType & "Source")

  arrVolumes        = Split(Replace(strVolList, ",", " "))
  For intVol = 0 To UBound(arrVolumes)
    strVolume       = Trim(arrVolumes(intVol))
    Select Case True
      Case strDirName = ""
        ' Nothing
      Case strVolSource <> "D"
        strPath     = strVolume & strDirBase
        Call PrepareFolderPath(strFType, strFAction, strPath, strSecDBA, strFAccount1, strFAccount2)
        Call SetupAVExclude(strFType, strPath)
      Case Else
        For intIdx = 1 To Len(strVolume)
          strVol    = Mid(strVolume, intIdx, 1)
          strPath   = strVol & strDirBase
          Call PrepareFolderPath(strFType, strFAction, strPath, strSecDBA, strFAccount1, strFAccount2)
          Call SetupAVExclude(strFType, strPath)
        Next
    End Select
  Next

  If strFFolder1 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder1, strFAccess, "", "")
  End If

  If strFFolder2 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder2, strFAccess, "", "")
  End If

  If strFFolder3 <> "" Then
    Call PrepareFolderPath(strFType, strFAction, strDirName & strFFolder3, strFAccess, "", "")
  End If

  objFolderParm     = ""

End Sub


Sub PrepareFolderPath(strType, strAction, strPath, strSec, strAccount1, strAccount2)
  Call DebugLog("PrepareFolderPath: " & strPath)
  Dim strPathFolder

  strPrepareFolderPath = strPath
  Select Case True
    Case strAction <> "ADDNODE"
      ' Nothing
    Case GetBuildfileValue("Vol" & strType & "Source") <> "D"
      ' Nothing
    Case GetBuildfileValue("Vol" & strType & "Type") <> "L"
      Exit Sub
  End Select

  strPathFolder     =  SetupFolder(strPath, strSec)
  If strAccount1 <> "" Then
    strCmd          = """" & strPathFolder & """ /T /C /E /G """ & FormatAccount(strAccount1) & """:F"
    Call RunCacls(strCmd)
  End If
  If strAccount2 <> "" Then
    strCmd          = """" & strPathFolder & """ /T /C /E /G """ & FormatAccount(strAccount2) & """:F"
    Call RunCacls(strCmd)
  End If

End Sub


Function SetupFolder(strPath, strSec)
  Call DebugLog("SetupFolder: " & strPath)
  Dim strNull, strPathAlt, strPathAltParent, strPathFolder, strPathParent, strPathRoot

  If Right(strPath, 1) = "\" Then
    strPath         = Left(strPath, Len(strPath) - 1)
  End If

  strPathAlt        = strPath
  Select Case True
    Case Left(strPath, 2) <> "\\"
      ' Nothing
    Case Instr(3, strPath, "\") = 0
      SetupFolder   = strPath
      Exit Function
    Case Else
      strPathRoot   = Left(strPathAlt, Instr(3, strPathAlt, "\") - 1)
      strPathAlt    = strPathRoot & Mid(strPathRoot, 2) & Mid(strPathAlt, Instr(3, strPathAlt, "\")) ' For SOFS
  End Select
  strPathParent     = Left(strPath, InstrRev(strPath, "\") - 1)
  strPathAltParent  = Left(strPathAlt, InstrRev(strPathAlt, "\") - 1)

  strDebugMsg1      = "PathParent: " & strPathParent
  strPathFolder     = ""
  Select Case True
    Case objFSO.FolderExists(strPath & "\")
      strPathFolder = strPath
    Case objFSO.FolderExists(strPathAlt & "\")
      strPathFolder = strPathAlt
    Case objFSO.FolderExists(strPathParent & "\")
      strPathFolder = strPath
      Call CreateThisFolder(strPathFolder, strSec)
    Case objFSO.FolderExists(strPathAltParent & "\")
      strPathFolder = strPathAlt
      Call CreateThisFolder(strPathFolder, strSec)
    Case Else
      strNull       = SetupFolder(strPathParent, strSec)
      Select Case True
        Case objFSO.FolderExists(strPathParent & "\")
          strPathFolder = strPath
          Call CreateThisFolder(strPathFolder, strSec)
        Case objFSO.FolderExists(strPathAltParent & "\")
          strPathFolder = strPathAlt
          Call CreateThisFolder(strPathFolder, strSec)
      End Select
  End Select

  SetupFolder       = strPathFolder

End Function


Sub CreateThisFolder(strFolder, strSec)
  Call DebugLog("CreateThisFolder: " & strFolder)
  Dim strCreate

  strCreate         = "N"
  Select Case True
    Case objFSO.FolderExists(strFolder)
      ' Nothing
    Case Else
      objFSO.CreateFolder(strFolder)
      Wscript.Sleep strWaitShort
      strCreate     = "Y"
  End Select

  Select Case True
    Case strSec = ""
      ' Nothing
    Case (strSec = strSecDBA) And (strCreate = "N")
      ' Nothing
    Case strSec = strSecDBA
      strCmd        = """" & strFolder & """ /T /C /G " & strSec 
      Select Case True
        Case strGroupDBANonSA = ""
          ' Nothing
        Case strFolder = strDirDBA
          strCmd    = strCmd & " """ & FormatAccount(strGroupDBANonSA) & """:F "
        Case strFolder = strPathTemp
          strCmd    = strCmd & " """ & FormatAccount(strGroupDBANonSA) & """:F "
        Case Else
          strCmd    = strCmd & " """ & FormatAccount(strGroupDBANonSA) & """:R "
      End Select
      Call RunCacls(strCmd)
    Case strSec = strSecTemp
      strCmd        = """" & strFolder & """ /T /C /G " & strSec 
      Call RunCacls(strCmd)     
    Case Else
      Call ProcessAccounts("AssignFolderRights", strFolder)
  End Select

  strPrepareFolderPath = ""

End Sub


Sub SetupAVExclude(strType, strPath)
  Call DebugLog("SetupAVExclude: " & strType & " for " & strPath)

  Select Case True
    Case strOSVersion < "6.0"
      ' Nothing
    Case Else
      strCmd        = strAVCmd & """" & strPath & """"
      Call Util_RunExec(strCmd, "", "", -1)
  End Select

End Sub


Sub ProcessAccounts(strProcess, strParameter)
  Call DebugLog("ProcessAccounts: " & strProcess)
  Dim intDomIdx, strParm

  intDomIdx         = InStr(strNTAuthAccount, "\")

  If strParameter <> "" Then
    strParm         = ",""" & strParameter & """"
  End If

  Select Case True
    Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSQLAccount = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & strSQLAccount & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

  Call ProcessServiceAccount(strProcess, strSetupSQLDBAG,                     strAgtAccount,                        strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLAS,                       strASAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupDRUCtlr"),   strCtlrAccount,                       strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupDRUClt"),    strCltAccount,                        strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupSQLDB"),     strSQLBrowserAccount,                 strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupAnalytics"), GetBuildfileValue("ExtSvcAccount"),   strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLDBFT,                     strFTAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLIS,                       strIsAccount,                         strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupIsMaster"),  GetBuildfileValue("IsMasterAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupISWorker"),  GetBuildfileValue("IsWorkerAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupPolyBase"),  GetBuildfileValue("PBDMSSvcAccount"), strParm)
  Call ProcessServiceAccount(strProcess, GetBuildfileValue("SetupPolyBase"),  GetBuildfileValue("PBEngSvcAccount"), strParm)
  Call ProcessServiceAccount(strProcess, strSetupSQLRS,                       strRSAccount,                         strParm)

  Select Case True
    Case strGroupDBA = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & FormatAccount(strGroupDBA) & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & FormatAccount(strGroupDBANonSA) & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

End Sub


Sub ProcessServiceAccount(strProcess, strSetup, strAccount, strParm)
  Call DebugLog("ProcessServiceAccount: " & strAccount)
  Dim strMSAGroup

  Select Case True
    Case strSetup <> "YES"
      ' Nothing
    Case strAccount = ""
      ' Nothing
    Case strAccount = strSqlAccount
      ' Nothing
    Case Else
      strCmd        = strProcess & "(""" & strAccount & """" & strParm & ")"
      Execute "Call " & strCmd
  End Select

End Sub


Sub AssignUserGroups(strAccount)
  Call DebugLog("AssignUserGroups: " & strAccount)
  Dim intServerLen

  intServerLen      = Len(strServer) + 1

  Select Case True
    Case Left(strGroupDBA, intServerLen) = strServer & "\"
      ' Nothing
    Case strGroupDBA = strLocalAdmin
      ' Nothing
    Case Ucase(strAccount) = strGroupDBA
      strCmd        = "NET LOCALGROUP """ & strGroupRDUsers & """ """ & strAccount & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
      Call AssignAccountGroups(strAccount)
  End Select

  Select Case True
    Case strGroupDBANonSA = ""
      ' Nothing
    Case Left(strGroupDBANonSA, intServerLen) = strServer & "\"
      ' Nothing
    Case Ucase(strAccount) = strGroupDBANonSA
      strCmd        = "NET LOCALGROUP """ & strGroupRDUsers & """ """ & strAccount & """ /ADD"
      Call Util_RunExec(strCmd, "", strResponseYes, 2)
      Call AssignAccountGroups(strAccount)
  End Select

  Select Case True
    Case Ucase(strAccount) = strGroupDBA
      ' Nothing
    Case Ucase(strAccount) = strGroupDBANonSA
      ' Nothing
    Case Left(strAccount, intServerLen) = strServer & "\"
      ' Nothing
    Case Ucase(strAccount) = strNTAuthAccount 
      ' Nothing
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      Call AssignAccountGroups(strAccount)
  End Select

End Sub


Sub AssignAccountGroups(strAccount)
  Call DebugLog("AssignAccountGroups: " & strAccount)

  strCmd            = "NET LOCALGROUP """ & strGroupUsers & """ """ & strAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  strCmd            = "NET LOCALGROUP """ & strGroupDistComUsers & """ """ & strAccount & """ /ADD"
  Call Util_RunExec(strCmd, "", strResponseYes, 2)

  If strGroupPerfLogUsers <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupPerfLogUsers & """ """ & strAccount & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End If

  If strGroupPerfMonUsers <> "" Then
    strCmd          = "NET LOCALGROUP """ & strGroupPerfMonUsers & """ """ & strAccount & """ /ADD"
    Call Util_RunExec(strCmd, "", strResponseYes, 2)
  End If

End Sub


Sub AssignAccountRights(strAccount)
  Call DebugLog("AssignAccountRights: " & strAccount)

  Select Case True
    Case Ucase(strAccount) = Ucase(strGroupDBA)
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeManageVolumePrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeRemoteShutdownPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeShutdownPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeSystemProfilePrivilege")
    Case (Ucase(strAccount) = Ucase(strGroupDBANonSA)) And (strGroupDBANonSA <> "")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & FormatAccount(strAccount) & """ +r SeSystemProfilePrivilege")
    Case strAccount = strSqlAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")    ' Replace a process-level token
      Call RunNTRights("""" & strAccount & """ +r SeBatchLogonRight")                ' Log on as a Batch Job
      Call RunNTRights("""" & strAccount & """ +r SeCreateGlobalPrivilege")          ' Create Global objects
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")          ' Bypass traverse checking
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")           ' Impersonate a client after Authentication
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseBasePriorityPrivilege")  ' Adjust scheduling priority
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")         ' Adjust memory quotas
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseWorkingSetPrivilege")    ' Adjust Working Set
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")            ' Lock pages in memory
      Call RunNTRights("""" & strAccount & """ +r SeManageVolumePrivilege")          ' Manage files on a volume
      Call RunNTRights("""" & strAccount & """ +r SeProfileSingleProcessPrivilege")  ' Profile a process
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")              ' Log on as a Service
      Call RunNTRights("""" & strAccount & """ +r SeSystemProfilePrivilege")         ' Profile System performance
      Call RunNTRights("""" & strAccount & """ +r SeTcbPrivilege")                   ' Act as part of the Operating System
    Case strAccount = strAgtAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeBatchLogonRight")
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strFTAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strAsAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseBasePriorityPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseWorkingSetPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeLockMemoryPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
      Call RunNTRights("""" & strAccount & """ +r SeSystemProfilePrivilege")
    Case strAccount = strIsAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case strAccount = strRsAccount
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeImpersonatePrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeProfileSingleProcessPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
      Call RunNTRights("""" & strAccount & """ +r SeSystemProfilePrivilege")
    Case strAccount = strExtSvcAccount
      Call RunNTRights("""" & strAccount & """ +r SeAssignPrimaryTokenPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeIncreaseQuotaPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      Call RunNTRights("""" & strAccount & """ +r SeChangeNotifyPrivilege")
      Call RunNTRights("""" & strAccount & """ +r SeServiceLogonRight")
  End Select

End Sub


Sub AssignFolderRights(strAccount, strFolder)
  Call DebugLog("AssignFolderRights: " & strFolder)

  Select Case True
    Case strAccount = strGroupDBANonSA
      strCmd        = """" & strFolder & """ /T /C /E /G """ & FormatAccount(strAccount) & """:R "
    Case Left(strAccount, Len(strNTService) + 1) = strNTService & "\"
      ' Nothing
    Case Else
      strCmd        = """" & strFolder & """ /T /C /E /G """ & FormatAccount(strAccount) & """:F "
  End Select
  Call RunCacls(strCmd)
  
End Sub


Sub PostPreparation()
  Call SetProcessId("1H", "Post Preparation Tasks")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HA"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case Else
      Call SystemFolderPermissions()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "1HB"
      ' Nothing
    Case Else
      Call GPUpdate()
  End Select

  Call SetProcessId("1HZ", " Post Preparation Tasks" & strStatusComplete)
  Call ProcessEnd("")

End Sub 


Sub SystemFolderPermissions()
  Call SetProcessId("1HA", "Set System Folder Permissions")

  Call SetKB2811566Permissions() 

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetKB2811566Permissions()
  Call SetProcessId("1HAA", "Set KB2811566 Permissions")
  Dim strLogPath

  strLogPath        = strDirSys & "\system32\LogFiles\Sum"

  Select Case True
   Case strSetupSQLDB <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case Not objFSO.FolderExists(strLogPath)
      ' Nothing
    Case Else
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strSQLAccount) & """:R"
      Call RunCacls(strCmd)
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strSQLAccount) & """:W"
      Call RunCacls(strCmd)
  End Select

  Select Case True
    Case strSetupSQLAS <> "YES"
      ' Nothing
    Case strSQLVersion < "SQL2012"
      ' Nothing
    Case strOSVersion < "6.2"
      ' Nothing
    Case Not objFSO.FolderExists(strLogPath)
      ' Nothing
    Case Else
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strAsAccount) & """:R"
      Call RunCacls(strCmd)
      strCmd        = """" & strLogPath & """ /T /C /E /G """ & FormatAccount(strAsAccount) & """:W"
      Call RunCacls(strCmd)
  End Select

End Sub


Sub GPUpdate()
  Call SetProcessId("1HB", "Run GPUpdate to apply permissions")

  strCmd            = "HKLM\SOFTWARE\CLS\ITInfra\FFGPO_Update"
  Call Util_RegWrite(strCmd, 1, "REG_DWORD")

  strCmd            = "GPUPDATE /Target:Computer /Force"
  Call Util_RunExec(strCmd, "", strResponseNo, -1)

  Call ProcessEnd(strStatusComplete)

End Sub


Sub UserPreparation()
  Call SetProcessId("1U", "User Preparation Tasks")
  Dim objInstParm

  Call SetXMLParm(objInstParm, "PathMain",    strPathFBScripts)
  Call SetXMLParm(objInstParm, "ParmXtra",    GetBuildfileValue("FBParm"))
  Call RunInstall("UserPreparation", GetBuildfileValue("UserPreparationvbs"), objInstParm)

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


Sub RunNTRights(strCmd)
  Call DebugLog("RunNTRights: " & strCmd)

  Call Util_RunExec(strProgNtrights & " -u " & strCmd, "", strResponseYes, -1)
  Select Case True
    Case intErrSave = 0
      ' Nothing
    Case intErrSave = 2
      ' Nothing
    Case Else
      Call SetBuildMessage(strMsgError, "Error " & Cstr(intErrSave) & " " & strErrSave & " returned by " & strCmd)
  End Select

End Sub


Function GetPathLog()

  GetPathLog        = strSetupLog & strInstLog & strProcessIdLabel & " " & strProcessIdDesc & ".txt"""

End Function


End Class


Sub SetTempLoc(strKeyValue, strKey, strSid)
  Call DebugLog("SetTempLoc:")
  Dim objWMIReg
  Dim strCmd, strPath, strPathTempUser, strTempVar

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strPathTempUser   = GetBuildfileValue("PathTempUser")

  strPath           = strSid & "\Environment"
  objWMIReg.GetStringValue strKeyValue,strPath,"TEMP",strTempVar

  If Not IsNull(strTempVar) Then
    strCmd          = strKey & strPath & "\TEMP" 
    Call Util_RegWrite(strCmd, strPathTempUser, "REG_SZ")
    strCmd          = strKey & strPath & "\TMP" 
    Call Util_RegWrite(strCmd, strPathTempUser, "REG_SZ")
  End If

End Sub