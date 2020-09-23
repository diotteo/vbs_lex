''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  FineBuild6ConfigureUsers.vbs  
'  Copyright FineBuild Team © 2008 - 2020.  Distributed under Ms-Pl License
'
'  Purpose:      Setup User Preferences for SQL Server
'
'  Author:       Ed Vassie
'
'  Date:         02 Jul 2008
'
'  Change History
'  Version  Author        Date         Description
'  2.1.0    Ed Vassie     18 Jun 2010  Initial SQL Server 2008 R2 version
'  2.0.0    Ed Vassie     02 Jul 2008  Initial SQL Server 2008 version for FineBuild V1.2
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colsysEnvVars, colUsrEnvVars
Dim objExcList, objFile, objFolder, objFSO, objShell, objWMIReg
Dim intIdx, intTrust
Dim strAllUserProf, strAnyKey, strClusterInstall, strClusterName, strCmd, strCmdSQL, strDfltDoc, strDfltProf, strDfltRoot, strDefaultUser, strDirDBA, strDomain
Dim strEdition, strExcList, strRegSQL, strRegVS, strRegWin, strHKLMSQL, strInstance, strInstNode, strInstName
Dim strMainInstance, strMailServer, strMenuAccessories, strMenuPrograms, strMenuSQL, strMenuSSMS
Dim strNTAuthAccount, strNumLogins
Dim strOSName, strOSType, strOSVersion, strPath, strPathFB, strPathFBScripts, strPathNew, strPathOld, strPathTemp, strPathVS, strProfileName, strProcArc, strProfDir, strProgReg
Dim strRegItem, strResponseYes
Dim strServer, strSetupBIDS, strSetupSQLIS, strSetupSQLTools, strSQLVersion, strSQLVersionNum
Dim strType, strUserAccount, strUserName, strUserDNSDomain, strUserRoot, strUserSid, strVSVersionNum

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "6Z"
      ' Nothing
    Case Else
      Call SetupUsers()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "6Z"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("User setup complete.", 2, "User setup" ,64)
      Call FBLog("*")
      Call FBLog("* FineBuild completed successfully")
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
      Call FBLog(" SQL Server User Setup failed")
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
  Include "FBManageAccount.vbs"
  Include "FBManageInstall.vbs"
  Include "FBManageService.vbs"
  Call SetProcessIdCode("FB6C")

  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colSysEnvVars = objShell.Environment("System")
  Set colUsrEnvVars = objShell.Environment("User")

  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAllUserProf    = GetBuildfileValue("AllUserProf")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strClusterInstall = GetBuildfileValue("ClusterInstall")
  strClusterName    = GetBuildfileValue("ClusterName")
  strCmdSQL         = GetBuildfileValue("CmdSQL")
  strDefaultUser    = GetBuildfileValue("DefaultUser")
  strDfltDoc        = GetBuildfileValue("DfltDoc")
  strDfltProf       = GetBuildfileValue("DfltProf")
  strDfltRoot       = GetBuildfileValue("DfltRoot")
  strDirDBA         = GetBuildfileValue("DirDBA")
  strDomain         = GetBuildfileValue("Domain")
  strEdition        = GetBuildfileValue("AuditEdition")
  strInstance       = GetBuildfileValue("Instance")
  strInstNode       = GetBuildfileValue("InstNode")
  strMainInstance   = GetBuildfileValue("MainInstance")
  strMailServer     = GetBuildfileValue("MailServer")
  strMenuAccessories  = GetBuildfileValue("MenuAccessories")
  strMenuPrograms   = GetBuildfileValue("MenuPrograms")
  strMenuSQL        = GetBuildfileValue("MenuSQL")
  strMenuSSMS       = GetBuildfileValue("MenuSSMS")
  strNTAuthAccount  = GetBuildfileValue("NTAuthAccount")
  strNumLogins      = GetBuildfileValue("NumLogins")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathFBScripts  = GetBuildfileValue("PathFBScripts")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strPathVS         = GetBuildfileValue("PathVS")
  strProfileName    = GetBuildfileValue("ProfileName")
  strProcArc        = GetBuildfileValue("ProcArc")
  strProfDir        = GetBuildfileValue("ProfDir")
  strProgReg        = GetBuildfileValue("ProgReg")
  strRegWin         = "\Software\Microsoft\Windows\CurrentVersion\"
  strResponseYes    = GetBuildfileValue("ResponseYes")
  strServer         = GetBuildfileValue("AuditServer")
  strSetupBIDS      = GetBuildfileValue("SetupBIDS")
  strSetupSQLIS     = GetBuildfileValue("SetupSQLIS")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strSQLVersionNum  = GetBuildfileValue("SQLVersionNum")
  strType           = GetBuildfileValue("Type")
  strUserAccount    = GetBuildfileValue("UserAccount")
  strUserDNSDomain  = GetBuildfileValue("UserDNSDomain")
  strUserName       = GetBuildfileValue("AuditUser")
  strVSVersionNum   = GetBuildfileValue("VSVersionNum")
  strRegSQL         = "\Software\Microsoft\Microsoft SQL Server\" & strSQLVersionNum & "\"
  strRegVS          = "\Software\Microsoft\VisualStudio\" & strVSVersionNum & "\"

  strExcList        = strPathTemp & "\ExcludeList.txt"
  Call SetBuildfileValue("ExcList", strExcList)
  If objFSO.FileExists(strExcList) Then
    objFSO.DeleteFile(strExcList)
  End If
  Set objExcList    = objFSO.CreatetextFile(strExcList)
  objExcList.Writeline("My Music")
  objExcList.Writeline("My Pictures")
  objExcList.Writeline("My Videos")
  objExcList.Close 

End Sub


Sub SetupUsers()
  Call SetProcessId("6", "User Setup processing (FineBuild6ConfigureUsers.vbs)")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AZ"
      ' Nothing
    Case Else
      SetupUserRegistry()
  End Select

  Dim strSetupMAPIProfile
  strSetupMAPIProfile   = GetBuildfileValue("SetupMAPIProfile")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6B"
      ' Nothing
    Case strEdition = "EXPRESS"
      ' Nothing
    Case strSetupMAPIProfile <> "YES"
      ' Nothing
    Case strInstName <> ""
      ' Nothing
    Case Else
      Call ConfigSQLMail()
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6CZ"
      ' Nothing
    Case Else
      Call FineBuildEnd()
  End Select

  Call SetProcessId("6Z", " User Setup processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub SetupUserRegistry()
  Call SetProcessId("6A", "User Registry processing")

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AA"
      ' Nothing
    Case strSetupSQLTools <> "YES"
      ' Nothing
    Case GetBuildfileValue("SetupBOL") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AA", "BOL preferences", "ProcessBol")
      Call SetBuildfileValue("SetupBOLStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AB"
      ' Nothing
    Case GetBuildfileValue("SetupCMD") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AB", "Command Window preferences", "ProcessCmd")
      Call SetBuildfileValue("SetupCMDStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AC"
      ' Nothing
    Case GetBuildfileValue("SetupSSMS") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AC", "SQL Server Management Studio preferences", "ProcessSSMS")
      Call SetBuildfileValue("SetupSSMSStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AD"
      ' Nothing
    Case GetBuildfileValue("SetupVS") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AD", "Configure Visual Studio Preferences", "ProcessVS")
      Call SetBuildfileValue("SetupVSStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AE"
      ' Nothing
    Case GetBuildfileValue("SetupNetTrust") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AE", "Configure Internet Preferences", "ProcessNetTrust")
      Call SetBuildfileValue("SetupNetTrustStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AF"
      ' Nothing
    Case GetBuildfileValue("SetupWindows") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AF", "Configure Windows Preferences", "ProcessWindows")
      Call SetBuildfileValue("SetupWindowsStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AG"
      ' Nothing
    Case GetBuildfileValue("SetupWindows") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AG", "Windows Menus", "ProcessMenus")
      Call SetBuildfileValue("SetupWindowsStatus", strStatusComplete)
  End Select

  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6AH"
      ' Nothing
    Case GetBuildfileValue("SetupMyDocs") <> "YES"
      ' Nothing
    Case Else
      Call ProcessUser("6AH", "Configure 'My Documents' location", "ProcessMyDocs")
      Call SetBuildfileValue("SetupMyDocsStatus", strStatusComplete)
  End Select

  Call SetProcessId("6AZ", " User Registry processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub ConfigSQLMail()
  Call SetProcessId("6B", "SQLMail Profile")
  Dim strServAcnt

  strPath           = "SYSTEM\CurrentControlSet\Services\MSSQLSERVER\"
  objWMIReg.GetStringValue strHKLM,strPath,"ObjectName",strServAcnt
  If IsNull(strServAcnt) Then
    strServAcnt     = ""
  End If
  If Left(strServAcnt, 2) = ".\" Then
    strServAcnt     = strServer & Mid(strServAcnt, 2)
  End If

  intIdx            = InStr(strServAcnt, "\")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strUserDNSDomain = ""
      Call SetBuildMessage(strMsgInfo, "Unable to create SQLMail profile: SQL is not running in a domain")
    Case intIdx = 0
      Call SetBuildMessage(strMsgInfo, "Unable to create SQLMail profile: SQL is not running using a domain account")
    Case Left(strServAcnt, intIdx) = Left(strNTAuthAccount, InStr(strNTAuthAccount, "\"))
      Call SetBuildMessage(strMsgInfo, "Unable to create SQLMail profile: SQL is not running using a domain account")
    Case strMailServer = ""
      Call SetBuildMessage(strMsgInfo, "Unable to create SQLMail profile: Mail Server can not be found")
    Case GetBuildfileValue("MailServertype") <> "E"
      Call SetBuildMessage(strMsgInfo, "Unable to create SQLMail profile: Mail Server type is not Exchange")
    Case Ucase(strUserName) <> Ucase(Mid(strServAcnt, intIdx + 1))
      ' Nothing - Current user is not the Service Account
    Case Else
      Call SetupMAPIProfile(strServAcnt, strMailServer)
   End Select

  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetupMAPIProfile(strServAcnt, strMailServer)
  Call DebugLog("SetupMAPIProfile: " & strServAcnt & " for " & strMailServer)
  Dim strMailAcnt, strNewData, strOldData

  strMailAcnt       = Mid(strServAcnt, intIdx + 1) & "@" & strUserDNSDomain

  Call DebugLog("Copy Mail Profile file to \Temp folder")
  Set objFile       = objFSO.GetFile(strPathFBScripts & "Set-OutlookProfile.PRF")
  strPathNew        = objShell.ExpandEnvironmentStrings("%TEMP%") & "\" & objFile.Name
  objFile.Copy strPathNew

  Call DebugLog("Tailor Mail Profile file")
  Set objFile       = objFSO.OpenTextFile(strPathNew, ForReading)
  strOldData        = objFile.ReadAll
  objFile.Close

  strNewData        = Replace(strOldData, "%MAILACCOUNT%", strMailAcnt,    1, -1 , 1) ' 1=Start Pos;-1=Replace all;1=Ignore case
  strOldData        = Replace(strNewData, "%MAILSERVER%",  strMailServer,  1, -1 , 1) ' 1=Start Pos;-1=Replace all;1=Ignore case

  Set objFile      = objFSO.OpenTextFile(strPathNew, ForWriting)
  objFile.WriteLine strOldData
  objFile.Close

  Call DebugLog("Start Outlook to create Mail Profile")
  err.Number        = objShell.Run("OUTLOOK.EXE /importprf " & strPathNew,7,False)
  err.Number        = objShell.Popup("Outlook has been started to create the Mail Profile.  Outlook can be closed whenever convenient", 2, "Outlook Mail Profile Setup" ,64)

End Sub


Sub FineBuildEnd()
  Call SetProcessId("6C", "FineBuild Completion Processing")

  Dim strSetupOldAccounts
  strSetupOldAccounts = GetBuildfileValue("SetupOldAccounts")
  Select Case True
    Case err.Number <> 0
      ' Nothing
    Case strProcessId > "6CA"
      ' Nothing
    Case strSetupOldAccounts <> "YES"
      ' Nothing
    Case Else
      Call DropInstallLogin()
  End Select

  Call SetProcessId("6CZ", " FineBuild Completion Processing" & strStatusComplete)
  Call ProcessEnd("")

End Sub


Sub DropInstallLogin()
  Call SetProcessId("6CA", "Drop Install Login")

  Select Case True
    Case strType = "WORKSTATION"
      ' Nothing
    Case GetBuildfileValue("IsInstallDBA") = 1
      Call Util_ExecSQL(strCmdSQL & "-Q", """DROP LOGIN [" & strUserAccount & "]""", 1)
    Case strClusterInstall = "YES"
      ' Nothing
    Case Else
      Call Util_ExecSQL(strCmdSQL & "-Q", """ALTER LOGIN [" & strUserAccount & "] DISABLE""", 0)
  End Select

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


Sub ProcessBol(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessBol: 6AA")
  Dim objWMIReg
  Dim strPath, strRegSQL, strRegItem

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strRegSQL         = GetBuildfileValue("RegSQL")

  strPath           = strSid & strRegSQL & "Tools\Shell\Help"
  objWMIReg.GetStringValue strKeyValue,strPath,"RegToken",strRegItem

  If Not IsNull(strRegItem) Then
    Call Util_RegWrite(strKey & strPath & "\UseOnlineContent", 0, "REG_DWORD")				                              ' No online content
  End If

  strPath           = strSid & "\Software\Microsoft\HTLMHelp\1.x\ItssRestrictions"                                                     ' Show Compiled HTML (.chm) file contents
  Call Util_RegWrite(strKey & strPath & "\MaxAllowedZone", 2, "REG_DWORD")				                              ' 0 - My Computer, 1 - Local Intranet, 2 - Trusted Sites, 3 - Internet Zone, 4 - Restricted Zone

End Sub


Sub ProcessCmd(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessCmd: 6AB")
  Dim objWMIReg
  Dim strPath, strRegItem

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

  strPath           = strSid & "Console"
  objWMIReg.GetDWORDValue strKeyValue,strPath,"ScreenBufferSize",strRegItem

  If Not IsNull(strRegItem) Then
    Call Util_RegWrite(strKey & strPath & "\HistoryBufferSize", 100, "REG_DWORD")                                                     ' Number of characters in command history buffer
    Call Util_RegWrite(strKey & strPath & "\HistoryNoDup", 1, "REG_DWORD")                                                            ' Drop duplicate command history buffers
    Call Util_RegWrite(strKey & strPath & "\NumberOfHistoryBuffers", 10, "REG_DWORD")                                                 ' Number of commands to store
    Call Util_RegWrite(strKey & strPath & "\QuickEdit", 1, "REG_DWORD")                                                               ' Enable Quick Edit mode
    Call Util_RegWrite(strKey & strPath & "\ScreenBufferSize", 196608080, "REG_DWORD")                                                ' Number of characters in screen buffer (3000 lines)
  End If

End Sub


Sub ProcessSSMS(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessSSMS: 6AC")
  Dim objWMIReg
  Dim strPath, strRegItem, strRegSQL

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strRegSQL         = GetBuildfileValue("RegSQL")

  strPath           = strSid & strRegSQL & "Tools\Shell\Help"
  objWMIReg.GetStringValue strKeyValue,strPath,"RegToken",strRegItem

  If Not IsNull(strRegItem) Then
    Call Util_RegWrite(strKey & strSid & strRegSQL & "CustomerFeedback", 0, "REG_DWORD")                                              ' No customer feedback
    Call Util_RegWrite(strKey & strPath & "\OnlineF1DialogShown", 1, "REG_DWORD")                                                     ' Mark local Help dialogue as displayed
    Call Util_RegWrite(strKey & strPath & "\UseLocalHelpF1", 1, "REG_DWORD")                                                          ' Always use local Help
    Call Util_RegWrite(strKey & strPath & "\UseMSDNOnlineF1", 0, "REG_DWORD")                                                         ' No online content
    Call Util_RegWrite(strKey & strPath & "\UseMSDNOnlineF1First", 0, "REG_DWORD")                                                    ' No online content
    Call Util_RegWrite(strKey & strSid & strRegSQL & "Tools\Client\SQLiMailWizard\ShowTitlePageInConfigureWizard", 1, "REG_DWORD")    ' No splash screen for DB Mail Wizard
  End If

End Sub


Sub ProcessVS(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessVS: 6AD")
  Dim objWMIReg
  Dim strPath, strRegItem, strRegVS, strSetupBIDS

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strRegVS          = GetBuildfileValue("RegVS")
  strSetupBIDS      = GetBuildfileValue("SetupBIDS")

  strPath           = strSid & strRegVS
  objWMIReg.GetStringValue strKeyValue,strPath & "Help","UseLocalHelpF1",strRegItem

  If Not IsNull(strRegItem) Then
    Call Util_RegWrite(strKey & strPath & "General\StartPage\IsDownloadRefreshEnabled", 0, "REG_DWORD")                               ' Disable automatic content downloads
    Call Util_RegWrite(strKey & strPath & "Help\OnlineF1DialogShown", 1, "REG_DWORD")					                              ' Mark local Help dialogue as displayed
    Call Util_RegWrite(strKey & strPath & "Help\UseLocalHelpF1", 1, "REG_DWORD")						                              ' Always use local Help
    Call Util_RegWrite(strKey & strPath & "Help\UseMSDNOnlineF1", 0, "REG_DWORD")						                              ' No online content
    Call Util_RegWrite(strKey & strPath & "Help\UseMSDNOnlineF1First", 0, "REG_DWORD")					                              ' No online content
    Call Util_RegWrite(strKey & strPath & "Help\UseOnlineContent", 0, "REG_DWORD")						                              ' No online content

    If strSetupBIDS = "YES" Then
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolNumKeys", 1, "REG_DWORD")					                          ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolArg0", "", "REG_SZ")						                          ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolCmd0", strPathVS & "Tools\" & "RawFileReader.exe", "REG_SZ")	      ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolDir0", "", "REG_SZ")						                          ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolOpt0", 17, "REG_DWORD")					                              ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolSourceKey0", "", "REG_SZ")					                          ' Raw File Reader
      Call Util_RegWrite(strKey & strPath & "External Tools\ToolTitle0", "SSIS Raw File Reader", "REG_SZ")		             	      ' Raw File Reader
    End If

  End If

End Sub


Sub ProcessNetTrust(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessNetTrust: 6AE")
  Dim objWMIReg
  Dim intTrust
  Dim strPath, strRegItem, strRegWin

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strRegWin         = GetBuildfileValue("RegWin")

  strPath           = strSid & strRegWin
  objWMIReg.GetStringValue strKeyValue,strPath & "Internet Settings", "CertificateRevocation",strRegItem

  If Not IsNull(strRegItem) Then
    intTrust        = objShell.RegRead(strKey & strPath & "WinTrust\Trust Providers\Software Publishing\State")	                      ' Get Publisher Certificate Check state
    intTrust        = intTrust Or &H200
    Call Util_RegWrite(strKey & strPath & "WinTrust\Trust Providers\Software Publishing\State", intTrust, "REG_DWORD")	              ' Do not check for Publisher Certificate Check 
    Call Util_RegWrite(strKey & strPath & "Internet Settings\CertificateRevocation", 0, "REG_DWORD")			                      ' Do not check for server certificate revocation
    Call Util_RegWrite(strKey & strPath & "Internet Settings\WarnonBadCertRecving", 0, "REG_DWORD")				                      ' Do not warn about invalid site certificates
  End If

End Sub


Sub ProcessWindows(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessWindows: 6AF")
  Dim objWMIReg
  Dim strPath, strRegItem, strRegWin

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strRegWin         = GetBuildfileValue("RegWin")

  strPath           = strSid & strRegWin & "Explorer"
  objWMIReg.GetStringValue strKeyValue,strPath & "\Advanced", "HideFileExt",strRegItem

  If Not IsNull(strRegItem) Then
    Call Util_RegWrite(strKey & strSid & "\Control Panel\Mouse\SnapToDefaultButton", 1, "REG_SZ")				                      ' Snap mouse pointer to default button in dialogues
    Call Util_RegWrite(strKey & strPath & "\Advanced\TaskBarSizeMove", 0, "REG_DWORD")				                                  ' Lock Taskbar
    Call Util_RegWrite(strKey & strPath & "\Advanced\IntelliMenus", 0, "REG_DWORD")					                                  ' Do not use Personalised Menus
    Call Util_RegWrite(strKey & strPath & "\Advanced\CascadeControlPanel", "YES", "REG_SZ")				                              ' Show Control Panel as Menu	
    Call Util_RegWrite(strKey & strPath & "\CabinetState\Fullpath", 0, "REG_DWORD")					                                  ' Do not display full path in Title Bar
    Call Util_RegWrite(strKey & strPath & "\Advanced\NoNetCrawling", 1, "REG_DWORD")				                                  ' Do not search for Network folders
    Call Util_RegWrite(strKey & strPath & "\Advanced\HideFileExt", 0, "REG_DWORD")					                                  ' Show file extensions
    Call Util_RegWrite(strKey & strPath & "\Advanced\Hidden", 1, "REG_DWORD")					                                      ' Show hidden files
    Call Util_RegWrite(strKey & strPath & "\Advanced\ShowSuperHidden", 1, "REG_DWORD")				                                  ' Show Operating system files
    Call Util_RegWrite("HKCR\Folder\shell\", "explore", "REG_SZ")	                                                                  ' Default mode for Windows Explorer is Explore
  End If

End Sub


Sub ProcessMenus(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessMenus: 6AG")
  Dim objFile, objFSO, objWMIReg
  Dim strAllUserProf, strDfltProf, strDfltRoot, strMenuAccessories, strMenuPrograms, strMenuSQL, strMenuSSMS, strPath, strPathNew, strOSVersion, strPathOld, strRegWin

  Set objFSO        = CreateObject("Scripting.FileSystemObject")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strAllUserProf    = GetBuildfileValue("AllUserProf")
  strDfltProf       = GetBuildfileValue("DfltProf")
  strDfltRoot       = GetBuildfileValue("DfltRoot")
  strMenuAccessories  = GetBuildfileValue("MenuAccessories")
  strMenuPrograms   = GetBuildfileValue("MenuPrograms")
  strMenuSQL        = GetBuildfileValue("MenuSQL")
  strMenuSSMS       = GetBuildfileValue("MenuSSMS")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strRegWin         = GetBuildfileValue("RegWin")

  Select Case True
    Case Ucase(strSid) = ".DEFAULT"
      strPath       = strDfltRoot
    Case Else
      strPathOld    = strSid & strRegWin & "Explorer\Shell Folders"
      strDebugMsg1  = "Source: " & strPathOld
      objWMIReg.GetStringValue strKeyValue, strPathold,"AppData",strPath
      If IsNull(strPath ) Then
        Exit Sub
      End If
      If Not objFSO.FolderExists(strPath) Then
        Exit Sub
      End If
  End Select

  strPathOld        = strAllUserProf & "\" & strMenuPrograms & "\" & strMenuSQL & "\" & strMenuSSMS & ".lnk"
  strDebugMsg1      = "Source: " & strPathOld
  If objFSO.FileExists(strPathOld) Then
    Set objFile     = objFSO.GetFile(strPathOld)
    strPathNew      = strPath & "\Microsoft\Internet Explorer"
    strDebugMsg2    = "Target: " & strPathNew
    If Not objFSO.FolderExists(strPathNew) Then
      objFSO.CreateFolder(strPathNew)
    End If
    strPathNew      = strPathNew & "\Quick Launch"
    strDebugMsg2    = "Target: " & strPathNew
    If Not objFSO.FolderExists(strPathNew) Then
      objFSO.CreateFolder(strPathNew)
    End If
    strPathNew      = strPathNew & "\" & objFile.Name
    objFile.Copy strPathNew
  End If

  Select Case True
    Case strOSVersion <= "6.1"
      strPathOld    = strDfltProf & "\" & strMenuPrograms & "\" & strMenuAccessories & "\Windows Explorer.lnk"
      strDebugMsg1  = "Source: " & strPathOld
      Set objFile   = objFSO.GetFile(strPathOld)
      strPathNew    = strPath & "\Microsoft\Internet Explorer\Quick Launch\" & objFile.Name
      strDebugMsg2  = "Target: " & strPathNew
      objFile.Copy strPathNew
    Case Else
      strPathOld    = strDfltProf & "\" & strMenuPrograms & "\" & strMenuAccessories & "\File Explorer.lnk"
  End Select

End Sub


Sub ProcessMyDocs(strKeyValue, strKey, strSid)
  Call DebugLog("ProcessMyDocs: 6AH")
  Dim objWMIReg
  Dim strDfltDoc, strExcList, strPath, strPathNew, strPathOld, strRegItem, strRegWin

  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  strDfltDoc        = GetBuildfileValue("DfltDoc")
  strRegWin         = GetBuildfileValue("RegWin")
  strExcList        = GetBuildfileValue("ExcList")

  Select Case True
    Case Ucase(strSid) = ".DEFAULT"
      strPath       = strDfltDoc
    Case Else
      strPathOld    = strSid & strRegWin & "Explorer\Shell Folders"
      objWMIReg.GetStringValue strKeyValue, strPathOld,"Personal",strPath
      If IsNull(strPath) Then
        Exit Sub
      End If
      If Not objFSO.FolderExists(strPath) Then
        Exit Sub
      End If
  End Select
  strPathNew        = strDirDBA
  If Ucase(strPath) <> Ucase(strPathNew) Then
    strCmd          = "XCOPY """ & strPath & "\*.*"" """ & strPathNew & "\*.*"" /C /EXCLUDE:" & strExcList & " /E /V /H /R /Y"
    Call Util_RunExec(strCmd, "", "", 0)
  End If

  Call DebugLog("Ensure 'My Documents' location is saved")
  strPath           = strKey & strSid & strRegWin & "Explorer\Shell Folders\Personal"
  Call Util_RegWrite(strPath, strPathNew, "REG_EXPAND_SZ")
  strPath           = strKey & strSid & strRegWin & "Explorer\User Shell Folders\Personal"
  Call Util_RegWrite(strPath, strPathNew, "REG_EXPAND_SZ")

End Sub