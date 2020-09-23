''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  User1Preparation.vbs  
'  Copyright FineBuild Team © 2008 - 2018.  Distributed under Ms-Pl License
'
'  Purpose:      Script to perform user processing requirements prior to running the SQL Server install.
'                This script should be changed as needed to perform the processing required.
'
'                It is suggested that this script is used only for work that must take place before
'                SQL Server 2008 is installed. E.G. Initial cluster configuration.
'                
'                This script can use ProcessId values 1YA to 1ZZ
'
'  Author:       Ed Vassie
'
'  Change History
'  Version  Author        Date         Description
'  2.0      Ed Vassie     15 Oct 2008  Rewritten for FineBuild 2.0
'  1.0      Ed Vassie     01 Jan 2008  Initial version
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit
Dim SQLBuild : Set SQLBuild = New FineBuild

Class FineBuild

Dim colArgs, colBuild, colFiles, colFlags, colGlobal, colStrings, colPrcEnvVars, colVolEnvVars
Dim intIndex, intprocNum
Dim objApp, objAutoUpdate, objConfig, objDrive, objFSO, objShell, objWMI, objWMIREG
Dim strAnyKey, strBuildShares, strPathFB, strProgCacls, strConfig, strCmd
Dim strDirBackup, strDirData, strDirDataFT, strDirProg, strDirLog, strDirTempData, strDirTempLog, strDirSQL, strDrive, strDrives, strDrvProg, strDrvSys, strDrvUsed
Dim strEdition, strFilePerm, strHKLMFB, strHKLMSQL, strInstance, strInstAgent, strInstAS, strInstIS, strInstNode, strInstRS, strInstSQL
Dim strSetupSQLAS, strSetupSQLDB, strSetupSQLRS, strSetupSQLTools, strMainInstance, strOSName, strOSType, strOSVersion
Dim strSetupLog, strPath, strPathSys, strPathTemp, strProcArc, strServer, strServInst, strSetupShares, strSQLVersion, strStopAt, strType, strUserName, strXMLNode

Private Sub Class_Initialize
' Perform FineBuild processing

  err.Clear
  Call Initialisation()

  Select Case True
    Case err.Number <> 0 
      ' Nothing
    Case strProcessId >= "1ZZ"
      ' Nothing
    Case Else
      Call ProcessUserConfig()
  End Select

End Sub


Private Sub Class_Terminate
' Error handling and termination

  Select Case True
    Case strPathFB = "%SQLFBFOLDER%"
      ' Nothing
    Case strProcessId > "1ZZ"
      ' Nothing
    Case err.Number = 0 
      Call objShell.Popup("User Preparation complete", 2, "Instance Preparation" ,64)
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
      Call FBLog(" User Configuration failed")
    End Select

  Wscript.quit(err.Number)

End Sub


Sub Initialisation()
' Perform initialisation processing

  Set objShell      = WScript.CreateObject ("Wscript.Shell")
  strPathFB         = objShell.ExpandEnvironmentStrings("%SQLFBFOLDER%")
  Include "FBManageBuildfile.vbs"
  Include "FBManageLog.vbs"
  Include "FBUtils.vbs"
  Call SetProcessIdCode("FBU1")

  Set objApp        = CreateObject ("Shell.Application")
  Set objAutoUpdate = CreateObject("Microsoft.Update.AutoUpdate")
  Set objConfig     = CreateObject ("Microsoft.XMLDOM") 
  Set objFSO        = CreateObject ("Scripting.FileSystemObject")
  Set objWMI        = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
  Set objWMIReg     = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
  Set colArgs       = Wscript.Arguments.Named
  Set colPrcEnvVars = objShell.Environment("Process")

  strSQLVersion     = GetBuildfileValue("AuditVersion")
  strType           = GetBuildfileValue("Type")
  strXMLNode        = GetBuildfileValue("TypeNode")
  strConfig         = strPathFB & "\" & GetBuildfileValue("Config")
  strDebugMsg1      = "Config: " & strConfig
  objConfig.async   = "false"
  objConfig.load(strConfig)
  Set colGlobal     = objConfig.documentElement.selectSingleNode("Global")
  Set colBuild      = objConfig.documentElement.selectSingleNode(strXMLNode)
  Set colFiles      = objConfig.documentElement.selectSingleNode("Files")
  Set colFlags      = objConfig.documentElement.selectSingleNode(strXMLNode & "/Flags")
  Set colStrings    = objConfig.documentElement.selectSingleNode("Global/Strings")

  strHKLMSQL        = GetBuildfileValue("HKLMSQL")
  strAnyKey         = GetBuildfileValue("AnyKey")
  strDirProg        = GetBuildfileValue("DirProg")
  strDrives         = GetBuildfileValue("DrvList")
  strDrvSys         = GetBuildfileValue("DrvSys")
  strDrvProg        = GetBuildfileValue("DrvProg")
  strEdition        = GetBuildfileValue("AuditEdition")
  strFilePerm       = GetBuildfileValue("FilePerm")
  strInstance       = GetBuildfileValue("Instance")
  strInstNode       = GetBuildfileValue("InstNode")
  strInstAgent      = GetBuildfileValue("InstAgent")
  strInstAS         = GetBuildfileValue("InstAS")
  strInstIS         = GetBuildfileValue("InstIS")
  strInstRS         = GetBuildfileValue("InstRS")
  strInstSQL        = GetBuildfileValue("InstSQL")
  strMainInstance   = GetBuildfileValue("MainInstance")
  strOSName         = GetBuildfileValue("OSName")
  strOSType         = GetBuildfileValue("OSType")
  strOSVersion      = GetBuildfileValue("OSVersion")
  strPathTemp       = GetBuildfileValue("PathTemp")
  strProcArc        = GetBuildfileValue("ProcArc")
  intProcNum        = GetBuildfileValue("ProcNum")
  strProgCacls      = GetBuildfileValue("ProgCacls")
  strPathSys        = GetBuildfileValue("PathSys")
  strServer         = GetBuildfileValue("AuditServer")
  strServInst       = GetBuildfileValue("ServInst")
  strSetupSQLAS     = GetBuildfileValue("SetupSQLAS")
  strSetupSQLDB     = GetBuildfileValue("SetupSQLDB")
  strSetupSQLRS     = GetBuildfileValue("SetupSQLRS")
  strSetupSQLTools  = GetBuildfileValue("SetupSQLTools")
  strStopAt         = GetBuildfileValue("StopAt")
  strSetupShares    = GetBuildfileValue("SetupShares")
  strSetupLog       = Ucase(objShell.ExpandEnvironmentStrings("%SQLLOGTXT%"))
  strUserName       = GetBuildfileValue("AuditUser")

  If strSetupSQLDB = "YES" Then
    strDirBackup    = GetBuildfileValue("DirBackup")
    strDirData      = GetBuildfileValue("DirData")
    strDirDataFT    = GetBuildfileValue("DirDataFT")
    strDirLog       = GetBuildfileValue("DirLog")
    strDirSQL       = GetBuildfileValue("DirSQL")
    strDirTempData  = GetBuildfileValue("DirTemp") & "\Tempdb"
    strDirTempLog   = GetBuildfileValue("DirLog")
  End If

End Sub


Function GetParam(colParam, strParam, strDefault) 
' Get parameter value
  Dim strValue

' Find parameter value in XML configuration file
  Select Case True
    Case IsNull(colParam)
      strValue      = strDefault
    Case IsNull(colParam.getAttribute(strParam))
      strValue      = strDefault
    Case Else
      strValue      = colParam.getAttribute(strParam)
  End Select

' Apply any parameter overide from CSCRIPT arguments
  Select Case True
    Case Not colArgs.Exists(strParam)
      ' Nothing
    Case Else
      strValue      = colArgs.Item(strParam)
  End Select

  GetParam          = strValue

End Function


Sub ProcessUserConfig()
  Call SetProcessId("1U", "User Preparation processing (User1Preparation.vbs)")

  Call SetUpdate("ON")

' Main body of user processing is inserted here

  Call SetUpdate("OFF")
  Call SetProcessId("1UZ", " User Preparation processing" & strStatusComplete)
  Call ProcessEnd(strStatusComplete)

End Sub


Sub SetUpdate(strOnOff)
  Call DebugLog("SetUpdate: messages " & strOnOff)
  On Error Resume Next

  If strOnOff = "ON" Then
    colPrcEnvVars("SEE_MASK_NOZONECHECKS") = 1    ' Prevent Security Warning message hanging quiet install
    err.Number      = objAutoUpdate.Pause()       ' Prevent Windows Update service triggering a reboot prompt
  Else
    colPrcEnvVars.Remove("SEE_MASK_NOZONECHECKS") ' Allow Security Warning messages
    err.Number      =  objAutoUpdate.Resume()     ' Resume normal Window Update Service prompts
  End If

  Select Case True
    Case err.Number = 0
      ' No action
    Case err.Number < 0
      Call FBLog("Error " & Cstr(err.Number) & " returned by Windows Update Service when setting service to " & strOnOff & ".  This is for information only, processing is continuing.")
      err.Number    = 0
    Case Else
      err.Raise err.Number, "", "(" & strProcessIdLabel & ") " & "Error running Windows Update configuration " & strOnOff
  End Select

End Sub


Sub SetupDrive (strInstance, strUserDrives, strUserLabel)
  Call SetProcessId("1YA", "Setup User drives")
' Example code to set up a drive.  Change as required

  Call DebugLog("Setup User drive(s)")
  For intIndex = 1 To Len(strUserDrives)
    strDrive        = Mid(strUserDrives, intIndex, 1)
    Select Case True
      Case Instr(strDrives, strDrive) = 0
        ' No action, not a valid drive letter
      Case (intIndex = 1) Or (objFSO.FolderExists(strDrive & ":\"))
        Call SetupThisDrive(strDrive, strUserLabel)

        CreateThisFolder strDrive & strDirSQL, strSecMain
        CreateThisFolder strDrive & strDirSQL & "\" & strInstNode & ".Data", strSecNull

        If intIndex = 1 Then
          strCmd    = "HKLM\SOFTWARE\FineBuild\" & strInstNode & "\DirUser"
          Call Util_RegWrite(strCmd, strDrive & strDirSQL & "\" & strInstNode & ".Data", "REG_SZ")
          CreateThisFolder strDrive & strDirSQL & "\" & strInstNode & ".Data\DBA_Data", strSecNull
          CreateThisFolder strDrive & strDirSQL & "\" & strInstNode & ".Data\model", strSecNull
          CreateThisFolder strDrive & strDirSQL & "\" & strInstNode & ".Data\msdb", strSecNull
          CreateThisFolder strDrive & strDirSQL & "\" & strInstNode & ".Data\ReportServer", strSecNull

        End If
      Case Else
        Call FBLog("Setup " & strDrive & ": drive bypassed")
    End Select
  Next

End Sub


Sub CreateThisFolder(strFolderName)
  Call DebugLog("CreateThisFolder: " & strFolderName)

  If Not objFSO.FolderExists(strFolderName) Then
    objFSO.CreateFolder(strFolderName)
    strCmd    = strProgCacls & " " & strFolderName & " /T /C /G " & strFilePerm
    Call Util_RunExec(strCmd, "", strResponseYes, -1)
  End If

End Sub


Sub SetupThisDrive(strDrive, strDriveLabel)
  Call DebugLog("SetupThisDrive: " & strDrive)

  If Instr(strDrvUsed, strDrive) = 0 Then       
    Call FBLog("Setup " & strDrive & ": drive for " & strDriveLabel)
    strDrvUsed      = strDrvUsed & strDrive
    LabelThisDrive  strDrive & ":", strDriveLabel
    CreateThisShare strDrive & ":\", "(" & strDrive & ") " & strDriveLabel
  End If

End Sub


Sub LabelThisDrive(strDrive, strDriveLabel)
  Call DebugLog("LabelThisDrive: " & strDrive)
' Code to clear IndexingEnabled flag adapted from "Windows Server Cookbook" by Robbie Allen, ISBN 0-596-00633-0

  Call DebugLog("Labelling drive " & strDrive)
  If strOSVersion > "5.1" Then
    strCmd          = "SELECT * FROM Win32_Volume WHERE Name='" & strDrive & "\\'"
  Else
    strCmd          = "SELECT * FROM Win32_LogicalDisk WHERE DeviceID='" & strDrive & "'"
  End If
  Set colVol        = objWMI.ExecQuery(strCmd)
  If colVol.Count <> 1 Then
    err.Raise 8, "", "(" & strProcessIdLabel & ") " & "Volume not found:" & strDrive
  End If

  For Each objVol In colVol
    Select Case True
      Case strOSVersion > "5.1"
        objVol.IndexingEnabled = 0
        Select Case True
          Case objVol.Label      = strDriveLabel
            ' Nothing
          Case strInstance = "MSSQLSERVER"
            objVol.Label         = Left(strDriveLabel, 32)
          Case strInstance = "SQLEXPRESS"
            objVol.Label         = Left(strDriveLabel, 32)
          Case Else
            objVol.Label         = Left(strDriveLabel & "\" & strInstance, 32)
        End Select
      Case Else
        Select Case True
          Case objVol.VolumeName = strDriveLabel
            ' Nothing
          Case strInstance = "MSSQLSERVER"
            objVol.VolumeName    = Left(strDriveLabel, 32)
          Case strInstance = "SQLEXPRESS"
            objVol.VolumeName    = Left(strDriveLabel, 32)
          Case Else
            objVol.VolumeName    = Left(strDriveLabel & "\" & strInstance, 32)
        End Select
    End Select
    objVol.Put_
  Next

  strCmd            = "ATTRIB +I " & strDrive & "\*.* /D /S"
  Call Util_RunCmdAsync(strCmd, 0)

End Sub


Sub CreateThisShare(strDrive, strShareName)
  Call DebugLog("CreateThisShare: " & strShareName)
  Dim objACEAdmin, objACEUser, objSecDesc, objShare, objShareParm

  If strSetupShares = "Yes" Then
    Set objSecDesc           = objWMI.Get("Win32_SecurityDescriptor").SpawnInstance_
    Set objShare             = objWMI.Get("win32_Share")

    Set objACEAdmin          = objWMI.Get("Win32_Ace").SpawnInstance_
    objACEAdmin.AccessMask   = 2032127 ' = "Full"
    objACEAdmin.AceFlags     = 3
    objACEAdmin.AceType      = 0
    objACEAdmin.Trustee      = SetGroupTrustee("Administrators")
    Set objACEUser           = objWMI.Get("Win32_Ace").SpawnInstance_
    objACEUser.AccessMask    = 1245631 ' = "Change"
    objACEUser.AceFlags      = 3
    objACEUser.AceType       = 0
    objACEUser.Trustee       = SetGroupTrustee("Users")

    objSecDesc.DACL          = Array(objACEAdmin, objACEUser)

    Set objShareParm         = objShare.Methods_("Create").InParameters.SpawnInstance_ 
    objShareParm.Access      = objSecDesc
    objShareParm.Description = strShareName & " Share"
    objShareParm.Name        = strShareName
    objShareParm.Path        = strDrive
    objShareParm.Type        = 0
    objShare.ExecMethod_ "Create",  objShareParm
  End If

End Sub 


Function SetGroupTrustee(strGroup) 
  Call DebugLog("SetGroupTrustee: " & strGroup)

  Dim objTrustee
  Dim colAccount
  Dim objAccount
  Dim objAccountSID
  Dim strSID

  Set objTrustee    = objWMI.Get("Win32_Trustee").Spawninstance_ 
  Set colAccount    = objWMI.ExecQuery("Select * from Win32_Group Where Name='" & strGroup & "' AND LocalAccount=True")
  For Each objAccount in colAccount
    Set objAccountSID = objWMI.Get("Win32_SID.SID='" & objAccount.SID &"'") 
    Exit For
    Next
  objTrustee.Domain = strServer 
  objTrustee.Name   = strGroup 
  objTrustee.SID    = objaccountSID.BinaryRepresentation 

  Set colAccount    = nothing 
  Set objAccount    = nothing 
  Set objAccountSID = nothing 
  Set SetGroupTrustee = objTrustee 

End Function 


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
