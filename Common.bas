Attribute VB_Name = "Common"
Option Explicit

Public Const Const_Website = "Our website is ( www.mama520.cn )....\"
Public Const Const_UserDBFile = "PrivateUser.DB"
Public Const Const_CustomerFile = "Private.ini"
Public Const Const_SkinFile = "Private.dat"
Public Const Const_DefaultURL = "http://www.mama520.cn/Software_AD/PrivateMaMa/index.htm"
Public IsFirstRun As Boolean
Public g_UserPwd As String
Public g_URL As String


Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

'for top most window
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

  
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Function SetTopMostWindow(ByVal thwnd As Long, ByVal b As Boolean) As Boolean
If b Then
    If SetWindowPos(thwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE) <> 0 Then SetTopMostWindow = True Else SetTopMostWindow = False
Else
    If SetWindowPos(thwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE) <> 0 Then SetTopMostWindow = True Else SetTopMostWindow = False
End If

End Function

Public Function AppPath() As String

Dim sPath As String

sPath = App.Path

If Right(App.Path, 1) <> "\" Then
    sPath = sPath + "\"
End If

AppPath = LCase(sPath)

End Function
Public Function GetDriveTypeOfString(ByVal sDrvName As String) As String

                  Select Case GetDriveType(sDrvName)
                  Case 0
                          GetDriveTypeOfString = "不明"
                  Case 2
                          GetDriveTypeOfString = "移动盘" '包括软盘与U盘
                  Case 3
                          GetDriveTypeOfString = "硬盘"
                  Case 4
                          GetDriveTypeOfString = "网络盘"
                  Case 5
                          GetDriveTypeOfString = "光驱"
                  Case 6
                          GetDriveTypeOfString = "RamDisk"
                  Case Else
                          GetDriveTypeOfString = "不明"
                  End Select



End Function
Public Function GetSystemDrive() As String

GetSystemDrive = Environ("SystemDrive") + "\"

End Function

Public Function AppPrePath() As String
'当使用FilePacker压缩时，解压后，读取上级目录的功能。
'此处得到应用程序上级目录，以方便读取个性化配置文档
'Const_CustomerFile 常量所对应的文件。

Dim sPath As String

    Dim FileName As String
    Dim FilePrePath As String
    Dim sTmp As String
    Dim iPos As Integer
    
    sPath = App.Path

If Right(App.Path, 1) <> "\" Then
    
    sTmp = StrReverse(sPath)
    iPos = InStr(sTmp, "\")
    FilePrePath = Mid(sTmp, iPos)
    FilePrePath = StrReverse(FilePrePath)
    sPath = FilePrePath
    
End If

AppPrePath = LCase(sPath)

End Function

Public Function HasPassword() As Boolean
On Error GoTo err
'MkDir GetAppLocalDisk & Const_Website
'RmDir GetAppLocalDisk & Const_Website
HasPassword = False
Exit Function
err:
HasPassword = True

End Function

Public Sub InIt()

    On Error Resume Next
    Dim FileNum As Integer
    Dim sBatFile As String
    
   Kill AppPath() & Const_UserDBFile
    
    sBatFile = GetSystemDrive() & "HappyQQ__5_5abc5_Pwd520.bat"
    g_URL = Const_DefaultURL
       
    
    FileNum = FreeFile
    
    
    Open sBatFile For Output As #FileNum
    

        Print #FileNum, "Copy/y """ & GetAppLocalDisk & Const_Website & Const_UserDBFile; """ """ & AppPath() + Const_UserDBFile; """"
        Print #FileNum, "del %0"
                
    Close #FileNum
    
    Shell sBatFile, vbHide
'    Dim bHasPwd As Boolean
'
'  '  bHasPwd = HasPassword()
'
'
'    If bHasPwd Then
'
''        While Dir(AppPath() + Const_UserDBFile) = ""
''        Wend
'        Sleep (2000)
'
'    End If
    
      Sleep (2000)

    
    
    Dim sAppTitle, sAppBoldTitle, sAppMessage, sPrivateURL, sRegisterCode As String
    Dim sUserName, sPwd As String
    
    frmMain.txtMachineCode.Text = GetDriveSerialNumber()
    
    
    sRegisterCode = GetIniParam(AppPath() & Const_UserDBFile, "Config", "RegisterCode")
    
    frmMain.txtRegisterCode.Text = IIf(sRegisterCode = "", frmMain.txtRegisterCode.Text, sRegisterCode)
    
    '这一部分保存在User.DB里面
        
    sUserName = GetIniParam(AppPath() & Const_UserDBFile, "Config", "UserID")
    sPwd = GetIniParam(AppPath() & Const_UserDBFile, "Config", "UserPwd")
    
    
    If sRegisterCode = CalculateMD5("www.mama520.cn" & frmMain.txtMachineCode.Text & "HappyQQ520") Then
    
        sAppTitle = GetIniParam(AppPath() & Const_CustomerFile, "Config", "AppTitle")
        sAppBoldTitle = GetIniParam(AppPath() & Const_CustomerFile, "Config", "AppBoldTitle")
        sAppMessage = GetIniParam(AppPath() & Const_CustomerFile, "Config", "AppMessage")
        sPrivateURL = GetIniParam(AppPath() & Const_CustomerFile, "Config", "PrivateURL")
        
        g_URL = IIf(sPrivateURL = "", Const_DefaultURL, sPrivateURL)
        
        
        
        

    End If
    
    
    
    
    frmMain.Caption = IIf(sAppTitle = "", frmMain.Caption, sAppTitle)
    frmMain.lblBoldTitle = IIf(sAppBoldTitle = "", frmMain.lblBoldTitle.Caption, Left(sAppBoldTitle, 9))
    frmMain.lblMessage = IIf(sAppMessage = "", frmMain.lblMessage.Caption, sAppMessage)
    
    

    
    
    
    frmMain.txtUserID.Text = IIf(sUserName = "", frmMain.txtUserID.Text, sUserName)
    
'    frmMain.txtPwd.Text = IIf(sPwd = "", frmMain.txtPwd.Text, sPwd)


    If Dir(AppPath() + Const_UserDBFile) = "" Then
        IsFirstRun = True
        frmMain.txtPwd.Text = "888888"
        g_UserPwd = CalculateMD5(frmMain.txtPwd.Text)
        MkDir GetAppLocalDisk & Const_Website
       ' frmMain.Hide
       ' frmSetting.Show
       ' frmSetting.SetFocus
    Else
        frmMain.txtPwd.Text = ""
        g_UserPwd = sPwd
        IsFirstRun = False
    End If
    
    
   
    'MkDir AppPath() & Const_Website

    
    
   ' MsgBox """" & AppPath() & Const_Website & Const_UserDBFile & """"
    
   ' FileCopy """" & AppPath() & Const_Website & Const_UserDBFile & """", AppPath() + Const_UserDBFile
    
      
    

End Sub

Public Sub SaveUserDB()

    On Error Resume Next
    Dim FileNum As Integer
    Dim sBatFile As String
    
    WriteWinIniParam AppPath() & Const_UserDBFile, "Config", "UserID", frmSetting.txtUserID.Text
    WriteWinIniParam AppPath() & Const_UserDBFile, "Config", "UserPwd", CalculateMD5(frmSetting.txtNewPwd1.Text)
    
    WriteWinIniParam AppPath() & Const_UserDBFile, "Config", "RegisterCode", frmMain.txtRegisterCode.Text
    
    
    
    
    
    
    
    
    
    g_UserPwd = CalculateMD5(frmSetting.txtNewPwd1.Text)
    
    
    sBatFile = GetSystemDrive() & "HappyQQ__Save_5abcSave_Pwd520.bat"
    
    
    FileNum = FreeFile
    
    
    Open sBatFile For Output As #FileNum
    

        Print #FileNum, "Copy/y """ & AppPath() & Const_UserDBFile & """ """ & GetAppLocalDisk & Const_Website & """"
        Print #FileNum, "del """ & AppPath() & Const_UserDBFile & """"
        Print #FileNum, "del %0"
                
    Close #FileNum
    
    Shell sBatFile, vbHide
    frmMain.txtPwd.Text = frmSetting.txtNewPwd1.Text


End Sub

Public Function GetAppLocalDisk() As String
    
    GetAppLocalDisk = Left(App.Path, 3)
    
    
    
End Function


Public Function CheckPassword(ByVal sPwd As String) As Boolean

CheckPassword = (CalculateMD5(sPwd) = g_UserPwd)


End Function
Public Function CheckProtectFile(ByVal sFilePath As String) As Boolean

Dim sProtectFilePath(4) As String
Dim i As Integer

sProtectFilePath(0) = GetSystemDrive() & "RECYCLER"
sProtectFilePath(1) = GetSystemDrive() & "WINDOWS"
sProtectFilePath(2) = GetSystemDrive() & "Program Files"
sProtectFilePath(3) = GetSystemDrive() & "Documents and Settings"

For i = 0 To 3
    If InStr(1, sFilePath, sProtectFilePath(i)) <> 0 Then
        Exit For
    End If
Next i

If i > 3 Then

    CheckProtectFile = False
Else

    CheckProtectFile = True

End If






End Function

Public Sub EncryptFilePath(ByVal sFilePath As String)

    On Error Resume Next
    Dim FileNum As Integer
    Dim sBatFile As String
    
    sBatFile = GetSystemDrive() & "HappyQQ__0_0abc0_1314520.bat"
    
    
    FileNum = FreeFile
    
    
    Open sBatFile For Output As #FileNum
    
        Print #FileNum, "md """ & sFilePath; "..\"""
        Print #FileNum, "md """ & sFilePath; "...\"""
        Print #FileNum, "move """ & sFilePath; """ """ & sFilePath; "...\"""
        
        Print #FileNum, "rd """ & sFilePath; "...\"""
        Print #FileNum, "del %0"
                
    Close #FileNum
    
    Shell sBatFile, vbHide



End Sub

Public Sub DecryptFilePath(ByVal sFilePath As String)

    On Error Resume Next
    Dim FileNum As Integer
    Dim FileName As String
    Dim FilePrePath As String
    Dim sTmp As String
    Dim iPos As Integer
    Dim sBatFile As String
    
    sBatFile = GetSystemDrive() & "HappyQQ__1_1cab1_1314520.bat"
    
    sTmp = StrReverse(sFilePath)
    iPos = InStr(sTmp, "\")
    
    
    FileName = Left(sTmp, iPos - 1)
    FileName = StrReverse(FileName)
    
    FilePrePath = Mid(sTmp, iPos)
    FilePrePath = StrReverse(FilePrePath)
    
    
    FileNum = FreeFile
    
    
    Open sBatFile For Output As #FileNum
    
        Print #FileNum, "md """ & sFilePath; "...\"""
        Print #FileNum, "move """ & sFilePath; "...\\" & FileName & """ """ & FilePrePath & """"
        Print #FileNum, "rd """ & sFilePath; "...\"""
        Print #FileNum, "rd """ & sFilePath; "..\"""
        Print #FileNum, "del %0"
                
    Close #FileNum
    
    Shell sBatFile, vbHide



End Sub
