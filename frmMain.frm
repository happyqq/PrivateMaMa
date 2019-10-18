VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "隐私文件夹——妈妈社区出品 ( http://www.mama520.cn )"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox PicAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   825
      ScaleWidth      =   2085
      TabIndex        =   10
      ToolTipText     =   "双击此处可以看到更多隐私噢！"
      Top             =   0
      Width           =   2115
   End
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   480
      OleObjectBlob   =   "frmMain.frx":39A2
      Top             =   3960
   End
   Begin VB.Frame frameMenu 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   840
      Width           =   6615
      Begin VB.TextBox txtRegisterCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   2400
         Width           =   4575
      End
      Begin VB.TextBox txtMachineCode 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1920
         Width           =   4575
      End
      Begin VB.CommandButton cmdSetting 
         Caption         =   "密码设置(&S)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   12
         Top             =   3000
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "frmMain.frx":3BD6
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtPwd 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   5
         Text            =   "888888"
         Top             =   1440
         Width           =   4575
      End
      Begin VB.TextBox txtUserID 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   4
         Text            =   "www.mama520.cn"
         Top             =   960
         Width           =   4575
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&Q)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   3
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdDecrypt 
         Caption         =   "解密(&D)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdEncrypt 
         Caption         =   "加密(&E)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         TabIndex        =   1
         Top             =   3000
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "frmMain.frx":3C3A
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "frmMain.frx":3CA4
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel lblMessage 
         Height          =   495
         Left            =   1440
         OleObjectBlob   =   "frmMain.frx":3D0C
         TabIndex        =   9
         Top             =   480
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "frmMain.frx":3DE8
         TabIndex        =   13
         Top             =   1920
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   600
         OleObjectBlob   =   "frmMain.frx":3E4C
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblBoldTitle 
      Height          =   615
      Left            =   2160
      OleObjectBlob   =   "frmMain.frx":3EB0
      TabIndex        =   11
      ToolTipText     =   "妈妈社区出品 (详情访问 http://www.mama520.cn )"
      Top             =   240
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Form_Load()

On Error Resume Next

InIt
    

'   SetTopMostWindow Me.hwnd, True
   
End Sub
Private Sub PicAbout_DblClick()
    frmMiniWeb.Show
End Sub

Private Sub txtPwd_GotFocus()

txtPwd.SelStart = 0
txtPwd.SelLength = Len(txtPwd.Text)

End Sub

Private Sub Do_Private(ByVal bEncrypt As Boolean)
    On Error Resume Next

    Dim sFilePath As String
    
    If Not CheckPassword(txtPwd.Text) Then

      MsgBox "密码错误，不能够进行" & IIf(bEncrypt, "加密", "解密") & "操作！", vbInformation + vbOKOnly
      Exit Sub
    
    End If

    sFilePath = BrowseForFolder(Me.hWnd, "请选择要" & IIf(bEncrypt, "加密", "解密") & "的文件夹")
    
    If sFilePath = "" Or Right(sFilePath, 1) = "\" Then
    
          MsgBox "请选择对应的文件夹！", vbInformation + vbOKOnly
          Exit Sub
    
    
    End If
    
    If Left(sFilePath, 3) <> GetAppLocalDisk Then
    
          MsgBox "^_^ 我在哪个盘，你就针对这个盘进行加解密操作吧！只允许对" + GetAppLocalDisk + "盘进行操作！", vbInformation + vbOKOnly
          Exit Sub
    
    End If
    

    If CheckProtectFile(sFilePath) Then
    
          MsgBox "文件夹为系统文件夹，不允许对其进行" & IIf(bEncrypt, "加密", "解密") & "操作！", vbInformation + vbOKOnly
          Exit Sub
    End If
    
    If Not bEncrypt And Left(StrReverse(sFilePath), 1) <> "." Then
    
        MsgBox "文件夹没有被加密，请选择一个被加密的文件夹进行解密操作！", vbInformation + vbOKOnly
        Exit Sub
    
    End If
    
    
    If Not bEncrypt Then
        sFilePath = Left(sFilePath, Len(sFilePath) - 1)
    End If


    If MsgBox("你确定对“" & sFilePath & "”进行" & IIf(bEncrypt, "加密", "解密") & "？", vbYesNo + vbQuestion) = vbNo Then Exit Sub

    If bEncrypt Then
    
        EncryptFilePath (sFilePath)
        
    Else
    
        DecryptFilePath (sFilePath)
        
    End If
    
    MsgBox IIf(bEncrypt, "加密", "解密") & "操作已完成！", vbInformation + vbOKOnly
    


End Sub



Private Sub cmdDecrypt_Click()

Do_Private (False)



'    On Error Resume Next
'
'    Dim sFilePath As String
'sFilePath = BrowseForFolder(Me.hWnd, "请选择要解密的文件夹")
'
'
'
'If sFilePath = "" Or Right(sFilePath, 1) = "\" Then
'
'      MsgBox "请选择对应的文件夹！", vbInformation + vbOKOnly
'      Exit Sub
'
'
'End If
'
'If CheckProtectFile(sFilePath) Then
'
'      MsgBox "文件夹为系统文件夹，不允许对其进行解密操作！", vbInformation + vbOKOnly
'      Exit Sub
'End If
'
'sFilePath = Left(sFilePath, Len(sFilePath) - 1)
'
'
'If MsgBox("你确定对“" & sFilePath & "”进行解密？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
'
'
'
'    DecryptFilePath (sFilePath)
'
'    MsgBox "解密操作已完成！", vbInformation + vbOKOnlyaaa
    


End Sub




Private Sub cmdEncrypt_Click()

Do_Private (True)

'    On Error Resume Next
'
'Dim sFilePath As String
'
'If Not CheckPassword(txtPwd.Text) Then
'
'      MsgBox "密码错误，不能够进行相关的操作！", vbInformation + vbOKOnly
'      Exit Sub
'
'End If
'
'
'sFilePath = BrowseForFolder(Me.hWnd, "请选择要加密的文件夹")
'
'
'
'If sFilePath = "" Or Right(sFilePath, 1) = "\" Then
'
'      MsgBox "请选择对应的文件夹！", vbInformation + vbOKOnly
'      Exit Sub
'
'End If
'
'If CheckProtectFile(sFilePath) Then
'
'      MsgBox "文件夹为系统文件夹，不允许对其进行加密操作！", vbInformation + vbOKOnly
'      Exit Sub
'End If
'
'
'
'If MsgBox("你确定对“" & sFilePath & "”进行加密？", vbYesNo + vbQuestion) = vbNo Then Exit Sub
'
'
'EncryptFilePath (sFilePath)
'
'    MsgBox "加密操作已完成！", vbInformation + vbOKOnly
'



End Sub

Private Sub cmdSetting_Click()

    frmSetting.Show

End Sub



Private Sub Form_DblClick()

MsgBox "软件名称：隐私文件夹  1.0" & Chr(13) & "官方网站：http://www.mama520.cn  ［妈妈社区］" & Chr(13) & "开发作者：黄启清" & Chr(13) & "联系方式：HuangQiQing@gmail.com" & Chr(13) & "      QQ号：18367144 ［添加好友时，请注明您所使用的软件名称］", vbInformation + vbOKOnly

End Sub

Private Sub Form_Initialize()

    If Dir(AppPath() & Const_SkinFile) = vbNullString Then
        
        MsgBox "数据文件已经被破坏，无法加载数据！", vbCritical + vbOKOnly
        End
    
    End If

    Skn.LoadSkin AppPath() + Const_SkinFile
    Skn.ApplySkin Me.hWnd
    


End Sub


Private Sub cmdExit_Click()

    End

End Sub



Private Sub txtRegisterCode_Change()

If Trim(txtRegisterCode.Text = "") Then Exit Sub

If CalculateMD5("www.mama520.cn" & txtMachineCode.Text & "HappyQQ520") = Trim(txtRegisterCode.Text) Then
    txtUserID.Locked = False
    txtUserID.BackColor = &H80000005
Else
    txtUserID.Locked = True
    txtUserID.BackColor = &H8000000F
End If




End Sub
