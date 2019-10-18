VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "密码设置"
   ClientHeight    =   2265
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4785
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skn 
      Left            =   3480
      OleObjectBlob   =   "frmSetting.frx":08CA
      Top             =   1560
   End
   Begin VB.TextBox txtNewPwd2 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1680
      Width           =   2295
   End
   Begin VB.TextBox txtNewPwd1 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1160
      Width           =   2295
   End
   Begin VB.TextBox txtUserID 
      BackColor       =   &H8000000F&
      Height          =   375
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "www.mama520.cn"
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtPwd 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   960
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "888888"
      Top             =   640
      Width           =   2295
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确认"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmSetting.frx":0AFE
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmSetting.frx":0B62
      TabIndex        =   5
      Top             =   640
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmSetting.frx":0BC6
      TabIndex        =   7
      Top             =   1160
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "frmSetting.frx":0C2A
      TabIndex        =   9
      Top             =   1680
      Width           =   735
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdCancel_Click()

    Unload Me
    If frmMain.Visible = False Then
        frmMain.Show
    End If
    frmMain.SetFocus
    
    
End Sub

Private Sub Form_Load()
    If Dir(AppPath() & Const_SkinFile) = vbNullString Then
        MsgBox "数据文件已经被破坏，无法加载数据！", vbCritical + vbOKOnly
        End
    End If

    Skn.LoadSkin AppPath() + Const_SkinFile
    Skn.ApplySkin Me.hWnd
    
    Me.txtUserID.Text = frmMain.txtUserID.Text
    Me.txtPwd.Text = frmMain.txtPwd.Text
    If IsFirstRun Then
        Me.Caption = "密码初始化设置"
    Else
        Me.Caption = "密码设置"
    End If
    

End Sub

Private Sub txtPwd_GotFocus()

txtPwd.SelStart = 0
txtPwd.SelLength = Len(txtPwd.Text)

End Sub

Private Sub txtnewpwd1_GotFocus()

txtNewPwd1.SelStart = 0
txtNewPwd1.SelLength = Len(txtNewPwd1.Text)

End Sub

Private Sub txtNewPwd2_GotFocus()

txtNewPwd2.SelStart = 0
txtNewPwd2.SelLength = Len(txtNewPwd2.Text)

End Sub

Private Function CheckSave() As Boolean

CheckSave = False


If Not CheckPassword(txtPwd.Text) Then
    MsgBox "密码输入错误！", vbInformation + vbOKOnly
    CheckSave = False
    txtPwd.SetFocus
    Exit Function
End If

If Trim(txtNewPwd1.Text) = "" Then
    MsgBox "密码不能够为空！", vbInformation + vbOKOnly
    CheckSave = False
    txtNewPwd1.SetFocus
    Exit Function
End If

If Trim(txtNewPwd2.Text) = "" Then
    MsgBox "密码不能够为空！", vbInformation + vbOKOnly
    CheckSave = False
    txtNewPwd2.SetFocus
    Exit Function
End If


If txtNewPwd1.Text <> txtNewPwd2.Text Then
    MsgBox "密码输入不一致！", vbInformation + vbOKOnly
    CheckSave = False
    txtNewPwd1.SetFocus
    Exit Function
End If

CheckSave = True

End Function

Private Sub cmdOK_Click()

If Not CheckSave() Then
   Exit Sub
End If
SaveUserDB
MsgBox "密码修改成功！请牢记密码：" & txtNewPwd1.Text, vbInformation + vbOKOnly
Unload Me
frmMain.SetFocus




End Sub
