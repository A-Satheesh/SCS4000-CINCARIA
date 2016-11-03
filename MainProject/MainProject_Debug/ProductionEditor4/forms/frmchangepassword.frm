VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmchangepassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password to Profile Editor"
   ClientHeight    =   2610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2610
   ScaleWidth      =   5265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtOldPassword 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   360
      Width           =   2895
   End
   Begin VB.TextBox txtPassword2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtPassword1 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   840
      Width           =   2895
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "Update"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmchangepassword.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel label1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmchangepassword.frx":0234
      TabIndex        =   5
      Top             =   360
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmchangepassword.frx":02AA
      TabIndex        =   6
      Top             =   840
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmchangepassword.frx":0320
      TabIndex        =   7
      Top             =   1320
      Width           =   1695
   End
End
Attribute VB_Name = "frmchangepassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
'Dim Username As String
Dim Password As String
Dim loginsuccessful As Boolean
Dim MD5 As New clsMD5

frmlogin.Open_File
If Len(txtPassword1.Text) = 0 Or Len(txtPassword2.Text) = 0 Or Len(txtOldPassword.Text) = 0 Then
MsgBox "Please enter your password.", vbExclamation
Exit Sub
End If
    If Len(txtPassword1.Text) > 0 Or Len(txtPassword2.Text) > 0 Then
        If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Confirm password must be same as your new password.", vbExclamation
            Exit Sub
        Else
            If UCase(MD5.DigestStrToHexStr(txtOldPassword.Text)) <> fPassword Then
                MsgBox "Invalid old password." & vbNewLine & "You must enter the valid password.", vbInformation
                Exit Sub
            Else
                NewPassword = UCase(MD5.DigestStrToHexStr(Me.txtPassword1.Text))
                    'Dim fs As Object, A As Object
                    Set fs = CreateObject("Scripting.FileSystemObject")
                    Set A = fs.createtextfile("C:\Desktop Access\LogIn_Profile", True)
                    'A.writeline ("Username =" & Me.txtUserName.Text)
                    A.writeline ("Password =" & NewPassword)
                    A.Close
                    txtPassword1.Text = ""
                    txtPassword2.Text = ""
                    txtOldPassword.Text = ""
                    MsgBox "Your password has been changed successfully!"
                    Unload Me
            End If
        End If
    End If


End Sub

Private Sub Form_activate()
    'Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    Me.txtOldPassword.SetFocus
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub
