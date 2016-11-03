VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmregisteration 
   Caption         =   "Registeration Form"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword2 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1695
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtPassword1 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1695
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   360
      Width           =   2895
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdregister 
      Caption         =   "&Register"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1560
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "frmregisteration.frx":0000
      Top             =   1440
   End
   Begin ACTIVESKINLibCtl.SkinLabel label1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmregisteration.frx":0234
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel label2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmregisteration.frx":02AE
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
End
Attribute VB_Name = "frmregisteration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdregister_Click()
    Dim MD5 As New clsMD5, NewPassword As String
    
    If Len(txtPassword1.Text) > 0 Or Len(txtPassword2.Text) > 0 Then
        If txtPassword1.Text <> txtPassword2.Text Then
            MsgBox "Confirm password must the same as Password field", vbExclamation
            Exit Sub
        End If
    End If
    
    ' get the hash of the passwords
    NewPassword = UCase(MD5.DigestStrToHexStr(Me.txtPassword1.Text))
    'OldPassword = UCase(MD5.DigestStrToHexStr(Me.txtOldPassword.Text))
    
    Dim fs As Object, A As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If frmmainwindow.Combo1.Text = "Maintenance" Then
        Set A = fs.createtextfile("C:\Desktop Access\LogIn_Maintenance", True)
        A.writeline ("Password =" & NewPassword)
        A.Close
    Else
       Set A = fs.createtextfile("C:\Desktop Access\LogIn_Profile", True)
        A.writeline ("Password =" & NewPassword)
        A.Close
    End If
    
    txtPassword1.Text = ""
    txtPassword2.Text = ""
    exitregister = False
    Unload Me
    
End Sub

Private Sub cmdExit_Click()
    exitregister = True
    Unload Me
End Sub

Private Sub Form_activate()
Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn")
Skin1.ApplySkin Me.hWnd
If Len(Me.txtPassword1.Text) = 0 Then
    Me.txtPassword1.SetFocus
End If
End Sub


