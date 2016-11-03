VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmforgetpassword 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forget Password"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton cmdreset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtsecuritykey 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtuserlevel 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "frmforgetpassword.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel label1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmforgetpassword.frx":0234
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmforgetpassword.frx":02A6
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "frmforgetpassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
Unload Me
End Sub

Private Sub cmdreset_Click()
'Dim Password As String
Dim MD5 As New clsMD5
Dim securitykey As String
securitykey = "F2F9F5DBCE18280CB038072DA8CD1777"

'frmlogin.Open_File
If LCase(Me.txtuserlevel.Text) <> "engineer" Then
    MsgBox "Please enter correct user level!"
Else
    If Len(txtsecuritykey.Text) = 0 Then
        MsgBox "Please enter your security key.", vbExclamation
        Exit Sub
    End If
        If Len(txtsecuritykey.Text) > 0 Then
            If UCase(MD5.DigestStrToHexStr(txtsecuritykey.Text)) <> securitykey Then
            'If txtsecuritykey.Text <> securitykey Then
                MsgBox "Invalid securitykey." & vbNewLine & "You must enter the valid security key.", vbInformation
                Exit Sub
            Else
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set A = fs.createtextfile("C:\Desktop Access\LogIn_Maintenance", True)
                A.writeline ("")
                A.Close
                confirmreset = True
                Unload Me
            End If
        End If
    End If

End Sub

Private Sub Form_Activate()
    Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn")
    Skin1.ApplySkin Me.hWnd
    Me.txtuserlevel.SetFocus
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub
