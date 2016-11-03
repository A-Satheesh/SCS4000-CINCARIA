VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmlogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login to Maintenance "
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   2760
      Picture         =   "frmlogin.frx":0000
      ScaleHeight     =   300
      ScaleWidth      =   1335
      TabIndex        =   7
      Top             =   1080
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1160
      Picture         =   "frmlogin.frx":0544
      ScaleHeight     =   300
      ScaleWidth      =   1455
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      Default         =   -1  'True
      Height          =   390
      Left            =   3000
      TabIndex        =   2
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   1200
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   720
      Width           =   2895
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   -120
      OleObjectBlob   =   "frmlogin.frx":0B4A
      Top             =   1560
   End
   Begin ACTIVESKINLibCtl.SkinLabel label1 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmlogin.frx":0D7E
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "frmlogin.frx":0DEC
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "frmlogin.frx":0E60
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmlogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Password As String

Private Sub cmdcancel_Click()
    mCancel = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mCancel = False
    Password = Me.txtPassword.Text
    Call DoLogin
End Sub

Private Sub Form_Activate()
    Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn")
    Skin1.ApplySkin Me.hWnd
    Me.txtPassword.SetFocus
    
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Picture1_Click()
    Me.Hide
    SetWindowOnTop Me, False    '@$K
    frmforgetpassword.Show vbModal
    If confirmreset = True Then
        Unload Me
    Else
        Me.Show vbModal
    End If
End Sub

Private Sub Picture2_Click()
    Me.Hide
    frmchangepassword.Show vbModal
    Me.Show vbModal
End Sub

Public Function Open_File()
    Dim words() As String
    Dim line As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.OpenTextFile("C:\Desktop Access\LogIn_Maintenance", 1, False)

    Do While A.AtEndOfStream <> True
        line = A.ReadLine
        
        If (line <> "") Then
            words() = Split(line, "=")
            
                If StrComp(words(0), "Username", vbTextCompare) = 1 Then
                    'fUsername = words()(1)
                Else
                    fPassword = words()(1)
                End If
        End If
    Loop
    A.Close
End Function

Private Function DoLogin()
    Dim MD5 As New clsMD5
    Dim Respond As Integer
    Randomize
    Call Open_File 'compare with password from file
    
    'check if the password is correct
    If UCase(MD5.DigestStrToHexStr(Password)) = UCase(fPassword) Then
        loginsuccessful = True
        'Me.Hide
        Unload Me
    Else
        Picture1.Visible = True
        Respond = MsgBox("Invalid login, do you want to try again ?", vbQuestion + vbYesNo, "Invalid Login")
        If (Respond = vbNo) Then
            loginsuccessful = False
            Unload Me
        End If
    End If
End Function

Private Sub Form_Unload(cancel As Integer)
    SetWindowOnTop Me, False    '@$K
    Unload Me
End Sub
