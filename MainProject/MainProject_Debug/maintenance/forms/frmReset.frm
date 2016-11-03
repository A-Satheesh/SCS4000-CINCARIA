VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmReset 
   Caption         =   "ResetForm(Reset.frm)"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "frmReset(Form1.frm).frx":0000
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   840
      OleObjectBlob   =   "frmReset(Form1.frm).frx":0234
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    Call P1240MotWrReg(boardNum, Y_axis, WR3, &H200)
    Call moveToHome
    
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")

    Skin1.ApplySkin Me.hWnd
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub


