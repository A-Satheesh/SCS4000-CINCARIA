VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmAlarm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AlarmForm(Alarm.frm)"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1920
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "frmAlarm.frx":0000
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   720
      OleObjectBlob   =   "frmAlarm.frx":0234
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
End
Attribute VB_Name = "frmAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim A As Long
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, A))
    If A >= &H800 Then
        A = A And &HF7FF
        'Close the alarm
        checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, A))
    End If
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim A As Long
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, A))
    If A >= &H800 Then
        A = A And &HF7FF
        'Close the alarm
        checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, A))
    End If
End Sub


