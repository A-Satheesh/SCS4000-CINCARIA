VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmEmergencyStopForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frmEmergency Stop!"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmEmergencyStop.frx":0000
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1095
      Left            =   600
      OleObjectBlob   =   "frmEmergencyStop.frx":0234
      TabIndex        =   0
      Top             =   600
      Width           =   3735
   End
End
Attribute VB_Name = "frmEmergencyStopForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim Value1, Value2, Value3, DirReset As Long
    Dim message As String
    Dim DriverXYZ As Long
    
    Emergency_Stop = True
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, Value1, Value2, Value3, 0))
    Do While ((Value1 <> 0) Or (Value2 <> 0) Or (Value3 <> 0))
        message = "Please release the E-Stop button!"
        Call MsgBox(message, vbOKOnly, "Techno Digm")
        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, Value1, Value2, Value3, 0))
        Value1 = (Value1 And &H20)
        Value2 = (Value2 And &H20)
        Value3 = (Value3 And &H20)
    Loop
    
    Unload Me
    
    Dim A As Long
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, A))
    If A >= &H800 Then
        A = A And &HF7FF
        'Close the alarm
        checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, A))
    End If
    
    If (Red_Light = True) Then
        Red_Light = False
        Green_Light = True
    End If
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, DirReset))
    DirReset = DirReset Or &H200
    'Do reset
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, DirReset))
    
    Call Sleep(0.3)
    
    DirReset = DirReset And &HFDFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, DirReset))
    
    Call Sleep(1)
    
    Servo_On
End Sub

Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub

