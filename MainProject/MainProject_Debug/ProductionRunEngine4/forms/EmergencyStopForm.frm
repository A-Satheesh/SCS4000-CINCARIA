VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form EmergencyStopForm 
   Caption         =   "Emergency Stop!"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "EmergencyStopForm.frx":0000
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1455
      Left            =   600
      OleObjectBlob   =   "EmergencyStopForm.frx":0234
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
End
Attribute VB_Name = "EmergencyStopForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim resetValue, valueX, valueY, valueZ As Long
    Dim message As String
    'Check whether the e-stop release or not
    'if not, it will show a message to user.
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, valueX, valueY, valueZ, 0))
    valueX = (valueX And &H20)
    valueY = (valueY And &H20)
    valueZ = (valueZ And &H20)
    Do While ((valueX <> 0) Or (valueY <> 0) Or (valueZ <> 0))
        message = "Please release the E-Stop button!"
        Call MsgBox(message, vbOKOnly, "Techno Digm")
        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, valueX, valueY, valueZ, 0))
        valueX = (valueX And &H20)
        valueY = (valueY And &H20)
        valueZ = (valueZ And &H20)
    Loop
    
    Unload Me
    'Stop alarm
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, resetValue))
    resetValue = resetValue And &HF7FE
    checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, resetValue))
    
    If (errorStatus = True) Then
        Red_Light = False
    End If
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, resetValue))
    resetValue = resetValue Or &H200
    'Do reset
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
    
    Call Sleep(0.5)
    
    Close_Emg = True
    resetValue = resetValue And &HFDFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
End Sub
