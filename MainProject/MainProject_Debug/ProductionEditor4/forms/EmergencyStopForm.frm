VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form EmergencyStopForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Emergency Stop!"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "EmergencyStopForm.frx":0000
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1455
      Left            =   600
      OleObjectBlob   =   "EmergencyStopForm.frx":0234
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "EmergencyStopForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Unload(cancel As Integer)
    Dim resetValue, ValueX, ValueY, ValueZ As Long
    Dim message As String
    'Check whether the e-stop release or not
    'if not, it will show a message to user.
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, ValueX, ValueY, ValueZ, 0))
    ValueX = (ValueX And &H20)
    ValueY = (ValueY And &H20)
    ValueZ = (ValueZ And &H20)
    Do While (ValueX <> 0) Or (ValueY <> 0) Or (ValueZ <> 0)
        message = "Please release the E-Stop button!"
        Call MsgBox(message, vbOKOnly, "Techno Digm")
        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, ValueX, ValueY, ValueZ, 0))
        ValueX = (ValueX And &H20)
        ValueY = (ValueY And &H20)
        ValueZ = (ValueZ And &H20)
    Loop
    
    Unload Me
    'Stop alarm
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, resetValue))
    resetValue = resetValue And &HF7FE
    checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, resetValue))
    
    If (Red_Light = True) Then
        Red_Light = False
        Green_Light = True
    End If
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, resetValue))
    resetValue = resetValue Or &H200
    'Do reset
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
    
    Call Sleep(0.3)
    
    Close_Emg = True
    resetValue = resetValue And &HFDFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
    
    Call Sleep(1)
    
    Servo_On
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    
    With executionForm
    .cmdStartButton.Enabled = False
    .cmdStopButton.Enabled = False
    .cmdResumeButton.Enabled = False
    .abortButton.Enabled = False
    .homeButton.Enabled = False
    .cmdPurgeButton.Enabled = False
    .closeButton.Enabled = False
    .abortButtonTimer.Enabled = False
    .resumeTimer.Enabled = False
    .startButtonTimer.Enabled = False
    .pauseButtonTimer.Enabled = False
    .purgeButtonTimer.Enabled = False
    End With
    
End Sub

