VERSION 5.00
Begin VB.Form frmVisionCalibration 
   Caption         =   "Camera Calibration"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5010
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPrompt 
      Height          =   1455
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Width           =   4095
   End
   Begin VB.TextBox txtMsg 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   4095
   End
   Begin VB.CommandButton buttNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3720
   End
End
Attribute VB_Name = "frmVisionCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String * 4096
Dim step As Long
Dim prevStep As Long

Private Sub buttNext_Click()
    Timer1.Enabled = False
    step = VdeCalibrationDlg(VisionDlgNext, s)

    If (step = VisionDlgToFinish) Then
        step = prevStep + 1
        buttNext.Caption = "Finish"
    ElseIf (step = VisionDlgFinish) Then
        Unload Me
        Exit Sub
    Else
        buttNext.Caption = "Next"
    End If

    prevStep = step
    Me.Caption = "Camera Calibration: Step " + CStr(step)
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    step = VdeCalibrationDlg(VisionDlgInit, s)
    txtMsg.Text = "Step " & CStr(step)
    txtPrompt = CStr(step)
    s = ""
    prevStep = 1
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
    step = VdeCalibrationDlg(VisionDlgCancel, s)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    step = VdeCalibrationDlg(VisionDlgOnTimer, s)
    txtMsg.Text = s
    Timer1.Enabled = True
End Sub

