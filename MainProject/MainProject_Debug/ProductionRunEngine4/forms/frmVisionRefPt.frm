VERSION 5.00
Begin VB.Form frmVisionRefPt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   4665
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   3480
   End
   Begin VB.CommandButton buttNext 
      Caption         =   "Next"
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.TextBox txtMsg 
      Enabled         =   0   'False
      Height          =   1695
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   4095
   End
   Begin VB.TextBox txtPrompt 
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmVisionRefPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String * 4096
Dim step As Long
Dim prevStep As Long

Private Sub buttNext_Click()
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim tempStr As String


    Timer1.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgNext, s)

    If (step = VisionDlgToFinish) Then
        step = prevStep + 1
        buttNext.Caption = "Finish"
    ElseIf (step = VisionDlgFinish) Then
        VdeGetRefPtPos x1, y1, x2, y2
        'editorForm.txtRef1x = x1
        'editorForm.txtRef1y = y1
        'editorForm.txtRef2x = x2
        'editorForm.txtRef2y = y2
        
        tempStr = "fudicial(x=" & convertToPulses(x1, X_axis) & ", y=" & convertToPulses(y1, Y_axis) & "; x=" & convertToPulses(x2, X_axis) & ", y=" & convertToPulses(y2, Y_axis) & ")"
        
        Call editorForm.lstPattern.AddItem(tempStr, 0)


        Unload Me
        Exit Sub
    Else
        buttNext.Caption = "Next"
    End If

    prevStep = step
    Me.Caption = "Teach Ref Pt: Step " + CStr(step)
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    step = VdeTeachRefPtDlg(VisionDlgInit, s)
    txtMsg.Text = "Step " & CStr(step)
    txtPrompt = CStr(step)
    s = ""
    prevStep = 1
    Timer1.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    Timer1.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgCancel, s)
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgOnTimer, s)
    txtMsg.Text = s
    Timer1.Enabled = True
End Sub


