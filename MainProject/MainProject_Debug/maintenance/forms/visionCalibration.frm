VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form visionCalibration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vision Calibration"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   696
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   2  'CenterScreen
   Tag             =   "2"
   Begin VB.CheckBox BothNeedle 
      Caption         =   "Calibrate Both Needle"
      Height          =   255
      Left            =   7560
      TabIndex        =   50
      Top             =   480
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel txtOffsetCalibrationProcedure 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "visionCalibration.frx":0000
      TabIndex        =   49
      Top             =   6960
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7320
      TabIndex        =   48
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSavePosition 
      Caption         =   "Save C_Position"
      Height          =   375
      Left            =   5400
      TabIndex        =   47
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCameraNeedleOffset 
      Caption         =   "Camera_Needle_Offset"
      Height          =   495
      Left            =   5040
      TabIndex        =   46
      Top             =   360
      Width           =   2175
   End
   Begin VB.Frame NeedleMode 
      Caption         =   "Needle Mode"
      Enabled         =   0   'False
      Height          =   615
      Left            =   840
      TabIndex        =   43
      Top             =   7920
      Width           =   2535
      Begin VB.OptionButton LeftNeedle 
         Caption         =   "Left "
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton RightNeedle 
         Caption         =   "Right "
         Height          =   255
         Left            =   1440
         TabIndex        =   44
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton xMinus 
      Height          =   615
      Left            =   6000
      Picture         =   "visionCalibration.frx":005E
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton yPlus 
      Height          =   615
      Left            =   5280
      Picture         =   "visionCalibration.frx":042A
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7560
      Width           =   735
   End
   Begin VB.CommandButton yMinus 
      Height          =   615
      Left            =   5280
      Picture         =   "visionCalibration.frx":0814
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8760
      Width           =   735
   End
   Begin VB.CommandButton xPlus 
      Height          =   615
      Left            =   4560
      Picture         =   "visionCalibration.frx":0C1E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton zMinus 
      Height          =   615
      Left            =   8280
      Picture         =   "visionCalibration.frx":0FF9
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8760
      Width           =   735
   End
   Begin VB.CommandButton zPlus 
      Height          =   615
      Left            =   8280
      Picture         =   "visionCalibration.frx":1403
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7560
      Width           =   735
   End
   Begin VB.Frame JoggingMode 
      Caption         =   "Jogging Mode"
      Height          =   615
      Left            =   840
      TabIndex        =   22
      Top             =   8640
      Width           =   2535
      Begin VB.OptionButton Jogging 
         Caption         =   "Jog"
         Height          =   255
         Left            =   1320
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton JoggingStep 
         Caption         =   "Step"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox StepDistance 
      Height          =   285
      Left            =   1920
      TabIndex        =   21
      Text            =   "1.000"
      Top             =   9480
      Width           =   1095
   End
   Begin VB.Timer setfocusTimer2 
      Interval        =   1000
      Left            =   9360
      Top             =   4440
   End
   Begin VB.CommandButton fudicialCalibration 
      Caption         =   "Fiducial Calibration"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Left            =   3360
      OleObjectBlob   =   "visionCalibration.frx":17ED
      TabIndex        =   20
      Top             =   600
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   2040
      OleObjectBlob   =   "visionCalibration.frx":184D
      TabIndex        =   19
      Top             =   600
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   720
      OleObjectBlob   =   "visionCalibration.frx":18AD
      TabIndex        =   18
      Top             =   600
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel LabelWarning 
      Height          =   255
      Left            =   5280
      OleObjectBlob   =   "visionCalibration.frx":190D
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton abortNeedleOffset 
      Caption         =   "Abort"
      Height          =   375
      Left            =   7320
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer displayCoOrdsTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9360
      Top             =   3840
   End
   Begin VB.TextBox xCoOrd 
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox yCoOrd 
      Height          =   285
      Left            =   2280
      TabIndex        =   12
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox zCoOrd 
      Height          =   285
      Left            =   3600
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.SkinLabel txtMsg 
      Height          =   495
      Left            =   600
      OleObjectBlob   =   "visionCalibration.frx":19DF
      TabIndex        =   7
      Top             =   7080
      Width           =   4215
   End
   Begin VB.Timer fudicial 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   10080
      Top             =   3840
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "visionCalibration.frx":1A3D
      TabIndex        =   5
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox LightingIntensity 
      Height          =   285
      Left            =   600
      TabIndex        =   4
      Text            =   "20"
      Top             =   6480
      Width           =   495
   End
   Begin VB.CommandButton needleCalibration 
      Caption         =   "Needle Calibration"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   10080
      OleObjectBlob   =   "visionCalibration.frx":1AAF
      Top             =   1800
   End
   Begin ComCtl2.UpDown UpDownLighting 
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   6480
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      BuddyControl    =   "LightingIntensity"
      BuddyDispid     =   196634
      OrigLeft        =   504
      OrigTop         =   656
      OrigRight       =   520
      OrigBottom      =   673
      Max             =   255
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton needleCalibrationNext2 
      Caption         =   "Next"
      Height          =   375
      Left            =   5400
      TabIndex        =   9
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton needleCalibrationNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton buttFudNext 
      Caption         =   "Next"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton needleCalibration3 
      Caption         =   "Next"
      Height          =   375
      Left            =   5400
      TabIndex        =   10
      Top             =   6840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel LimitReachedLabel 
      Height          =   255
      Left            =   360
      OleObjectBlob   =   "visionCalibration.frx":1CE3
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Robot Position"
      Height          =   855
      Left            =   360
      TabIndex        =   17
      Top             =   240
      Width           =   4215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   9240
      OleObjectBlob   =   "visionCalibration.frx":1D9F
      TabIndex        =   31
      Top             =   9600
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   4080
      OleObjectBlob   =   "visionCalibration.frx":1E03
      TabIndex        =   32
      Top             =   9600
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   3480
      OleObjectBlob   =   "visionCalibration.frx":1E67
      TabIndex        =   33
      Top             =   9840
      Width           =   735
   End
   Begin ComctlLib.Slider jogSpeedSlider 
      Height          =   375
      Left            =   4080
      TabIndex        =   34
      Top             =   9840
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      Min             =   2
      Max             =   151
      SelStart        =   28
      TickStyle       =   2
      Value           =   28
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
      Height          =   255
      Index           =   0
      Left            =   8040
      OleObjectBlob   =   "visionCalibration.frx":1ED7
      TabIndex        =   35
      Top             =   7800
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
      Height          =   255
      Index           =   0
      Left            =   5520
      OleObjectBlob   =   "visionCalibration.frx":1F37
      TabIndex        =   36
      Top             =   7320
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
      Height          =   255
      Index           =   0
      Left            =   4320
      OleObjectBlob   =   "visionCalibration.frx":1F97
      TabIndex        =   37
      Top             =   8400
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
      Height          =   255
      Index           =   1
      Left            =   6840
      OleObjectBlob   =   "visionCalibration.frx":1FF7
      TabIndex        =   38
      Top             =   8400
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
      Height          =   255
      Index           =   1
      Left            =   5520
      OleObjectBlob   =   "visionCalibration.frx":2057
      TabIndex        =   39
      Top             =   9480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
      Height          =   255
      Index           =   1
      Left            =   8040
      OleObjectBlob   =   "visionCalibration.frx":20B7
      TabIndex        =   40
      Top             =   9000
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel LabelDistance 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "visionCalibration.frx":2117
      TabIndex        =   41
      Top             =   9480
      Width           =   1095
   End
   Begin ComCtl2.UpDown UpDownStep 
      Height          =   255
      Left            =   3000
      TabIndex        =   42
      Top             =   9480
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.PictureBox picImage 
      Height          =   4410
      Left            =   2040
      ScaleHeight     =   290
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   0
      Top             =   1560
      Width           =   5640
   End
End
Attribute VB_Name = "visionCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim s As String * 4096
Dim step As Long
Dim prevStep As Long
Dim Camera_Position(0 To 1) As Long, Left_Position(0 To 1) As Long, Right_Position(0 To 1) As Long, Left_ZPosition As Long, Right_ZPosition As Long
Dim L_Offset_Dis(0 To 1) As Long, R_Offset_Dis(0 To 1) As Long
Dim LeftValve As Boolean            'Just a flag for choosing left-valve.
Dim RightValve As Boolean           'Just a flag for choosing right-valve.

Private Sub displayCoOrds()

    Dim xValue, yValue, zValue, uValue As Long

    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
    
    xCoOrd.Text = convertToMM(xValue, X_axis)
    yCoOrd.Text = convertToMM(yValue, Y_axis)
    zCoOrd.Text = convertToMM(zValue, Z_axis)
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, xValue, yValue, zValue, uValue))
    xValue = xValue And &HC
    yValue = yValue And &HC
    zValue = zValue And &HC
    
    If (xValue <> 0) Or (yValue <> 0) Or (zValue <> 0) Then
        LimitReachedLabel.Visible = True
    Else
        LimitReachedLabel.Visible = False
    End If
    
End Sub

Private Sub buttFudNext_Click()
    fudicial.Enabled = False
    step = VdeCalibrationDlg(VisionDlgNext, s)
    
    'Added by XW (Change the slider's value)
    If (buttFudNext.Caption = "Finish") Then
        jogSpeedSlider.value = 28
        cmdCameraNeedleOffset.Visible = True
    End If

    If (step = VisionDlgToFinish) Then
        step = prevStep + 1
        buttFudNext.Caption = "Finish"
        txtMsg.Caption = "Calibration completed! To save data, click Finish->"
        buttFudNext.Refresh
        'needleCalibration.Visible = True 'xiong
    ElseIf (step = VisionDlgFinish) Then
        fudicial.Enabled = False
        buttFudNext.Visible = False
        buttFudNext.Caption = "Next"
        step = VdeCalibrationDlg(VisionDlgCancel, s)
        'needleCalibration.Visible = True
        'LabelWarning.Visible = True
        fudicialCalibration.Visible = True
        abortNeedleOffset.Visible = False
        txtMsg.Caption = ""
        'Do set focus vision form again
        Me.SetFocus        'XW
        Exit Sub
    Else
        buttFudNext.Caption = "Next"
    End If

    prevStep = step
    fudicial.Enabled = True
    'mainForm.purgeButtonTimer.Enabled = True
    
End Sub

Private Sub displayCoOrdsTimer_Timer()
    displayCoOrds
End Sub

Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If (KeyCode = vbKeyRight) And (Indicator = True) Then
        xMinus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, X_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) + CDbl(StepDistance.Text), X_axis), 0, 0, 0)
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
            Loop
        End If
    ElseIf (KeyCode = vbKeyLeft) And (Indicator = True) Then
        xPlus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, X_axis, 1)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) - CDbl(StepDistance.Text), X_axis), 0, 0, 0)
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
            Loop
        End If
    ElseIf (KeyCode = vbKeyUp) And (Indicator = True) Then
        yPlus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Y_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) - CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0)
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS) 'Loop while Y motor is still spinning
            Loop
        End If
    ElseIf (KeyCode = vbKeyDown) And (Indicator = True) Then
        yMinus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Y_axis, 2)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) + CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0)
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS)  'Loop while Y motor is still spinning
            Loop
        End If
    ElseIf (KeyCode = vbKeyUp) And (Reflector = True) Then
        zPlus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Z_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) - CDbl(StepDistance.Text), Z_axis)) * (-1), 0)
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
            Loop
        End If
    ElseIf (KeyCode = vbKeyDown) And (Reflector = True) Then
        zMinus.SetFocus
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Z_axis, 4)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) + CDbl(StepDistance.Text), Z_axis)) * (-1), 0)
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
            Loop
        End If
    End If
    
    If Shift = vbShiftMask + vbCtrlMask Then
        MsgBox ("Please don't press the two keys at the same time!")
        Exit Sub
    ElseIf (KeyCode = 17) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS) Then
            Exit Sub
        End If
        Reflector = False
        Indicator = True
        xPlus.SetFocus          'Just for testing or change the focus       'XW
    ElseIf (KeyCode = 16) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS) Then
            Exit Sub
        End If
        Indicator = False
        Reflector = True
        zPlus.SetFocus          'Just for testing or change the focus       'XW
    End If
End Sub

Private Sub LeftNeedle_Click()
    Move_To_Zero2
    
    'Do it when changing from the right valve.
    If (LeftValve = False) And (RightValve = True) Then
        LeftNeedleValve
    End If
    
    LeftValve = True
    RightValve = False
End Sub

Private Sub RightNeedle_Click()
    Move_To_Zero2
    
    'Do it when changing from the left valve.
    If (LeftValve = True) And (RightValve = False) Then
        RightNeedleValve
    End If
    
    LeftValve = False
    RightValve = True
End Sub

Private Sub Move_To_Zero2()
    jogSpeedSlider.value = 28
    setSpeed (28)
        
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS)
        DoEvents
    Loop
End Sub

Private Sub setfocusTimer2_Timer()
    setfocusTimer2.Enabled = False
    
    Dim xValue, yValue, zValue, uValue As Long
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, xValue, yValue, zValue, uValue))
    xValue = xValue And &HC
    yValue = yValue And &HC
    zValue = zValue And &HC
        
    If (xValue <> 0) Or (yValue <> 0) Or (zValue <> 0) Then
        Me.SetFocus
    End If
    setfocusTimer2.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Indicator = True) Then
        If KeyCode = vbKeyRight Then
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf KeyCode = vbKeyLeft Then
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf KeyCode = vbKeyUp Then
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = True
        ElseIf KeyCode = vbKeyDown Then
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = True
        Else
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = False
        End If
    ElseIf (Reflector = True) Then
        If KeyCode = vbKeyUp Then
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf KeyCode = vbKeyDown Then
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Then
            Indicator = False
            Reflector = True
        Else
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = False
        End If
    Else
        checkSuccess (P1240MotStop(boardNum, X_axis, 1))
        checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
        checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
        Indicator = False
        Reflector = False
    End If
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn") 'for login (NNO)
    
    Skin1.ApplySkin Me.hWnd
    
    picImage.Width = 461
    picImage.Height = 346
    
    mainForm.SetFocusTimer.Enabled = False
    
    Dim A As Integer
    'displayCoOrdsTimer.Enabled = False 'xiong
    'A = VdeInitializeVision(picImage.hWnd, 461, 346)
    'NNO (login)
    A = VdeInitializeVision(picImage.hWnd, 461, 346, 1)
    
    VdeSelectCamera 2
    VdeCameraLive 1
    
    Initialize_LightIntensity_Com        'for lightIntensity
    Call Turn_On_LightIntensity
    SetLightIntensity (LightingIntensity.Text)
    
    'VdeSetLightIntensity (LightingIntensity.Text)

    displayCoOrdsTimer.Enabled = True
    setfocusTimer2.Enabled = True
    
    LeftNeedleValve                         'Set Default as Let_Cylinder
    LeftValve = True                        'Just a default flag
    Jogging.value = True                    'XW
    JoggingStep.value = False               'XW
    UpDownStep.Enabled = False              'XW
    StepDistance.Enabled = False            'XW
    LabelDistance.Enabled = False           'XW
    visionCalibration.KeyPreview = True
    
    loadVisionCalibration = True
    
End Sub

Private Sub Form_Unload(cancel As Integer)
        
     
'    step = VdeCalibrationDlg(VisionDlgCancel, s)
'    VdeSelectCamera 2
    picImage.Enabled = False 'NNO
    AbortNeedleOffset_Click
    
    Call Sleep(0.3)
    
    fudicial.Enabled = False
    setfocusTimer2.Enabled = False
    displayCoOrdsTimer.Enabled = False
    
    VdeCameraLive False
    Call Sleep(0.2)
    VdeReleaseVision
    Call Sleep(0.8)
    
    
    Call Turn_Off_LightIntensity
    
    If mainForm.mscomLighIntensity.PortOpen = True Then
        mainForm.mscomLighIntensity.PortOpen = False
    End If
    
End Sub

Private Sub fudicialCalibration_Click()
    'mainForm.purgeButtonTimer.Enabled = False
    'frmVisionCalibration.Show vbModal
    txtMsg.Visible = True
    
    VdeSelectCamera 2
    step = VdeCalibrationDlg(VisionDlgCancel, s)
    step = VdeCalibrationDlg(VisionDlgInit, s)
    'txtMsg.Caption = "Step " & CStr(step)
    s = ""
    prevStep = 1
    abortNeedleOffset.Visible = True
    fudicial.Enabled = True
    buttFudNext.Visible = True
    'needleCalibration.Visible = False
    fudicialCalibration.Visible = False
    LightingIntensity.Text = VdeGetLightIntensity
    
    cmdCameraNeedleOffset.Visible = False
End Sub

Private Sub LightingIntensity_Change()
   'VdeSetLightIntensity LightingIntensity.Text
   
   If IsNumeric(LightingIntensity.Text) Then
        If (CInt(LightingIntensity.Text) < 256) And CLng(LightingIntensity.Text) > 0 Then
             SetLightIntensity (Val(LightingIntensity.Text))
        Else
            MsgBox ("Error input value!")
            LightingIntensity.Text = 20
        End If
    
    Else
        MsgBox ("Error input value!")
        LightingIntensity.Text = 20
    End If
End Sub

Private Sub needleCalibration_Click()

    'mainForm.purgeButtonTimer.Enabled = False
    needleCalibration.Enabled = False

    VdeSelectCamera 1
    'step = VdeCalibrationDlg(VisionDlgCancel, s)

    txtMsg.Caption = "Mount syringe with needle tip resting on datum, click Next->"

    setSpeed mainForm.xyDefaultSpeed.Text

    Dim xDatum, yDatum, zDatum, systemMoveHeight As Long
    
    '29July 05
    'xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    'yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    'zDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zDatum", "0"), Z_axis)
    
    xDatum = convertToPulses(mainForm.xDatum.Text, X_axis)
    yDatum = convertToPulses(mainForm.yDatum.Text, Y_axis)
    zDatum = convertToPulses(mainForm.zDatum.Text, Z_axis)
    systemMoveHeight = convertToPulses(mainForm.systemMoveHeight.Text, Z_axis)
    
    
    systemTrackMoveHeight = systemMoveHeight
    
    'Oirgin (XW)
    'PTPToXYZ xDatum, yDatum, systemMoveHeight
    
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)    'XW
    PTPToXYZ xDatum, yDatum, zDatum
    
    txtMsg.Visible = True
    abortNeedleOffset.Visible = True
    needleCalibration.Visible = False
    'LabelWarning.Visible = False
    fudicialCalibration.Visible = False
    needleCalibrationNext.Visible = True
    LightingIntensity.Text = VdeGetLightIntensity

End Sub

Private Sub needleCalibrationNext_Click()
    txtMsg.Caption = "Jog needle tip to Needle Calibration Camera and adjust height to focus, click Next->"
    needleCalibrationNext.Visible = False
    needleCalibrationNext2.Visible = True
End Sub

Private Sub needleCalibrationNext2_Click()

    Dim xValue, yValue, zValue, uValue As Long

    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
    
    SaveStringSetting "EpoxyDispenser", "NeedleOffset", "calibrationX", xValue
    SaveStringSetting "EpoxyDispenser", "NeedleOffset", "calibrationY", yValue
    SaveStringSetting "EpoxyDispenser", "NeedleOffset", "calibrationZ", zValue
    
    step = VdeCalibrationDlg(VisionDlgCancel, s)
    step = VdeCalibrationDlg(VisionDlgInit, s)
    'txtMsg.Caption = "Step " & CStr(step)
    txtMsg.Caption = "To calibrate needle, click Next->"
    s = ""
    prevStep = 1
    fudicial.Enabled = True
    
    buttFudNext.Visible = True
    needleCalibrationNext2.Visible = False

End Sub

Private Sub Jogging_Click()
    JoggingStep.value = False
    UpDownStep.Enabled = False
    StepDistance.Enabled = False
    LabelDistance.Enabled = False
    Jogging.value = True
    SkinLabel19.Enabled = True
    SkinLabel18.Enabled = True
    SkinLabel20.Enabled = True
    jogSpeedSlider.Enabled = True
    Call setSpeed(jogSpeedSlider.value - 1)
End Sub

Private Sub JoggingStep_Click()
    Jogging.value = False
    SkinLabel19.Enabled = False
    SkinLabel18.Enabled = False
    SkinLabel20.Enabled = False
    jogSpeedSlider.Enabled = False
    JoggingStep.value = True
    UpDownStep.Enabled = True
    StepDistance.Enabled = True
    LabelDistance.Enabled = True
    Call setSpeed(40)
End Sub

Private Sub StepDistance_Validate(cancel As Boolean)
    Call validateNumber(visionCalibration.StepDistance.Text, visionCalibration.LabelDistance.Caption)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(visionCalibration.StepDistance.Text) <= 0) Then
            visionCalibration.StepDistance.Text = "0.001"
        ElseIf (CDbl(visionCalibration.StepDistance.Text) > 10) Then
            visionCalibration.StepDistance.Text = "10.000"
        Else
            visionCalibration.StepDistance.Text = Format(visionCalibration.StepDistance.Text, "#0.000")
        End If
    End If
End Sub

Private Sub UpDownStep_DownClick()
    If StepDistance.Text <> 0 Then
        StepDistance.Text = CDbl(StepDistance.Text) - 0.001
    End If
End Sub

Private Sub UpDownStep_UpClick()
   StepDistance.Text = CDbl(StepDistance.Text) + 0.001
End Sub

Private Sub xMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value      'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, X_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) + CDbl(StepDistance.Text), X_axis), 0, 0, 0))
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Private Sub xPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value      'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, X_axis, 1))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) - CDbl(StepDistance.Text), X_axis), 0, 0, 0))
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Private Sub xMinus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub xPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub yMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value      'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 2))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) + CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0))
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS)  'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Private Sub yPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value      'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) - CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0))
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS) 'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Private Sub yMinus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub yPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub zMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value     'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 4))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) + CDbl(StepDistance.Text), Z_axis)) * (-1), 0))
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Private Sub zPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'mainForm.purgeButtonTimer.Enabled = False
    If Jogging.value = True Then
        'setSpeed jogSpeedSlider.value      'origin
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) - CDbl(StepDistance.Text), Z_axis)) * (-1), 0))
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Private Sub zMinus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub zPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
    'mainForm.purgeButtonTimer.Enabled = True
End Sub
Private Sub jogSpeedSlider_Change()
    setSpeed (jogSpeedSlider.value - 1)

End Sub

Private Sub fudicial_Timer()
    fudicial.Enabled = False
    step = VdeCalibrationDlg(VisionDlgOnTimer, s)
    'txtMsg.Caption = s
    fudicial.Enabled = True
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(Y)

    VdeOnLButtonDown Button, ix, iy
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(Y)

    VdeOnMouseMove Button, ix, iy
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(Y)

    VdeOnLButtonUp Button, CInt(X), CInt(Y)
End Sub

Private Sub AbortNeedleOffset_Click()
        fudicial.Enabled = False
        step = VdeCalibrationDlg(VisionDlgCancel, s)
        abortNeedleOffset.Visible = False
        txtMsg.Visible = False
        'needleCalibration.Visible = True
        'LabelWarning.Visible = True
        'fudicialCalibration.Visible = True
        needleCalibrationNext.Visible = False
        needleCalibrationNext2.Visible = False
        needleCalibration3.Visible = False
        buttFudNext.Visible = False
        VdeSelectCamera 2
        'needleCalibration.Enabled = True
        'For finished state, it will show wrong caption if we directly click "Abort button" and not to click the "Finish button"    'XW
        If buttFudNext.Caption = "Finish" Then
            buttFudNext.Caption = "Next"
        End If
        'put by XW (Change the slider's value)
        jogSpeedSlider.value = 28
        cmdCameraNeedleOffset.Visible = True
        'mainForm.purgeButtonTimer.Enabled = True
End Sub

Private Sub cmdCameraNeedleOffset_Click()
    Dim Camera_PositionX As Long, Camera_PositionY As Long
    
    cmdSavePosition.Visible = True
    cmdCancel.Visible = True
    fudicialCalibration.Enabled = False
    cmdCameraNeedleOffset.Enabled = False
    txtOffsetCalibrationProcedure.Visible = True
    txtOffsetCalibrationProcedure.Caption = "Move the camera to calibration position then press 'Save C_Position'"
    
    Camera_PositionX = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Camera_PositionX", "0")
    Camera_PositionY = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Camera_PositionY", "0")
    Camera_PositionY = Camera_PositionY * (-1)
    Call RightNeedle_Click

    setSpeed (CLng(mainForm.xyDefaultSpeed.Text))

    PTPToXYZ Camera_PositionX, Camera_PositionY, 0

    Call setSpeed(jogSpeedSlider.value - 1)
End Sub

Private Sub cmdSavePosition_Click()
    Dim Vresult As Boolean
    Dim Vx As Double, Vy As Double
    Dim PositionX As Long, PositionY As Long, PositionZ As Long, PositionU As Long

    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, Lcnt, PositionX, PositionY, PositionZ, PositionU))
    
    If (cmdSavePosition.Caption = "Save C_Position") Then

        'Vresult = VdeFindCameraOffset1(Vx, Vy)
        'Vresult = True

        'If (Vresult = True) Then
            checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, Lcnt, PositionX, PositionY, PositionZ, PositionU))

            Camera_Position(0) = PositionX
            Camera_Position(1) = PositionY
            
            SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Camera_PositionX", Camera_Position(0)
            SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Camera_PositionY", Camera_Position(1)
        
            cmdSavePosition.Caption = "Save L_Position"
            txtOffsetCalibrationProcedure.Caption = "Move the L.Needle to calibration position then press 'Save L_Position'"
        'Else
            'MsgBox "Please try to save camera again."
        'End If
        
        Call LeftNeedle_Click
        LeftNeedle.value = 1
        setSpeed (jogSpeedSlider.value)
    ElseIf (cmdSavePosition.Caption = "Save L_Position") Then
        Left_Position(0) = PositionX
        Left_Position(1) = PositionY
        Left_ZPosition = PositionZ
        
        If BothNeedle.value = 1 Then
            Call RightNeedle_Click
            RightNeedle.value = 1
            cmdSavePosition.Caption = "Save R_Position"
            txtOffsetCalibrationProcedure.Caption = "Move the R.Needle to calibration position then press 'Save R_Position'"
        Else
            cmdSavePosition.Caption = "Cal Offset_Position"
            txtOffsetCalibrationProcedure.Caption = "To finish, please click 'Cal Offset_Position'"
        End If
    ElseIf (cmdSavePosition.Caption = "Save R_Position") Then
        Right_Position(0) = PositionX
        Right_Position(1) = PositionY
        Right_ZPosition = PositionZ
        
        cmdSavePosition.Caption = "Cal Offset_Position"
        txtOffsetCalibrationProcedure.Caption = "To finish, please click 'Cal Offset_Position'"
    ElseIf (cmdSavePosition.Caption = "Cal Offset_Position") Then
        'Save the offset distance for "Left_Needle"
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_L_Needle", Left_Position(0) - Camera_Position(0)
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_L_Needle", Left_Position(1) - Camera_Position(1)
        
        Master_LN_Position = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Master_LN_Position", "0")
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "needleOffsetZ_L", Master_LN_Position - Left_ZPosition
        
        If BothNeedle.value = 1 Then
            'Save the offset distance for "Right_Needle"
            SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_R_Needle", Right_Position(0) - Camera_Position(0)
            SaveStringSetting "EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_R_Needle", Right_Position(1) - Camera_Position(1)
            
            Master_RN_Position = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Master_RN_Position", "0")
            SaveStringSetting "EpoxyDispenser", "NeedleOffset", "needleOffsetZ_R", Master_RN_Position - Right_ZPosition
        End If
        
        jogSpeedSlider.value = 28
        setSpeed (28)

        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS)
            DoEvents
        Loop

        Reset_Cal
    End If
End Sub

Private Sub cmdcancel_Click()
    checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis Or Z_axis, 0))
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS)
    Loop
    
    Reset_Cal
End Sub

Private Sub Reset_Cal()
    cmdSavePosition.Caption = "Save C_Position"
    cmdSavePosition.Visible = False
    cmdCancel.Visible = False
    fudicialCalibration.Enabled = True
    cmdCameraNeedleOffset.Enabled = True
    txtOffsetCalibrationProcedure.Visible = False
End Sub
