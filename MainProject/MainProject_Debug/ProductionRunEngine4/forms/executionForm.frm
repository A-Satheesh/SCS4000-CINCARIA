VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form executionForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pattern Execution Form"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   471
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   849
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Test Run"
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
      Begin VB.OptionButton Camera 
         Caption         =   "Camera"
         Height          =   255
         Left            =   2300
         TabIndex        =   48
         Top             =   200
         Width           =   855
      End
      Begin VB.OptionButton wetRun 
         Caption         =   "Wet"
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   200
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton dryRun 
         Caption         =   "Dry"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   200
         Width           =   855
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "repeat run"
      Height          =   255
      Left            =   6120
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox AlwaysPurge 
      Caption         =   "Always Purge"
      Height          =   255
      Left            =   9000
      TabIndex        =   45
      Top             =   1770
      Width           =   1455
   End
   Begin VB.Frame NeedleMode 
      Caption         =   "Needle Mode"
      Height          =   495
      Left            =   13920
      TabIndex        =   42
      Top             =   2280
      Width           =   3615
      Begin VB.OptionButton LeftNeedle 
         Caption         =   "Left "
         Height          =   195
         Left            =   240
         TabIndex        =   44
         Top             =   200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton RightNeedle 
         Caption         =   "Right "
         Height          =   195
         Left            =   2300
         TabIndex        =   43
         Top             =   200
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCycleStop 
      Caption         =   "Cycle Stop"
      Height          =   495
      Left            =   11640
      TabIndex        =   41
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox txtPurgeTime 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   12090
      TabIndex        =   40
      Text            =   "1.00"
      Top             =   1770
      Width           =   525
   End
   Begin ACTIVESKINLibCtl.SkinLabel PurgeTime 
      Height          =   255
      Left            =   11100
      OleObjectBlob   =   "executionForm.frx":0000
      TabIndex        =   39
      Top             =   1770
      Width           =   960
   End
   Begin VB.Timer SensorTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6480
      Top             =   0
   End
   Begin VB.Timer resetTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6960
      Top             =   0
   End
   Begin VB.Timer purgeButtonTimer 
      Interval        =   200
      Left            =   4200
      Top             =   0
   End
   Begin VB.CheckBox onlineOption 
      Caption         =   "Conveyor On"
      Height          =   255
      Left            =   7440
      TabIndex        =   38
      Top             =   1770
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.Timer OnlineTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   6000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "executionForm.frx":0076
      TabIndex        =   37
      Top             =   5880
      Width           =   855
   End
   Begin VB.Timer TimerDrawStatus 
      Interval        =   250
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer displayCoOrdsTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   7560
      OleObjectBlob   =   "executionForm.frx":00E8
      TabIndex        =   29
      Top             =   1200
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   9240
      OleObjectBlob   =   "executionForm.frx":0148
      TabIndex        =   28
      Top             =   1200
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   10920
      OleObjectBlob   =   "executionForm.frx":01A8
      TabIndex        =   27
      Top             =   1200
      Width           =   135
   End
   Begin VB.TextBox xCoOrd 
      Height          =   330
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "0"
      Top             =   1170
      Width           =   1455
   End
   Begin VB.TextBox yCoOrd 
      Height          =   330
      Left            =   9360
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "0"
      Top             =   1170
      Width           =   1455
   End
   Begin VB.TextBox zCoOrd 
      Height          =   330
      Left            =   11040
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "0"
      Top             =   1170
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Robot Position"
      Height          =   690
      Left            =   7440
      TabIndex        =   33
      Top             =   960
      Width           =   5175
   End
   Begin VB.PictureBox PictureReady 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   7440
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   26
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox PictureBusy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   8040
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   25
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox PictureError 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   8640
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   24
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton AbortNeedleOffset 
      Caption         =   "Abort"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4080
      TabIndex        =   22
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel fudMsgText 
      Height          =   735
      Left            =   1440
      OleObjectBlob   =   "executionForm.frx":0208
      TabIndex        =   18
      Top             =   6120
      Width           =   2415
   End
   Begin VB.TextBox LightingIntensity 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   16
      Text            =   "20"
      Top             =   6240
      Width           =   495
   End
   Begin VB.PictureBox PicImage 
      Height          =   5130
      Left            =   240
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   469
      TabIndex        =   14
      Top             =   600
      Width           =   7095
   End
   Begin VB.CommandButton PurgePosition 
      Caption         =   "Purge Position"
      Height          =   495
      Left            =   11280
      TabIndex        =   13
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton abortButton 
      Caption         =   "Abort"
      Height          =   495
      Left            =   10320
      TabIndex        =   12
      Top             =   6000
      Width           =   855
   End
   Begin VB.Timer redrawValveOnOff 
      Interval        =   250
      Left            =   5400
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   9840
      OleObjectBlob   =   "executionForm.frx":0266
      TabIndex        =   11
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox PictureValveOnOff 
      Height          =   495
      Left            =   9720
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   0
   End
   Begin VB.Timer startButtonTimer 
      Interval        =   500
      Left            =   1200
      Top             =   0
   End
   Begin VB.Timer pauseButtonTimer 
      Interval        =   500
      Left            =   1800
      Top             =   0
   End
   Begin VB.Timer abortButtonTimer 
      Interval        =   500
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer ContiPathTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3000
      Top             =   0
   End
   Begin VB.Timer eStopTimer 
      Interval        =   100
      Left            =   3600
      Top             =   0
   End
   Begin VB.Timer resumeTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4800
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "executionForm.frx":02CE
      Top             =   4440
   End
   Begin VB.CommandButton closeButton 
      Caption         =   "Close"
      Height          =   495
      Left            =   11640
      TabIndex        =   6
      Top             =   6480
      Width           =   975
   End
   Begin VB.TextBox txtMessage 
      Height          =   2055
      Left            =   7440
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   5175
   End
   Begin VB.CommandButton cmdResumeButton 
      Caption         =   "Resume"
      Height          =   495
      Left            =   9360
      TabIndex        =   4
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdStopButton 
      Caption         =   "Pause"
      Height          =   495
      Left            =   8400
      TabIndex        =   3
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdStartButton 
      Caption         =   "Run"
      Height          =   495
      Left            =   7440
      TabIndex        =   2
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton cmdPurgeButton 
      Caption         =   "Purge"
      Height          =   495
      Left            =   11280
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton homeButton 
      Caption         =   "Home"
      Height          =   495
      Left            =   11280
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin ComCtl2.UpDown UpDownLighting 
      Height          =   255
      Left            =   720
      TabIndex        =   15
      Top             =   6240
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      BuddyControl    =   "LightingIntensity"
      BuddyDispid     =   196635
      OrigLeft        =   504
      OrigTop         =   656
      OrigRight       =   520
      OrigBottom      =   673
      Max             =   255
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel NeedleOffsetErrMsg 
      Height          =   615
      Left            =   1560
      OleObjectBlob   =   "executionForm.frx":0502
      TabIndex        =   21
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
      Height          =   255
      Left            =   7800
      OleObjectBlob   =   "executionForm.frx":05DA
      TabIndex        =   23
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton FindNeedleOffset 
      Caption         =   "Needle Calibration"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   17
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Calibrate 
      Caption         =   "Calibrate"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5640
      TabIndex        =   20
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Calibrate1 
      Caption         =   "Calibrate"
      Height          =   495
      Left            =   5640
      TabIndex        =   19
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      Caption         =   "Failed Fiducial Mode"
      Height          =   495
      Left            =   7440
      TabIndex        =   34
      Top             =   2880
      Width           =   3615
      Begin VB.OptionButton OptionPause 
         Caption         =   "Pause"
         Height          =   195
         Left            =   2300
         TabIndex        =   36
         Top             =   200
         Width           =   1215
      End
      Begin VB.OptionButton OptionSkip 
         Caption         =   "Skip"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   200
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel CycleTime_BoardNo 
      Height          =   375
      Left            =   7440
      OleObjectBlob   =   "executionForm.frx":0652
      TabIndex        =   46
      Top             =   6600
      Width           =   4095
   End
End
Attribute VB_Name = "executionForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim returncode, contiPathArrayIndex, speedPathArrayIndex As Long
Dim contiPathArray As ContiPathData
Dim speedArray() As Long
Dim SignalCounter As Long                   'Count enable signal for purge timer
Dim PrePositionX As Long, PrePositionY As Long, PrePositionZ As Long, PrePositionU   'Save PrePosition
Dim DisablePurgeSignal As Boolean
Dim ClickPurgePosition As Boolean           'To recover the removal first point in translation (XW)
Dim ReStart As Boolean                      'To exit the on line timer
Dim LeftValve As Boolean                    'Just a flag for choosing left-valve.
Dim RightValve As Boolean                   'Just a flag for choosing right-valve.
Dim L_Dispense As Boolean, R_Dispense As Boolean    'Just flag which valve should be dispensed
Dim Button_Purge As Boolean
Dim Rotation_Angle_U As Long                'Save data for rotation
Dim Rotation_Flag As Boolean, Retilt As Boolean, Spray_Valve As Boolean
Dim Start_Form As Boolean                   'To flage the form load is finished
Dim Position(0 To 2) As Long                'Save X,Y and Z position
Dim Start_Time                              'To calculate the cycle time
Dim Board_Count As Long                     'Count the no of board
Dim Previous_Angle As Long                  'Save the angle and compare with the curren one
Dim firstTimeLeft As Boolean            'Flag for part array to avoid Up/Down many times
Dim firstTimeRight As Boolean           'Flag for part array to avoid Up/Down many times
Dim purgePartArray As Boolean           'Flag for auto purging not to do "Enable/Disable GUI" in partArray

Private Sub AbortNeedleOffset_Click()
    FudMsgText.Visible = False
    AbortNeedleOffset.Visible = False
    NeedleOffsetErrMsg.Visible = False
    Calibrate.Visible = False
    Calibrate1.Visible = False
    'FindNeedleOffset.Visible = True
    'VdeSelectCamera 2
    'VdeSetLightIntensity camera2LightSetting
    'LightingIntensity.Text = camera2LightSetting
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0")

    'FindNeedleOffset.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    abortButton.Enabled = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdStartButton.Enabled = True
    homeButton.Enabled = True
    startButtonTimer.Enabled = True
    PurgePosition.Enabled = True
End Sub

Private Sub Calibrate_Click()
   
    Dim xDatum, yDatum, zDatum As Long
   
    'camera2LightSetting = LightingIntensity.Text
   
    'VdeSelectCamera 1
    setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    
    systemTrackMoveHeight = systemMoveHeight
    
    PTPToXYZ xDatum, yDatum, systemMoveHeight
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), systemMoveHeight
    systemTrackMoveHeight = GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationZ", "0")

    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationZ", "0")
    Calibrate.Visible = False
    FindNeedleOffset.Visible = False
    'Calibrate1.Visible = True
    'VdeSelectCamera 1
    'VdeReadSettings ("VisionSetup.txt")
    'LightingIntensity.Text = VdeGetLightIntensity
    Calibrate1_Click
End Sub

Private Sub Calibrate1_Click()
    Dim OffSetX, OffSetY As Double
    
    'returncode = VdeFindNeedleOffset(offsetX, offsetY)
    
    If returncode = 1 Then
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "XOff", OffSetX
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "YOff", OffSetY
        needleOffsetX = -convertToPulses(OffSetX, X_axis)
        needleOffsetY = -convertToPulses(OffSetY, Y_axis)
        AbortNeedleOffset_Click
    Else
        'AbortNeedleOffset.Visible = True
        NeedleOffsetErrMsg.Visible = True
        FudMsgText.Visible = False
    End If
End Sub

Private Sub Camera_Click()
    SetLightIntensity (Val(LightingIntensity.Text))
End Sub

Private Sub cmdCycleStop_Click()
    PrintParseTree ("Cycle Stop!")
    CycleStop = True
    cmdCycleStop.Enabled = False
    cmdStopButton.Enabled = False       'Should not be pressed more than one time (XW)
    pauseButtonTimer.Enabled = False
    cmdResumeButton.Enabled = True
End Sub

Private Sub displayCoOrdsTimer_Timer()
    'XW
    If (CloseBoard = False) Then
        displayCoOrds
        
        'Check and show tower light
        Tower_Light
    End If
End Sub

Private Sub dryRun_Click()
    SetLightIntensity (0)
End Sub

Private Sub FindNeedleOffset_Click()
        
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    
    FindNeedleOffset.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    abortButton.Enabled = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdStartButton.Enabled = False
    homeButton.Enabled = False
    PurgePosition.Enabled = False
    
    Dim xDatum, yDatum, zDatum As Long
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    zDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zDatum", "0"), Z_axis)

    'Patch to remove up down movement of z-axis 19/1/06
    systemMoveHeight = 0
    systemTrackMoveHeight = systemMoveHeight
    PTPToXYZ xDatum, yDatum, systemMoveHeight
    
    systemTrackMoveHeight = zDatum
    PTPToXYZ xDatum, yDatum, zDatum
    FindNeedleOffset.Visible = False
    'Calibrate.Visible = True
    FudMsgText.Visible = True
    FudMsgText.Caption = "Mount syring with needle tip resting on datum, click Calibrate->"
    'AbortNeedleOffset.Visible = True
End Sub

Private Sub LightingIntensity_Change()
    If IsNumeric(LightingIntensity.Text) Then
        If CLng(LightingIntensity.Text) < 256 And CLng(LightingIntensity.Text) > 0 Then
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

Private Sub PrintParseTree(Text As String)
    txtMessage.Text = txtMessage.Text & Text & vbNewLine
    txtMessage.SelStart = 65535
End Sub

'These two procedures may not be used
'Private Sub LeftNeedle_Click()
'    'Do it when changing from the right valve.
'    If (LeftValve = False) And (RightValve = True) Then
'        LeftNeedleValve
'    End If
    
'    LeftValve = True
'    RightValve = False
'End Sub

'Private Sub RightNeedle_Click()
'    'Do it when changing from the left valve.
'    If (LeftValve = True) And (RightValve = False) Then
'        RightNeedleValve
'    End If
    
'    LeftValve = False
'    RightValve = True
'End Sub

Private Sub abortButton_Click()
    returncode = P1240MotStop(boardNum, X_axis Or Y_axis Or Z_axis, X_axis Or Y_axis Or Z_axis)
    'XW
    'After pressing "Abort" button, we need to check whether all axes are still busy or not.
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
    Loop
    abortValue = True
    If cmdResumeButton.Enabled = True Then
        stopValue = False
        cmdResumeButton.Enabled = False
        cmdResumeButton.Refresh
        pauseButtonTimer.Enabled = True     'XW
    End If
    PrintParseTree ("Dispensing Abort!")
    abortButton.Enabled = False
    
    prevDispenserValue = False
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub

Private Sub cmdPurgeButton_Mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Origin (NYP)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H0)
    
    'Dim Value As Long
    
    'returncode = P1240MotRdReg(boardNum, Z_axis, WR3, Value)
    'Value = (Value And &HFEFF)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, Value)
    'purgeButtonTimer.Enabled = True
    
    Button_Purge = False
End Sub

Private Sub cmdPurgeButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Origin (NYP)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H100)
    
    'Remove dispensing timer
    'Dim Value As Long
    'Dim Start

    'Start = Timer
    
    'purgeButtonTimer.Enabled = False
    'Do While (Timer < Start + CDbl(txtPurgeTime.Text))
    '    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, Value)
    '    Value = (Value Or &H100)
    '    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, Value)
    '    DoEvents
    'Loop
    
    Dim xpos, ypos, zpos, ReadValue As Long
    
    purgeButtonTimer.Enabled = False
    Disable_Button
    Button_Purge = True
    
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
       
    xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
    ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
    'To get the actual direction        'XW
    ypos = ypos * (-1)
    zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
    zpos = zpos * (-1)
    
    'XW
    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    systemMoveHeight = systemMoveHeight * (-1)
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.3)
    
    'To travel the right way (XW)
    If (systemMoveHeight > zpos) Then
        doPtp xpos, ypos, systemMoveHeight, zSystemTravelSpeed
    Else
        ClickPurgePosition = True
    End If
    
    doPtp xpos, ypos, zpos, xySystemTravelSpeed
    'Left slider go down
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.2)
    
    Dispensing
End Sub

Private Sub Dispensing()
    Dim ReadValue As Long

    Do While (Button_Purge = True)
        If leftside = False And rightside = True Then
            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H100
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
        ElseIf leftside = True And rightside = False Then
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H800
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
        Else
            'If (executionForm.LeftNeedle.Value = True) Then
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H800
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
            'ElseIf (executionForm.RightNeedle.Value = True) Then
            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H100
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
            'End If
        End If
        DoEvents
    Loop
    
    returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HF7FF
    returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
    
    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HFEFF
    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
    
    readyStatus = False
    busyStatus = True
    
    Call Sleep(0.2)
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.1)
    
    doPtp PrePositionX, PrePositionY, systemMoveHeight, zSystemTravelSpeed
    doPtp PrePositionX, PrePositionY, PrePositionZ, xySystemTravelSpeed
    readyStatus = True
    busyStatus = False
        
    'Left slider go down
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    Enable_Button
    purgeButtonTimer.Enabled = True
End Sub

Private Sub cmdstartbutton_Click()
    If (Inter_Lock = True) Then
        If (Door_Lock = True) Then
            MsgBox "Please lock the door first before running the application!"
            Exit Sub
        End If
    End If
    
    If onlineOption.Value = 1 Then
        'XW
        'onlineOption.Enabled = False
        OnlineTimer.Enabled = True
    Else
        'XW
        ReStart = True
        Start_Time = Timer
        
        If (AlwaysPurge.Value = 1) Then
            Always_Purge
        End If
        
        doDispense
        'temp NNO
        If Check1.Value = 1 Then
        cmdstartbutton_Click
        End If
    End If
    
End Sub
    
Private Sub doDispense()

    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    systemMoveHeight = systemMoveHeight * (-1)
    
    abortValue = False
    readyStatus = False     'For indicator (XW)
    busyStatus = True
        
    'To prevent flickering of start button
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    
    closeButton.Enabled = False
    homeButton.Enabled = False
    FindNeedleOffset.Enabled = False
    
    cmdPurgeButton.Enabled = False
    PurgePosition.Enabled = False
    abortButton.Enabled = True
    
    Call doNormalDispense
    
    'cmdStartButton.Enabled = False
    'cmdStopButton.Enabled = False
    'cmdResumeButton.Enabled = False
    'abortButton.Enabled = False
    'homeButton.Enabled = False
    'cmdPurgeButton.Enabled = False
    'closeButton.Enabled = False
    
    'doHome
    
    cmdStartButton.Enabled = True
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    PurgePosition.Enabled = True
    closeButton.Enabled = True
    
    'To prevent flickering of start button
    startButtonTimer.Enabled = True
    purgeButtonTimer.Enabled = True
    
    'FindNeedleOffset.Enabled = True        'done by XW
End Sub

Private Sub cmdResumeButton_Click()

    PrintParseTree ("Dispensing Resume...")
    stopValue = False
    CycleStop = False           'XW
    cmdResumeButton.Enabled = False
    pauseButtonTimer.Enabled = True

End Sub

Private Sub cmdStopButton_Click()
    PrintParseTree ("Dispensing Pause!")
    stopValue = True
    cmdStopButton.Enabled = False       'Should not be pressed more than one time (XW)
    pauseButtonTimer.Enabled = False
    cmdResumeButton.Enabled = True
    resumeTimer.Enabled = True
End Sub
Private Sub Form_Activate()
    SetWindowOnTop Me, True   '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False   '@$K
End Sub

Private Sub Form_Load()
    
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    
    'End all of the program after pressing the E-Stop button    'XW
    If (Close_Emg = True) Then
        End
    End If
    
    executionForm.displayCoOrdsTimer.Enabled = True
    executionForm.Caption = CaptionRunEngine
    
    'SensorTimer.Enabled = False
    resetTimer.Enabled = False
       
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H400)
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &H200))  'xu long
    
    abortValue = False
    estopValue = False
    stopValue = False
    
    Red_Light = False
    Yellow_Light = False
    Green_Light = True
    
    startButtonTimer.Enabled = False
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    
    'Testing (Not to do the homing when the door is opening)
    Dim dor As Boolean
    dor = True
    Do While (dor)
        If (Inter_Lock = True) Then
            If (Door_Lock = True) Then
                MsgBox "Please lock the door first before running the application!"
            Else
                dor = False
                Exit Do
            End If
        Else
            dor = False
            Exit Do
        End If
    Loop
    
    'directSoftHomeOption = False
    prevDispenserValue = False

    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    printErrorLimit = True
        
    'Run Engine
    
    Timer1.Enabled = False
    
    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    systemMoveHeight = systemMoveHeight * (-1)
    
    txtPurgeTime.Text = Format(GetStringSetting("EpoxyDispenser", "Setup", "DispensingTime", "10"), "##0.00")
    
    'These two parameters haven't been changed while doing the translation
    'XW
    systemHomeY = systemHomeY * (-1)
    systemHomeZ = systemHomeZ * (-1)
    
    Dim resetX, resetY, resetZ As Long
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, resetX, resetY, resetZ, 0))
    resetX = resetX And &H10
    resetY = resetY And &H10
    resetZ = resetZ And &H10
            
    If (resetX = &H0) And (resetY = &H0) And (resetZ = &H0) Then
        moveToHome
    Else
        MsgBox "Driver Error. Please check the Driver!"
          
        displayCoOrdsTimer.Enabled = False
        startButtonTimer.Enabled = False
        
        Servo_Off
        Close_TowerLight
        ResetDriver
        unInitializeBoard
        
        End
        Exit Sub
    End If
    
    Timer1.Enabled = True
    Leftslider_go_up
    doSoftHome
    Leftslider_go_down
    
    'SensorTimer.Enabled = True
    startButtonTimer.Enabled = True
    resetTimer.Enabled = True
    purgeButtonTimer.Enabled = True
    
    'LeftNeedleValve                         'Set Default as Let_Cylinder
    'LeftValve = True                        'Just a default flag
    'RightValve = False
    
    cmdStartButton.Enabled = True
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    
    deltaX = 0
    deltaY = 0
    deltaAngle = 0
    xOrgFid = 0
    yOrgFid = 0
    
    'Run Engine Additions
    
    PicImage.Width = 461
    PicImage.height = 346
    
    'VdeInitializeVision PicImage.hWnd, 461, 346
    VdeInitializeVision PicImage.hWnd, 461, 346, 3 'NNO
    
    VdeSelectCamera 2
    VdeCameraLive 1
    VdeReadSettings ("VisionSetup.txt")
    
    Initialize_LightIntensity_Com    'for lightintensity   '@$K
    Call Turn_On_LightIntensity
    'LightingIntensity.Text = VdeGetLightIntensity
    
    Start_Form = True
    
    'Initialize the cycle time and the no of board
    Board_Count = 0
    CycleTime_BoardNo.Caption = "The cycle time is 0sec. The successful board is " & Board_Count & "."
End Sub

Private Sub doNormalDispense()
    Dim buttonValue As Long
    
    'Check the metarial before starting the program
    If rightside = True Then
        Low_Level
    End If
        
    'Origin (NYP)
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H0)
    
    'Production mode signal
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H200) 'xu long
    
    'Production mode & Left needle signal
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &HA00))

    cmdStartButton.Enabled = False
    cmdResumeButton.Enabled = False
    closeButton.Enabled = False
    cmdStopButton.Enabled = True
    pauseButtonTimer.Enabled = True
    abortButton.Enabled = True
    NeedleMode.Enabled = False
    abortButtonTimer.Enabled = True
    
    firstTimeLeft = False
    firstTimeRight = False
    
    'If (systemHomeX <> 0 Or systemHomeY <> 0 Or systemHomeZ <> 0) Then
        'Timer1.Enabled = True
        'If (directSoftHomeOption = False) Then
                'Call doPtp(systemHomeX, systemHomeY, systemMoveHeight, 50)
        'End If
    'End If
    Timer1.Enabled = False
     
    PrintParseTree ("Dispensing Begin...")

    Dim docontiloop As Boolean

    Dim retstring, readstring As String
    
    doneFudicial = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.OpenTextFile(App.Path & "\translatedPattern.txt", 1, False)
    docontiloop = False
    retstring = ""
    Do While ((A.AtEndOfStream <> True) And (estopValue = False) And (abortValue = False))
        'Oringin (NYP)
        'readstring = A.ReadLine
        'retstring = retstring & readstring
        'If (readstring = "contibuffer") Then
        '    docontiloop = True
        'End If
    
        'If (readstring = "contiend") Then
        '    docontiloop = False
        'End If
        
        'If (Left(readstring, 3) = "ptp") Then
        '    docontiloop = False
        'End If
        'retstring = retstring & vbNewLine
        'If docontiloop = False Then
        '    DoParse (retstring)
        '    retstring = ""
        'End If
        
        readstring = A.ReadLine
        If (readstring <> "") Then
            Call doExecute(readstring)
            readstring = ""
        End If
    Loop
    A.Close
    
    'Move into "If" function (XW)
    'PrintParseTree ("Dispensing Complete!")
    
    'If the user presses "Abort" button, move "System Height" for safety.
    If (abortValue = True) Then
        Dim xpos, ypos, zpos, upos As Long, ValveClose As Long
        
        checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
        ValveClose = (ValveClose And &HF7FF)
        checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
            
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ValveClose))
        ValveClose = (ValveClose And &HFEFF)
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ValveClose))
        
        prevDispenserValue = False
    End If
    
    Leftslider_go_up
    'Tilting Off
    Move_System_Height
    Tilt_Off
    Rotation_U (0)
    
    'Change the left needle before moving to the system home position (XW)
    'LeftNeedleValve
    'executionForm.LeftNeedle.Value = True

    systemTrackMoveHeight = systemMoveHeight

    If (estopValue = False) Then
        cmdStopButton.Enabled = False
        pauseButtonTimer.Enabled = False
        'startButtonTimer.Enabled = True
        cmdStopButton.Enabled = False
                
        Leftslider_go_up
        'doSoftHome
        If GetStringSetting("EpoxyDispenser", "Setup", "AlwaysRobotHome", "0") = "1" Then
            doPtp 0, 0, 0, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50"))
            doSoftHome
            Leftslider_go_down
        Else
            'Call doPtp(0, 0, systemMoveHeight, 50)
            'Origin
            'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
            
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
            
            Leftslider_go_down
        End If
        
        'Cycle stop (XW)
        Do While (CycleStop = True And estopValue = False And abortValue = False)
            DoEvents
            returncode = P1240MotRdReg(boardNum, X_axis, RR4, buttonValue)
            buttonValue = buttonValue And &H2
            
            If (buttonValue = &H0) Then
                PrintParseTree ("Dispensing Resume...")
                CycleStop = False
                pauseButtonTimer.Enabled = True
            End If
        Loop
        
        PrintParseTree ("Dispensing Complete!")
        
        busyStatus = False
        readyStatus = True
        CycleStop = False
        
        cmdStopButton.Enabled = False
        abortButton.Enabled = False
        
        '''''''''''''''''''''''''''''
        '   Move to end of method.  '
        '                           '
        '''''''''''''''''''''''''''''
        'Enable after finish signal or updating cycle time
        'cmdCycleStop.Enabled = True
        'cmdStartButton.Enabled = True
        'closeButton.Enabled = True
        
        'purgeButtonTimer.Enabled = True
        'startButtonTimer.Enabled = True
        'NeedleMode.Enabled = True
    Else
        cmdStopButton.Enabled = False
        pauseButtonTimer.Enabled = False
        cmdStartButton.Enabled = False
        closeButton.Enabled = False
    End If
    
    'Origin (NYP)
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H400)
    
    'Only conveyor + single valve
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H600)    'XL
    
    'Need to consider "left" or "right" valve (2 head)
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &HE00))
    
    'checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &H400))
    
    If onlineOption.Value = 0 Then
        Call Sleep(1)
        'Need to consider "left" or "right" valve (2 head)
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &HA00))
    End If
    
    'Calculate the cycle time and the no of board
    If (abortValue = False) Then
        Start_Time = Timer - Start_Time
        Board_Count = Board_Count + 1
        If (Board_Count > 1) Then
            CycleTime_BoardNo.Caption = "The cycle time is " & CLng(Start_Time) & "sec. The successful boards are " & Board_Count & "."
        Else
            CycleTime_BoardNo.Caption = "The cycle time is " & CLng(Start_Time) & "sec. The successful board is " & Board_Count & "."
        End If
    Else
        If (Board_Count > 1) Then
            CycleTime_BoardNo.Caption = "The cycle time is 0sec. The successful boards are " & Board_Count & "."
        Else
            CycleTime_BoardNo.Caption = "The cycle time is 0sec. The successful board is " & Board_Count & "."
        End If
    End If
    
    'Check the metarial after finishing the program
    If rightside = True Then
        Low_Level
    End If
    Previous_Angle = 0
    Rotation_Flag = False
    
    If (estopValue = False) Then
        cmdCycleStop.Enabled = True
        cmdStartButton.Enabled = True
        closeButton.Enabled = True
        
        purgeButtonTimer.Enabled = True
        startButtonTimer.Enabled = True
        NeedleMode.Enabled = True
    End If
End Sub

Private Sub DoParse(dataline As String)
      
   Dim Response As GPMessageConstants
   Dim Parser   As New GOLDParser
   Dim Done As Boolean                                    'Controls when we leave the loop
   
   Dim ReductionNumber As Integer                         'Just for information
   Dim n As Integer, Text As String
      
   If Parser.LoadCompiledGrammar(txtCGTFilePath2) Then
       Parser.OpenTextString (dataline)
       Parser.TrimReductions = True
              
       Done = False
       Do Until Done
           Response = Parser.Parse()
              
           Select Case Response
           Case gpMsgLexicalError
              txtMessage.Text = "Line " & Parser.CurrentLineNumber & ": Lexical Error: Cannot recognize token: " & Parser.CurrentToken.Data
              Done = True
                 
           Case gpMsgSyntaxError
              Text = ""
              For n = 0 To Parser.TokenCount - 1
                  Text = Text & " " & Parser.Tokens(n).Name
              Next
              txtMessage.Text = "Line " & Parser.CurrentLineNumber & ": Syntax Error: Expecting the following tokens: " & LTrim(Text)
              Done = True
              
           Case gpMsgReduction
              ReductionNumber = ReductionNumber + 1
              Parser.CurrentReduction.Tag = ReductionNumber   'Mark the reduction
              
           Case gpMsgAccept
              '=== Success!
              doExecute Parser.CurrentReduction
              Done = True
              
           Case gpMsgTokenRead
              
           Case gpMsgInternalError
              Done = True
              
           Case gpMsgNotLoadedError
              Done = True
              
           Case gpMsgCommentError
              Done = True
           End Select
           
        Loop
    Else
        MsgBox "Could not load the CGT file", vbCritical
    End If
    
    Parser.Clear
    Set Parser = Nothing
    
End Sub

'Oringin (NYP)
'Private Sub doExecute(TheReduction As Reduction)
'Dim seqNum As Long, n As Integer, x As Long, y As Long, Z As Long, endx As Long, endy As Long, endz As Long, centerx As Long, centery As Long
'Dim tempEndX As Long, tempEndY As Long
'Dim tempEndX1 As Long, tempEndY1 As Long

'Dim Speed As Long
'Dim delay As Double
'Dim ccw As Long
'Dim amount As Long
'Dim dispenserControl As Boolean

'    Timer1.Enabled = True

'   For n = 0 To TheReduction.TokenCount - 1
'      Select Case TheReduction.Tokens(n).Kind
'       Case SymbolTypeNonterminal
'            doExecute TheReduction.Tokens(n).Data
'       Case Else
'          If (estopValue = True) Or (abortValue = True) Then
'               Exit Sub
'          End If
'          Select Case LCase(TheReduction.Tokens(n).Data)
'            Case "ptp"
'                Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
'                'Origin
'                endx = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data - needleOffsetX
'                'endy = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data - needleOffsetY
'                endy = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data - (needleOffsetY * (-1))
'                endz = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
      
'                'Testing (two valves) XW
'                If (55555 = CLng(endx + needleOffsetX)) And (55555 = CLng(endy - needleOffsetY)) And (endz = 55555) And (Speed = 55555) Then
'                    Move_System_Height
                    
'                    'Old method
'                    'RightNeedleValve
'                    LeftNeedle.Value = False
'                    L_Dispense = True
'                    R_Dispense = False
                    
'                    RightNeedleValve
                    
'                    RightNeedle.Value = True
'                    L_Dispense = False
'                    R_Dispense = True
                    
'                    Spray_Valve = True
'                ElseIf (66666 = CLng(endx + needleOffsetX)) And (66666 = CLng(endy - needleOffsetY)) And (endz = 66666) And (Speed = 66666) Then
'                    Move_System_Height
'
'                    'If (LeftValve = False) Then
'                    '    LeftNeedleValve
'                    'End If
'                    'LeftNeedleValve
'                    RightNeedle.Value = False
'                    R_Dispense = True
'                    L_Dispense = False
                    
'                    LeftNeedleValve
                    
'                    LeftNeedle.Value = True
'                    R_Dispense = False
'                    L_Dispense = True
'
'                    Tilt_Off
                    
'                    If (Spray_Valve = True) Then
'                        Rotation_U (0)
'                    End If
                    
'                    Spray_Valve = False
'                ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 10101) Then
'                    'No tilt
'                    Tilt_Off
'                    If (Spray_Valve = True) Then
'                        Rotation_U (0)
'                    End If
'                    Rotation_Flag = False
'                ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 36363) Then
'                    'MsgBox "Tilt first and then rotate 0 or 360."
'                    'Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
'                    'Loop
'                    Tilt_ON
'                    Rotation_U (0)
'                    Rotation_Flag = True
'                ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 99999) Then
'                    'MsgBox "Tilt first and then rotate -90."
'                    'Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
'                    'Loop
'                    Tilt_ON
'                    'Rotation_U (1250)
'                    Rotation_U (2500)
'                    Rotation_Flag = True
'                ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 18181) Then
'                    'MsgBox "Tilt first and then rotate -180."
'                    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
'                    Loop
'                    Tilt_ON
'                    'Rotation_U (2500)
'                    Rotation_U (5000)
'                    Rotation_Flag = True
'                ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 27272) Then
'                    'MsgBox "Tilt first and then rotate -270."
'                    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
'                    Loop
'                    Tilt_ON
'                    'Rotation_U (3750)
'                    Rotation_U (7500)
'                    Rotation_Flag = True
'                ElseIf (11111 = CLng(endx + needleOffsetX)) And (11111 = CLng(endy - needleOffsetY)) And (endz = 11111) And (Speed = 11111) Then
'                    No_Start_Stop = True
'                    Rotation_Flag = False
'                Else
'                    origX = endx
'                    origY = endy
'
'                    detXYfromFudicial endx, endy, endx, endy, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
'
'                    presentX = endx
'                    presentY = endy
                           
'                    'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
'                    'To recover the removable first point in translation    (XW)
'                    'If (ClickPurgePosition = True) Then
'                    '    Call doPtp(endx, endy, SystemMoveHeight, Speed)
'                    '    ClickPurgePosition = False
'                    'End If
                
'                    'If "Purge Position" is higher than "systemMoveHeight", the robort should take z-Height of "Purge Position".
'                    If (Spray_Valve = False) Then
'                        If (ClickPurgePosition = True) Then
'                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
'                            ClickPurgePosition = False
'                        Else
'                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, endz, Speed)
'                        End If
'                    Else
'                        If (Rotation_Flag = False) Then
'                            If (ClickPurgePosition = True) Then
'                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
'                                ClickPurgePosition = False
'                            Else
'                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, endz, Speed)
'                            End If
'                        Else
'                            If (ClickPurgePosition = True) Then
'                               Call doPtp(endx, endy, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
'                                ClickPurgePosition = False
'                            Else
'                                Call doPtp(endx, endy, endz, Speed)
'                            End If
'                        End If
'                    End If
                
'                    'Call doPtp(endx, endy, endz, Speed)    (origin)
'                End If
'            Case "dispenseon"
'                Call doDispenseOn
'            Case "dispenseoff"
'                Call doDispenseOff
'            Case "delay"
'                delay = TheReduction.Tokens(2).Data.Tokens(0).Data
'                Call doDelay(delay)
'            Case "contibuffer"
'                Call doContiBuffer
'            Case "line"
'                endx = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
'                endy = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
'
'                tempEndX1 = origX + endx
'                tempEndY1 = origY + endy
                
'                tempEndX = tempEndX1
'                tempEndY = tempEndY1
                
'                origX = tempEndX1
'                origY = tempEndY1
                                
'                detXYfromFudicial tempEndX, tempEndY, tempEndX, tempEndY, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
'
'                endx = tempEndX - presentX
'                endy = tempEndY - presentY
                                
'                Call doContiLine(endx, endy)

'                presentX = tempEndX
'                presentY = tempEndY

'            Case "line3d"
'                endx = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
'                endy = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
'                endz = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                
'                'initialize
'                Position(0) = 0
'                Position(1) = 0
'                Position(2) = 0
                
'                tempEndX1 = origX + endx
'                tempEndY1 = origY + endy
                
'                tempEndX = tempEndX1
'                tempEndY = tempEndY1
                
'                origX = tempEndX1
'                origY = tempEndY1
                                
'                detXYfromFudicial tempEndX, tempEndY, tempEndX, tempEndY, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
                                
'                endx = tempEndX - presentX
'                endy = tempEndY - presentY
                
'                'Origin (NYP)
'                If (No_Start_Stop = True) Then
'                    Call doContiLine3D(endx, endy, endz)
'                Else
'                    Call doContiLine3D_XW(endx, endy, endz)
'                End If
            
'                presentX = tempEndX
'                presentY = tempEndY
            
'            Case "arc"
'                endx = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
'                endy = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
'                centerx = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
'                centery = TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data
'                ccw = TheReduction.Tokens(6).Data.Tokens(2).Data
                
'                tempEndX1 = origX + endx
'                tempEndY1 = origY + endy
                
'                tempEndX = tempEndX1
'                tempEndY = tempEndY1
                
'                origX = tempEndX1
'                origY = tempEndY1
                                
'                detXYfromFudicial tempEndX, tempEndY, tempEndX, tempEndY, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
                
'                endx = tempEndX - presentX
'                endy = tempEndY - presentY
                                
'                Call doContiArc(endx, endy, centerx, centery, ccw)
            
'                presentX = tempEndX
'                presentY = tempEndY
            
'            Case "segmentproperty"
'                Speed = TheReduction.Tokens(2).Data.Tokens(2).Data
'                dispenserControl = TheReduction.Tokens(4).Data.Tokens(0).Data
'                seqNum = TheReduction.Tokens(6).Data.Tokens(0).Data
'                Call doSegmentProperty(Speed, dispenserControl, seqNum)
'            Case "segmentproperty3d"
'                Speed = TheReduction.Tokens(2).Data.Tokens(2).Data
'                dispenserControl = TheReduction.Tokens(4).Data.Tokens(0).Data
'                seqNum = TheReduction.Tokens(6).Data.Tokens(0).Data
'                'Origin (NYP)
'                If (No_Start_Stop = True) Then
'                    Call doSegmentProperty3D(Speed, dispenserControl, seqNum)
'                Else
'                    Call doSegmentProperty3D_XW(Speed, dispenserControl, seqNum)
'                End If
'            Case "contiend"
'                Call doContiEnd
'                No_Start_Stop = False
'            Case "fudicial"
'                Call doFudicial(TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data, TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data, TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data, TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data, TheReduction.Tokens(6).Data, TheReduction.Tokens(8).Data.Tokens(0).Data)
'            End Select
'       End Select
'   Next
   
'End Sub

'No error checking
Private Sub doExecute(ByVal ExecutionString As String)
    Dim Speed As Long, seqNum As Long, endx As Long, endy As Long, endz As Long
    Dim tempEndX As Long, tempEndY As Long
    Dim tempEndX1 As Long, tempEndY1 As Long
    Dim delay As Double
    Dim dispenserControl As Integer
    
    Dim word1() As String, word2() As String, word3() As String, word4() As String, _
        word5() As String, word6() As String, word7() As String, word8() As String

    Timer1.Enabled = True

    If (estopValue = True) Or (abortValue = True) Then
        Exit Sub
    End If
    
    word1() = Split(ExecutionString, "(")
    
    Select Case LCase(Trim(word1(0)))
        Case "ptp"
            word2() = Split(ExecutionString, "=")
            word3() = Split(word2(1), ",")
            word4() = Split(word2(2), ",")
            word5() = Split(word2(3), ";")
            word6() = Split(word2(4), ")")
            
            Speed = CLng(Trim(word6(0)))
                
            endx = CLng(Trim(word3(0))) - needleOffsetX
            endy = CLng(Trim(word4(0))) - (needleOffsetY * (-1))
            endz = CLng(Trim(word5(0)))
            
            'Testing (two valves) XW
            If (55555 = CLng(endx + needleOffsetX)) And (55555 = CLng(endy - needleOffsetY)) And (endz = 55555) And (Speed = 55555) Then
                If (firstTimeLeft = False) Then
                    Move_System_Height
                    
                    LeftNeedle.Value = False
                    L_Dispense = True
                    R_Dispense = False
                    
                    RightNeedleValve
                    
                    RightNeedle.Value = True
                    L_Dispense = False
                    R_Dispense = True
                    
'                    Tilt_Off
'                    'Need to go back to "angle 0" because we don't know all offset value
'                    If (Spray_Valve = True) Then
'                        Rotation_U (0)
'                    End If
                
                    Spray_Valve = True    '@$K
                    
                    firstTimeLeft = True
                    firstTimeRight = False
                End If
            ElseIf (66666 = CLng(endx + needleOffsetX)) And (66666 = CLng(endy - needleOffsetY)) And (endz = 66666) And (Speed = 66666) Then
                If (firstTimeRight = False) Then
                    Move_System_Height
                
                    RightNeedle.Value = False
                    R_Dispense = True
                    L_Dispense = False
                    
                    LeftNeedleValve
                    
                    LeftNeedle.Value = True
                    R_Dispense = False
                    L_Dispense = True
                    
                    Tilt_Off
                    'Need to go back to "angle 0" because we don't know all offset value
                    If (Spray_Valve = True) Then
                        Rotation_U (0)
                    End If

                    Spray_Valve = False '@$K
                
                    firstTimeRight = True
                    firstTimeLeft = False
                End If
            ElseIf (endx = 0) And (endy = 0) And (endz = 0) And (Speed = 77777) Then
                'Go to the system home and do the purging for part array
                If (AlwaysPurge.Value = 1) Then
                    purgePartArray = True
                    Always_Purge_Part_Array
                    purgePartArray = False
                End If
            ElseIf (77777 = CLng(endx + needleOffsetX)) And (77777 = CLng(endy - needleOffsetY)) And (endz = 77777) And (Speed = 77777) Then
                'Go to the system home and do the purging for part array
                If (AlwaysPurge.Value = 1) Then
                    purgePartArray = True
                    Always_Purge_Part_Array
                    purgePartArray = False
                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 10101) Then
'                'No tilt
'                Tilt_Off
'                'If (Spray_Valve = True) Then
'                    Rotation_U (0)
'                'End If
'                Rotation_Flag = False
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 36363) Then
'                If (Previous_Angle <> 36363) Then
'                'If (Previous_Angle <> 36363) And (Camera.Value = False) Then
'                    Tilt_ON
'                    Rotation_U (0)
'                    Previous_Angle = 36363
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 99999) Then
'                If (Previous_Angle <> 99999) Then
'                'If (Previous_Angle <> 99999) And (Camera.Value = False) Then
'                    Tilt_ON
'                    'Rotation_U (2500)
'                    Rotation_U (900)
'                    Previous_Angle = 99999
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 18181) Then
'                If (Previous_Angle <> 18181) Then
'                'If (Previous_Angle <> 18181) And (Camera.Value = False) Then
'                    Tilt_ON
'                    'Rotation_U (5000)
'                    Rotation_U (1800)
'                    Previous_Angle = 18181
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 27272) Then
'                If (Previous_Angle <> 27272) Then
'                'If (Previous_Angle <> 27272) And (Camera.Value = False) Then
'                    Tilt_ON
'                    'Rotation_U (7500)
'                    Rotation_U (2700)
'                    Previous_Angle = 27272
'                    Rotation_Flag = True
'                End If
            
            '@$K
            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) Then
                If (Speed = 10101) Then
                    Tilt_Off
                    Rotation_U (0)
                    Rotation_Flag = False
                Else
                    Tilt_ON
                    Rotation_U (Speed * 10)
                    Rotation_Flag = True
                End If
                
            ElseIf (11111 = CLng(endx + needleOffsetX)) And (11111 = CLng(endy - needleOffsetY)) And (endz = 11111) And (Speed = 11111) Then
                No_Start_Stop = True
                Rotation_Flag = False
            Else
                origX = endx
                origY = endy
                                
                detXYfromFudicial endx, endy, endx, endy, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
                
                presentX = endx
                presentY = endy
                
                If (Camera.Value = True) Then
                    endx = CLng(Trim(word3(0)))
                    endy = CLng(Trim(word4(0)))
                    endz = 0
                    
                    If (Spray_Valve = False) Then
                        If (ClickPurgePosition = True) Then
                            Call doPtp(endx, endy, 0, Speed)
                            ClickPurgePosition = False
                        Else
                            Call doPtp(endx, endy, 0, Speed)
                        End If
                    Else
                        If (Rotation_Flag = False) Then
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx, endy, 0, Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx, endy, 0, Speed)
                            End If
                        Else
                            'Tilting/Rotation is needle teach. So, we need to change as camera's position
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx - Offset_DistanceX_Camera_R_Needle, endy - Offset_DistanceY_Camera_R_Needle, 0, Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx - Offset_DistanceX_Camera_R_Needle, endy - Offset_DistanceY_Camera_R_Needle, 0, Speed)
                            End If
                        End If
                    End If
                Else
                    'If "Purge Position" is higher than "systemMoveHeight", the robort should take z-Height of "Purge Position".
                    If (Spray_Valve = False) Then
                        If (ClickPurgePosition = True) Then
                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                            ClickPurgePosition = False
                        Else
                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, endz - needleoffsetZ_L, Speed)
                        End If
                    Else
                        If (Rotation_Flag = False) Then
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, endz - needleoffsetZ_R, Speed)
                            End If
                        Else
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, endz - needleoffsetZ_R, Speed)
                            End If
                        End If
                    End If
                End If
            End If
        Case "dispenseon"
            Call doDispenseOn
        Case "dispenseoff"
            Call doDispenseOff
        Case "delay"
            word2() = Split(word1(1), ")")
        
            delay = CDbl(Trim(word2(0)))
            Call doDelay(delay)
        Case "contibuffer"
            Call doContiBuffer
            
            'No start stop
            'If (RightNeedle.Value = True) Then
                No_Start_Stop = True
            'End If
        Case "line3d"
            word2() = Split(ExecutionString, "=")
            word3() = Split(word2(1), ",")
            word4() = Split(word2(2), ",")
            word5() = Split(word2(3), ")")
            
            endx = CLng(Trim(word3(0)))
            endy = CLng(Trim(word4(0)))
            endz = CLng(Trim(word5(0)))
            
            'initialize
            Position(0) = 0
            Position(1) = 0
            Position(2) = 0
                
            tempEndX1 = origX + endx
            tempEndY1 = origY + endy
                
            tempEndX = tempEndX1
            tempEndY = tempEndY1
                
            origX = tempEndX1
            origY = tempEndY1
                
            detXYfromFudicial tempEndX, tempEndY, tempEndX, tempEndY, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
            
            endx = tempEndX - presentX
            endy = tempEndY - presentY
            
            'Origin (NYP)
            If (No_Start_Stop = True) Then
                Call doContiLine3D(endx, endy, endz)
            Else
                Call doContiLine3D_XW(endx, endy, endz)
            End If
                
            presentX = tempEndX
            presentY = tempEndY
        Case "segmentproperty3d"
            word2() = Split(ExecutionString, "=")
            word3() = Split(word2(1), ";")
            word4() = Split(word3(2), ")")
        
            Speed = CLng(Trim(word3(0)))
            dispenserControl = (Trim(word3(1)))
            seqNum = CLng(Trim(word4(0)))
            
            'Origin (NYP)
            If (No_Start_Stop = True) Then
                Call doSegmentProperty3D(Speed, dispenserControl, seqNum)
            Else
                Call doSegmentProperty3D_XW(Speed, dispenserControl, seqNum)
            End If
        Case "contiend"
            Call doContiEnd
            No_Start_Stop = False
        Case "fudicial"
            word2() = Split(ExecutionString, "=")
            word3() = Split(ExecutionString, ";")
            word4() = Split(word2(1), ",")
            word5() = Split(word2(2), ";")
            word6() = Split(word2(3), ",")
            word7() = Split(word2(4), ";")
            word8() = Split(word3(3), ")")
            
            deltaX = 0
            deltaY = 0
            deltaAngle = 0
            
            Call doFudicial(CLng(Trim(word4(0))), CLng(Trim(word5(0))), CLng(Trim(word6(0))), CLng(Trim(word7(0))), Trim(word3(2)), CLng(Trim(word8(0))))
            If Camera.Value = True Then
                SetLightIntensity (Val(LightingIntensity.Text))
            Else
                SetLightIntensity (0)
            End If
    End Select
End Sub

Private Sub doFudicial(x1 As Long, y1 As Long, x2 As Long, y2 As Long, FileName As String, LightingAmt As Long)
    
    Dim patFile As String
    Dim Length As Long
    
    'Move z_axis to "zero" before doing fiducial (XW)
    checkSuccess (P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(30, Z_axis), 2000000, 1200000, 9158400))
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
    Loop
    
    xOrgFid = convertToMM(x1, X_axis)
    yOrgFid = convertToMM(y1, Y_axis)
    
    Length = Len(FileName)

    patFile = Mid(FileName, 2, Length - 2)

    'Use "S-curve" (XW)
    'returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis), MaxVelocity, AccelSpeed, AccelRate)
    returncode = P1240MotAxisParaSet(boardNum, 0, X_axis Or Y_axis Or Z_axis, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis), MaxVelocity, AccelSpeed, AccelRate)

    visionRetry = True
    
    SetLightIntensity (LightingAmt)
    
    Do While (visionRetry)
        returncode = VdeFindRefPt(patFile, convertToMM(x1, X_axis), convertToMM(y1, Y_axis), convertToMM(x2, X_axis), convertToMM(y2, Y_axis), deltaX, deltaY, deltaAngle)
        If returncode <> 1 And OptionPause.Value = True Then
            RetryFudicial.Show (vbModal)
        End If
        If returncode <> 1 And OptionSkip.Value = True Then
            'FailFudicialSkipForm.Show (vbModal)
            '29 July 05
            PrintParseTree ("Failed Fuidicial!")
            abortValue = True
            PrintParseTree ("Dispensing Abort!")
            visionRetry = False
            doneFudicial = True
        End If
        If returncode = 1 Then
            visionRetry = False
            doneFudicial = True
        End If
    Loop
    
    VdeSelectCamera 2
    VdeCameraLive 1
 
End Sub

Private Function doContiStart(Speed As Long, dispense As Integer, axisMovement As Long)
    
    'returnCode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(Speed, Z_axis), MaxVelocity, AccelSpeed, AccelRate)
    
    'Use "S-curve" (XW)
    'returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)
    returncode = P1240MotAxisParaSet(boardNum, 0, &H7, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)
    
    '@$K
    If (dispense = 0 Or dispense = 10 Or dispense = 20 Or dispense = 30 Or dispense = 40) Then
        Call doDispenseOff
    ElseIf (dispense = 1 Or dispense = 11 Or dispense = 21 Or dispense = 31 Or dispense = 41) Then
        Call doDispenseOn
    End If
        
    If (axisMovement = 2) Then
        returncode = P1240StartContiDrive(boardNum, X_axis Or Y_axis, 0)
    ElseIf (axisMovement = 3) Then
        returncode = P1240StartContiDrive(boardNum, X_axis Or Y_axis Or Z_axis, 0)
    Else
        returncode = P1240StartContiDrive(boardNum, X_axis Or Y_axis Or U_axis, 0)
    End If
    
    Debug.Print "Did Contistart " & Speed
End Function

Private Function doContiEnd()
    
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
        DoEvents
    Loop
    returncode = P1240FreeContiBuf(boardNum)
    
    contiPathArrayIndex = 1
    
    Debug.Print "Did free contibuf " & returncode
    
End Function

Private Function doSegmentProperty(Speed As Long, dispenser As Integer, sequenceNum As Long)
    Dim PresentSequenceNum As Long
    
    If (sequenceNum = 1) Then
        Call doContiStart(Speed, dispenser, 2)
    Else
    
        Debug.Print "Waiting for sequence " & sequenceNum
    
        returncode = P1240GetCurContiNum(boardNum, PresentSequenceNum)
        
        Do While ((PresentSequenceNum <> sequenceNum) And (estopValue = False) And (abortValue = False))
            DoEvents
            returncode = P1240GetCurContiNum(boardNum, PresentSequenceNum)
        Loop
            
        Debug.Print "Doing speed change and dispenser control"
        
        If (dispenser = False Or estopValue = True Or abortValue = True) Then
            Call doDispenseOff
        Else
            Call doDispenseOn
        End If
        
        returncode = P1240MotChgLineArcDV(boardNum, convertSpeed(Speed, X_axis Or Y_axis))
    
        Debug.Print "Changespeed = " & returncode
    End If
    
End Function

Private Function doSegmentProperty3D(Speed As Long, dispenser As Integer, sequenceNum As Long)
    Dim PresentSequenceNum As Long
        
    '@$K
    If (sequenceNum = 1) Then
        If (dispenser = 0 Or dispenser = 1 Or dispenser = 10 Or dispenser = 11) Then
            Call doContiStart(Speed, dispenser, 3)
        Else
            If (Speed > 20) Then
                Speed = 20
            End If
            Call doContiStart(Speed, dispenser, 4)
        End If
        
        If (dispenser = 0 Or dispenser = 10 Or dispenser = 20 Or dispenser = 30 Or dispenser = 40) Then
            dispenser = 0
        ElseIf (dispenser = 1 Or dispenser = 11 Or dispenser = 21 Or dispenser = 31 Or dispenser = 41) Then
            dispenser = 1
        End If
    Else
        Debug.Print "Waiting for sequence " & sequenceNum
    
        returncode = P1240GetCurContiNum(boardNum, PresentSequenceNum)
        
        'Origin (NYP)
            'Do While ((PresentSequenceNum <> sequenceNum) And (estopValue = False) And (abortValue = False))
            Do While ((PresentSequenceNum < sequenceNum) And (estopValue = False) And (abortValue = False))
            'Not to do the infinite loop (XW)
            If (CloseBoard = True) Then
                Exit Function
            End If
            
            DoEvents
            returncode = P1240GetCurContiNum(boardNum, PresentSequenceNum)
        Loop
            
        Debug.Print "Doing speed change and dispenser control"
        
        '@$K
        If (dispenser = 0 Or dispenser = 10 Or dispenser = 20 Or dispenser = 40 Or estopValue = True Or abortValue = True) Then
                Call doDispenseOff
        Else
            If (dispenser = 1 Or dispenser = 11 Or dispenser = 21 Or dispenser = 31 Or dispenser = 41) Then
                Call doDispenseOn
            End If
        End If
        
'        If (dispenser = False Or estopValue = True Or abortValue = True) Then
'            Call doDispenseOff
'        Else
'            Call doDispenseOn
'        End If
        
        returncode = P1240MotChgLineArcDV(boardNum, convertSpeed(Speed, X_axis Or Y_axis))
    
        Debug.Print "Changespeed = " & returncode
    End If
    
End Function

Private Function doContiArc(endx As Long, endy As Long, CenX As Long, ceny As Long, ccw)
        
    If (ccw = 1) Then
        'Origin
        contiPathArray.PathType = IPO_CCW
        'contiPathArray.PathType = IPO_CW
    Else
        contiPathArray.PathType = IPO_CW
        'contiPathArray.PathType = IPO_CCW
    End If
    contiPathArray.EndPoint_1 = endx
    contiPathArray.EndPoint_2 = endy
    contiPathArray.EndPoint_3 = 0
    contiPathArray.CenPoint_1 = CenX
    contiPathArray.CenPoint_2 = ceny
    returncode = P1240SetContiData(boardNum, contiPathArray, contiPathArrayIndex)
    contiPathArrayIndex = contiPathArrayIndex + 1

End Function

Private Function doContiLine(endx As Long, endy As Long)
        
    contiPathArray.PathType = IPO_L2
    contiPathArray.EndPoint_1 = endx
    contiPathArray.EndPoint_2 = endy
    contiPathArray.EndPoint_3 = 0
    returncode = P1240SetContiData(boardNum, contiPathArray, contiPathArrayIndex)
    contiPathArrayIndex = contiPathArrayIndex + 1
    
    Debug.Print "Did contiline"

End Function

Private Function doContiLine3D(endx As Long, endy As Long, endz As Long)

    contiPathArray.PathType = IPO_L3
    contiPathArray.EndPoint_1 = endx
    contiPathArray.EndPoint_2 = endy
    contiPathArray.EndPoint_3 = endz
    returncode = P1240SetContiData(boardNum, contiPathArray, contiPathArrayIndex)
    contiPathArrayIndex = contiPathArrayIndex + 1
    
    Debug.Print "Did contiline3D"

End Function

Private Function doContiBuffer()
    'Original
    'returncode = P1240InitialContiBuf(0, 100)
    
    'Define as 300 lines for buffer (XW)
    returncode = P1240InitialContiBuf(0, 300)
    
    Debug.Print "Initialize contibuffer " & returncode
    
    contiPathArrayIndex = 1
End Function

Private Function doDispenseOn()
    Dim Value As Long
    
    If wetRun.Value = True Then
        If (prevDispenserValue = False) Then
    
            If (externalDispenserControl = True) Then
            
                Dim returnValue As Long
            
                returnValue = &H400
        
                Do While returnValue = &H400
                    returncode = P1240MotRdReg(boardNum, 8, RR5, returnValue)
                    returnValue = returnValue & &H400
                    DoEvents
                Loop
        
            End If

            Debug.Print "Dispense On"
            'Origin (NYP)
            'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H100)
            
            If (L_Dispense = True) Then
            'If (R_Dispense = True) Then
                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, Value)
                Value = (Value Or &H800)
                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, Value)
            ElseIf (R_Dispense = True) Then
            'ElseIf (L_Dispense = True) Then
                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, Value)
                Value = (Value Or &H100)
                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, Value)
            End If
        End If
    
        prevDispenserValue = True
        
    End If
    
End Function

Private Function doDispenseOff()
    Dim Value As Long
    
    If (prevDispenserValue = True) Then
    
        Debug.Print "Dispense Off"
    
        'Origin (NYP)
        'returncode = P1240MotWrReg(0, 4, WR3, &H0)
        
        returncode = P1240MotRdReg(boardNum, Y_axis, WR3, Value)
        Value = (Value And &HF7FF)
        returncode = P1240MotWrReg(boardNum, Y_axis, WR3, Value)
        
        'returncode = P1240MotRdReg(boardNum, Z_axis, WR3, Value)
        'Value = (Value And &H600)
        'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, Value)
        
        returncode = P1240MotRdReg(boardNum, Z_axis, WR3, Value)
        Value = (Value And &HFEFF)
        returncode = P1240MotWrReg(boardNum, Z_axis, WR3, Value)
    End If
    
    prevDispenserValue = False
    
End Function

Private Function doPtp(ByVal x As Long, ByVal y As Long, ByVal Z As Long, ByVal Speed As Long)
    Debug.Print "DoPtp " & x & " " & y & " " & Z & " " & Speed
    
    Dim Value, valueX, valueY, valueZ, ValueU As Long
    Dim AccelSpeed, AccelSpeedZ As Double
    Dim AccelRate, AccelRateZ As Double
    Dim factor, factorZ As Double
    
    If ballScrew = 1 Then
        '0.12G
        AccelSpeedZ = 1200000
        AccelRateZ = 9158400
        '0.5G
        AccelSpeed = 5000000
        AccelRate = 30000000
        
        'AccelSpeedZ = 260000       'origin
        'AccelRateZ = 500000
        'AccelSpeed = 260000
        'AccelRate = 500000
     
        'If Speed <= 10 Then
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            'AccelRateZ = AccelSpeedZ / 0.05

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.1
            'AccelRate = AccelSpeed / 0.03
        'ElseIf Speed <= 90 Then
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            'AccelRateZ = AccelSpeedZ / 0.05

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.2
            'AccelRate = AccelSpeed / 0.06
        'Else
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.8
            'AccelRateZ = AccelSpeedZ / 0.08

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.7
            'AccelRate = AccelSpeed / 0.2

        'End If
    Else
    
        If Speed <= 10 Then
            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            AccelRateZ = AccelSpeedZ / 0.05
            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.01
            AccelRate = AccelSpeed / 0.05

        Else
            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.4
            AccelRateZ = AccelSpeedZ / 0.05
            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.05
            AccelRate = AccelSpeed / 0.05
        End If
    
    End If
    
    returncode = P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(Speed, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate)
    returncode = P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(Speed, Z_axis), 2000000, AccelSpeedZ, AccelRateZ)
    
    Do While (stopValue = True And estopValue = False And abortValue = False)
        DoEvents
        pauseButtonTimer.Enabled = True
    Loop
    
    If (x = 0 And y = 0 And Z = 0) Then
        returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, -1000, 0)
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
            DoEvents
        Loop
        returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, 1000, -1000, 0, 0)
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
            DoEvents
        Loop
    Else
        returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, Z, 0)
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
            DoEvents
        Loop
        returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, x, y, 0, 0)
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
            DoEvents
        Loop
    End If
    
    proceed = False
    checkProceed
            
    cmdStartButton.Enabled = False
            
    If (x = 0 And y = 0 And Z = 0) Then
        home_limit_flag = True  'xu long
        'Chagne from T_curve to S_curve
        'returncode = P1240MotHome(boardNum, X_axis Or Y_axis Or Z_axis)    'origin
        If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(20, Z_axis), 200000, 1200000, 9158400))) Then
            If (checkSuccess(P1240MotCmove(boardNum, Z_axis, 0))) Then  'Move Z in clockwise direction
                Value = 0
                
                Do While ((Value And &H4) <> &H4)  'Do loop if Z Limit switch still not reached
                    checkSuccess (P1240MotRdReg(boardNum, Z_axis, RR2, Value))
                    DoEvents
                Loop
                If ((Value And &H4) = &H4) Then 'Do immediate stop on Z axis
                    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
                End If
                
                Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)  'Loop while Z motor is still spinning
                Loop
                If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
                    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)  'Loop while Z motor is still spinning
                        DoEvents
                    Loop
                    
                    '''''''''''''''''''''
                    '   U_axis (Homing) '
                    '''''''''''''''''''''
                    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 300, 300, 8300, 53000, 9000000))
                    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 600, 600, 16600, 106000, 18000000))
                    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 200, 10000, 50000, 9000000))
    
                    checkSuccess (P1240MotCmove(boardNum, U_axis, 8))
                    ValueU = 0
                
                    Do While ((ValueU And &H8) <> &H8)
                        checkSuccess (P1240MotRdReg(boardNum, U_axis, RR2, ValueU))
            
                        If ((ValueU And &H8) = &H8) Then 'Do immediate stop on U axis
                            checkSuccess (P1240MotStop(boardNum, U_axis, 8))
                        End If
                        DoEvents
                        If (Emergency_Stop = True) And (Ext = True) Then
                            readyStatus = True
                            busyStatus = False
                            Emergency_Stop = False
                            Exit Function
                        End If
                    Loop
                    Do While (P1240MotAxisBusy(boardNum, U_axis) <> Success)
                        DoEvents
                    Loop
            
                    checkSuccess (P1240MotHome(boardNum, U_axis))
                    Do While (P1240MotAxisBusy(boardNum, U_axis) <> Success)
                        DoEvents
                        If (Emergency_Stop = True) And (Ext = True) Then
                            readyStatus = True
                            busyStatus = False
                            Emergency_Stop = False
                            Exit Function
                        End If
                    Loop
                    
                    If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(20, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
                        If (checkSuccess(P1240MotCmove(boardNum, X_axis, 1))) Or (checkSuccess(P1240MotCmove(boardNum, Y_axis, 0))) Then
                        'Move X and Y motors in clockwise direction and anti-clockwise direction
                            valueX = 0
                            valueY = 0
                            Do While (((valueX And &H8) <> &H8) Or ((valueY And &H4) <> &H4))
                                checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, valueX, valueY, valueZ, ValueU))
                                If ((valueY And &H4) = &H4) Then
                                    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
                                End If
                                If ((valueX And &H8) = &H8) Then
                                    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
                                End If
                                DoEvents
                            Loop
                            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                            Loop
                            If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
                                Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                                    DoEvents
                                Loop
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
        Loop
    End If
    
    'For U_axis
    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 500, 8300, 8300, 53000, 9000000))
    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 1000, 16600, 16600, 106000, 18000000))
    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 3000, 10000, 50000, 9000000))
    
    home_limit_flag = False
    
End Function
Private Function checkProceed()

    Do While proceed = False
        DoEvents
    Loop
    
End Function

Private Function doDelay(delay As Double)

    Debug.Print "DoDelay " & delay

Dim Start

Start = Timer

Do While ((Timer < Start + delay) And (estopValue = False) And (abortValue = False))
    If (Timer < Start) Then
        Start = (86400 - Start)
    End If
        
    DoEvents
Loop

End Function

Private Sub eStopTimer_Timer()
    eStopTimer.Enabled = False
    Dim buttonValue As Long
    Dim ValveClose As Long, ReadValue As Long, DriverXYZ As Long

    'returncode = P1240MotRdReg(boardNum, X_axis, RR0, buttonValue)     'origin
    'buttonValue = buttonValue And &HF0                                 'origin
    
    returncode = P1240MotRdReg(boardNum, X_axis, RR2, buttonValue)
    buttonValue = buttonValue And &H20
    
    'If (buttonValue = &HF0) Then       'origin
     If (buttonValue = &H20) Then
        abortButtonTimer.Enabled = False
        startButtonTimer.Enabled = False
        purgeButtonTimer.Enabled = False
        pauseButtonTimer.Enabled = False
        
        Servo_Off
         
        'Close valve first before the Emergency Form come out (XW)
        'Call P1240MotRdReg(boardNum, Z_axis, WR3, ValveClose)
        'If (ValveClose = &H304) Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
            ValveClose = (ValveClose And &HF7FF)
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
            
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &H0))
            
            'Right-needle will be gone up.
            checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HFEFF
            checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
    
            'Left-needle will be gone up.
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HF7FF
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        'End If
        
        estopValue = True
        Green_Light = False
        Red_Light = True
        returncode = P1240MotWrReg(boardNum, X_axis, WR3, &H801) 'Activate the buzzer
        EmergencyStopForm.Show (vbModal)
        'Close executinFrom firs to get the enough processing time to end the whole program
        'XW
        If (Close_Emg = True) Then
            Unload Me
        End If
    End If
    eStopTimer.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim ReadValue As Long, Tower_Light_Value As Long
        
    'Change them to the original states
    'XW
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, &H0))
    
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &H800
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    'Remove this one because it will make the left_needle moved up when we close the form. (XW)
    'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H0) 'xu long
    systemHomeY = systemHomeY * (-1)
    systemHomeZ = systemHomeZ * (-1)
        
    Leftslider_go_up
    
    Call setSpeed(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "100"))
    
    '@$K
    If GetStringSetting("EpoxyDispenser", "Setup", "EnableSolventPosition", "0") = "1" Then
        'Move to Solvent Position
        PTPToXYZ solventposX, solventposY, solventposZ
    Else
        'Move to System Home Position
        PTPToXYZ systemHomeX, systemHomeY, systemHomeZ
    End If
    
    'Left needles go down
    Leftslider_go_down
            
    If (Close_Emg = True) Then
        displayCoOrdsTimer.Enabled = False
          
        'Right-needle will be gone up.
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
        ReadValue = ReadValue And &HFEFF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
    
        'Left-needle will be gone up.
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
        ReadValue = ReadValue And &HF7FF
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
         
        Call Sleep(0.5)
         
        'Close "tilting valve"
        checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
        ReadValue = ReadValue And &HFEFF
        checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
           
        'Disable Red_Light,Yellow_light and Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Tower_Light_Value = Tower_Light_Value And &HF1FF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
        
        PicImage.Enabled = False 'NNO
        returncode = P1240FreeContiBuf(0)
        If (Start_Form = True) Then
            Call Sleep(0.3)
            VdeCameraLive False
            VdeReleaseVision
        End If
        unInitializeBoard
    Else
        PicImage.Enabled = False 'NNO
        Call Sleep(0.3)
        VdeCameraLive False
        VdeReleaseVision
        
        'Remove this part because of the customer's spect
        'Right-needle will be gone up.
        'Call P1240MotRdReg(boardNum, U_axis, WR3, ReadValue)
        'ReadValue = ReadValue And &HFEFF
        'Call P1240MotWrReg(boardNum, U_axis, WR3, ReadValue)
    
        'Left-needle will be gone up.
        'Call P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
        'ReadValue = ReadValue And &HF7FF
        'Call P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
         
        'Sleep (0.5)
        
        'Disable Red_Light,Yellow_light and Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Tower_Light_Value = Tower_Light_Value And &HF1FF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
        
        Red_Light = False
        Yellow_Light = False
        Green_Light = False
        
        Call Sleep(0.2)
        
        Call Turn_Off_LightIntensity
        If (editorForm.mscomLighIntensity.PortOpen = True) Then
            editorForm.mscomLighIntensity.PortOpen = False
        End If
        
    End If
    
    SaveStringSetting "EpoxyDispenser", "Setup", "DispensingTime", txtPurgeTime.Text
    StopTimer
    EstopLimit
    Close_PCI1750
    CloseBoard = True
End Sub

Private Sub homeButton_Click()
    readyStatus = False         'For indicator (XW)
    busyStatus = True
    
    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    PurgePosition.Enabled = False
    FindNeedleOffset.Enabled = False

    Leftslider_go_up
    'doHome
    'To provide a shortcut for machine home 3 May 2005
    doPtp 0, 0, 0, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50"))
    doSoftHome
    Leftslider_go_down
    
    cmdStartButton.Enabled = True
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    PurgePosition.Enabled = True
    'FindNeedleOffset.Enabled = True        'done by XW
    closeButton.Enabled = True
    
    busyStatus = False
    readyStatus = True
End Sub

Private Sub OnlineTimer_Timer()
    Dim boardActive As Long
    Dim Solvent As Boolean
    'XW
    If (CloseBoard = False) Then
    
        OnlineTimer.Enabled = False
        purgeButtonTimer.Enabled = False
        'XW
        'StopTimer
        
        'If (onlineOption.Enabled = True) And (onlineOption.Value = 1) Then
        '    onlineOption.Enabled = False
        'End If
        
        Start_Time = Timer
        
        boardActive = &H200
        Do While (boardActive = &H200)
            'XW
            If (CloseBoard = True) Then
                Exit Sub
            ElseIf (ReStart = True) Then
                ReStart = False
                Exit Sub
            Else
                checkSuccess (P1240MotRdReg(boardNum, Y_axis, RR4, boardActive))
                boardActive = boardActive And &H200
                DoEvents
            End If
            
            '@$K
            If (Timer - Start_Time > 20) And Solvent = False Then
                Call setSpeed(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "100"))
                PTPToXYZ solventposX, solventposY, solventposZ
                Solvent = True
            End If
        Loop
        
        Start_Time = Timer
        If (AlwaysPurge.Value = 1) Or Solvent = True Then
            Always_Purge
        End If
        
        doDispense
    
        boardActive = &H0
    
        Do While (boardActive = &H0)
            'XW
            If (CloseBoard = True) Then
                Exit Sub
            ElseIf (ReStart = True) Then
                ReStart = False
                Exit Sub
            Else
                checkSuccess (P1240MotRdReg(boardNum, Y_axis, RR4, boardActive))
                boardActive = boardActive And &H200
                DoEvents
            End If
        Loop
        
        'Give Production mode signal
        'Call P1240MotWrReg(boardNum, Z_axis, WR3, &H200)    'XL
        
        'Give Production mode & Left valve signals
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &HA00))
        
        'XW
        'OnTimer
        
        If abortValue = False Then
            'XW (enable the check box)
            'onlineOption.Enabled = True
            OnlineTimer.Enabled = True
        End If
        purgeButtonTimer.Enabled = True
    End If
End Sub

Private Sub PurgePosition_Click()
    readyStatus = False     'For indicator (XW)
    busyStatus = True
    
    Dim xpos, ypos, zpos As Long
        
    xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
    ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
    'To get the actual direction        'XW
    ypos = ypos * (-1)
    zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
    zpos = zpos * (-1)
    
    'XW
    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    systemMoveHeight = systemMoveHeight * (-1)
    
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False

    FindNeedleOffset.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    abortButton.Enabled = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdStartButton.Enabled = False
    FindNeedleOffset.Enabled = False
    
    Leftslider_go_up
    
    'To travel the right way (XW)
    If (systemMoveHeight > zpos) Then
        doPtp xpos, ypos, systemMoveHeight, zSystemTravelSpeed
    Else
        ClickPurgePosition = True
    End If
    
    doPtp xpos, ypos, zpos, xySystemTravelSpeed
    
    startButtonTimer.Enabled = True
    If (DisablePurgeSignal = False) And (Button_Purge = False) Then
        purgeButtonTimer.Enabled = True
    End If
        
    'FindNeedleOffset.Enabled = True        'done by XW
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    'abortButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'cmdStopButton.Enabled = True
    cmdStartButton.Enabled = True
    'FindNeedleOffset.Enabled = True        'done by XW
    
    busyStatus = False
    readyStatus = True
    Leftslider_go_down
End Sub

Private Sub redrawValveOnOff_Timer()
    Dim PW, PH
    Dim ValveStatus As Long, ValveStatus2 As Long
   
    'XW
    If (CloseBoard = False) Then
        PictureValveOnOff.FillStyle = vbFSSolid
    
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ValveStatus))
        ValveStatus = ValveStatus And &H100
        
        checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveStatus2))
        ValveStatus2 = ValveStatus2 And &H800
    
        If (ValveStatus = &H100) Or (ValveStatus2 = &H800) Then
            PictureValveOnOff.FillColor = QBColor(10)
        Else
            PictureValveOnOff.FillColor = QBColor(2)
        End If
        
        PW = PictureValveOnOff.ScaleWidth
        PH = PictureValveOnOff.ScaleHeight
        
        ' Draw circle
        PictureValveOnOff.Circle (PW / 2, PH / 2), PH / 3
    End If
End Sub

Private Sub abortButtonTimer_Timer()
    
    Dim buttonValue As Long
    Dim skip As Boolean
    
    'XW
    If (CloseBoard = False) Then
        abortButtonTimer.Enabled = False
        
        'Check Door Lock Sensor
        'If (Door_Lock = True) Then
        '    abortButton_Click
        '    MsgBox "Please lock the door first before running the application!"
        '    Exit Sub
        'End If
    
        returncode = P1240MotRdReg(boardNum, Z_axis, RR5, buttonValue)
    
        buttonValue = buttonValue And &H4
    
        skip = False
    
        If (buttonValue = &H0) Then
    
            returncode = P1240MotStop(boardNum, X_axis Or Y_axis Or Z_axis, X_axis Or Y_axis Or Z_axis)
            'XW
            'After pressing external "Stop" button, we need to check whether all axes are still busy or not.
            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
            Loop
            abortValue = True
            
            'Make the same as GUI
            If cmdResumeButton.Enabled = True Then
                stopValue = False
                cmdResumeButton.Enabled = False
                cmdResumeButton.Refresh
                pauseButtonTimer.Enabled = True     'XW
            End If
            
            PrintParseTree ("Dispensing Abort!")
            abortButton.Enabled = False
            skip = True
            
            prevDispenserValue = False
        End If

        If skip = False Then
            abortButtonTimer.Enabled = True
        End If
    End If
End Sub

Private Sub Stop_MotAndValve()
    Dim ValveClose As Long, ReadValue As Long
    
    'Close both valves
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ValveClose))
    ValveClose = (ValveClose And &HFEFF)
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ValveClose))
    
    Call Sleep(0.03)
    
    'Close both valves
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
    ValveClose = (ValveClose And &HF7FF)
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
    
    Call Sleep(0.03)
    
    'Right-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HFEFF
    checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
     
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
            
    Call Sleep(0.03)
            
    checkSuccess (P1240MotStop(boardNum, XYZU_axis, XYZU_axis))
    'XW
    'After pressing external "Stop" button, we need to check whether all axes are still busy or not.
    Do While (P1240MotAxisBusy(boardNum, XYZU_axis) <> Success)
    Loop
End Sub

Private Sub resumeTimer_Timer()

    Dim buttonValue, pauseButtonValue As Long
    Dim skip As Boolean

    'XW
    If (CloseBoard = False) Then
        resumeTimer.Enabled = False
    
        returncode = P1240MotRdReg(boardNum, X_axis, RR4, pauseButtonValue)
    
        pauseButtonValue = pauseButtonValue And &H4
    
    
        returncode = P1240MotRdReg(boardNum, Z_axis, RR5, buttonValue)
    
        buttonValue = buttonValue And &H2
        skip = False
    
        If (buttonValue = &H0) And (pauseButtonValue = &H4) Then
            PrintParseTree ("Dispensing Resume...")
            stopValue = False
            'To match GUI
            cmdResumeButton.Enabled = False
            cmdStopButton.Enabled = True
            pauseButtonTimer.Enabled = True
            skip = True
        End If
        If skip = False Then
            resumeTimer.Enabled = True
        End If
    End If
End Sub

Private Sub startButtonTimer_Timer()

    Dim buttonValue1, buttonValue2, abortButtonValue As Long
    Dim skip As Boolean
    
    'Check Door Lock Sensor
    If (Inter_Lock = True) Then
        If (Door_Lock = True) Then
            MsgBox "Please lock the door first before running the application!"
            Exit Sub
        End If
    End If
    
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    
    'Move to "doNormalDispense" procedure
    'abortButtonTimer.Enabled = True
    
    If (CloseBoard = False) Then
        returncode = P1240MotRdReg(boardNum, X_axis, RR4, abortButtonValue)
    
        abortButtonValue = abortButtonValue And &H4
    
        returncode = P1240MotRdReg(boardNum, X_axis, RR4, buttonValue1)
        buttonValue1 = buttonValue1 And &H2
        returncode = P1240MotRdReg(boardNum, Z_axis, RR5, buttonValue2)
        buttonValue2 = buttonValue2 And &H2
    
        skip = False
    
        If (buttonValue1 = &H0) And (buttonValue2 = &H0) And (abortButtonValue = &H4) Then
            stopValue = False
            pauseButtonTimer.Enabled = True
            skip = True

            'abortValue = False
            'Call doNormalDispense
        
            If onlineOption.Value = 0 Then
                Start_Time = Timer
                ReStart = True
                'Do the purging first
                If (AlwaysPurge.Value = 1) Then
                    Always_Purge
                End If
                doDispense
            Else
                OnlineTimer.Enabled = True
            End If
        
            'If (estopValue = False) Then
                'doHome
                'doPtp 0, 0, 0, 50
                'doSoftHome 'See comments of 300605
            'End If
        End If
        
        If skip = False Then
            startButtonTimer.Enabled = True
        End If
    End If
    'Commented to prevent flickering
    cmdStartButton.Enabled = True
    closeButton.Enabled = True
    If (DisablePurgeSignal = False) And (Button_Purge = False) Then
        purgeButtonTimer.Enabled = True
    End If
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True

End Sub

Private Sub purgeButtonTimer_Timer()
    Dim buttonValue As Long
    
    'XW
    If (CloseBoard = False) Then
        purgeButtonTimer.Enabled = False
        
        returncode = P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue)
        buttonValue = buttonValue And &H400
    
        If (buttonValue = 0) Then
            'PrintParseTree ("Purge...")
            
            cmdPurgeButton.Enabled = False
            Disable_Button
            
            Dim xpos, ypos, zpos, ReadValue As Long
        
            xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
            ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
            'To get the actual direction        'XW
            ypos = ypos * (-1)
            zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
            zpos = zpos * (-1)
    
            'XW
            systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
            systemMoveHeight = systemMoveHeight * (-1)
            
            DisablePurgeSignal = True
            checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
            
            'Left-needle will be gone up.
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HF7FF
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
            Call Sleep(0.2)
    
            'To travel the right way (XW)
            If (systemMoveHeight > zpos) Then
                doPtp xpos, ypos, systemMoveHeight, zSystemTravelSpeed
            Else
                ClickPurgePosition = True
            End If
    
            doPtp xpos, ypos, zpos, xySystemTravelSpeed
        
            'Left slider go down
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H800
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
            Call Sleep(0.2)
    
            Do While (buttonValue = 0)
                If leftside = False And rightside = True Then
                    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, buttonValue)
                    buttonValue = (buttonValue Or &H100)
                    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, buttonValue)
                ElseIf leftside = True And rightside = False Then
                    returncode = P1240MotRdReg(boardNum, Y_axis, WR3, buttonValue)
                    buttonValue = (buttonValue Or &H800)
                    returncode = P1240MotWrReg(boardNum, Y_axis, WR3, buttonValue)
                Else
                    'If (executionForm.LeftNeedle.Value = True) Then
                    returncode = P1240MotRdReg(boardNum, Y_axis, WR3, buttonValue)
                    buttonValue = (buttonValue Or &H800)
                    returncode = P1240MotWrReg(boardNum, Y_axis, WR3, buttonValue)
                    'ElseIf (executionForm.RightNeedle.Value = True) Then
                    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, buttonValue)
                    buttonValue = (buttonValue Or &H100)
                    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, buttonValue)
                    'End If
                End If
                returncode = P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue)
                buttonValue = buttonValue And &H400
            Loop
     
            
            If (DisablePurgeSignal = True) Then
                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, buttonValue)
                buttonValue = buttonValue And &HF7FF
                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, buttonValue)
            
                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, buttonValue)
                buttonValue = buttonValue And &HFEFF
                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, buttonValue)
            
                readyStatus = False
                busyStatus = True
                Call Sleep(0.2)
                'Left-needle will be gone up.
                checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
                ReadValue = ReadValue And &HF7FF
                checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
                Call Sleep(0.2)
    
                doPtp PrePositionX, PrePositionY, systemMoveHeight, zSystemTravelSpeed
                doPtp PrePositionX, PrePositionY, PrePositionZ, xySystemTravelSpeed
                
                'Left slider go down
                checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
                ReadValue = ReadValue Or &H800
                checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
                
    
                DisablePurgeSignal = False
                readyStatus = True
                busyStatus = False
            End If
            
            cmdPurgeButton.Enabled = True
            Enable_Button
        End If
    
        purgeButtonTimer.Enabled = True
    End If
End Sub

'''''''''''''''''''''
'   Wrong Testing   '
'''''''''''''''''''''

'Private Sub purgeButtonTimer_Timer()
'    Dim buttonValue As Long
'    'Dim skip As Boolean
    
'    'XW
'    If (CloseBoard = False) Then
'        purgeButtonTimer.Enabled = False
'        startButtonTimer.Enabled = False
'        pauseButtonTimer.Enabled = False
'        abortButtonTimer.Enabled = False
'        resumeTimer.Enabled = False
        
'        returncode = P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue)
'        buttonValue = buttonValue And &H400
    
'        If (buttonValue = 0) Then
'            'PrintParseTree ("Purge...")
            
'            cmdPurgeButton.Enabled = False
            
'            SignalCounter = SignalCounter + 1
'            PurgeTimer
            
'            If (LeftNeedle.Value = True) Then
'                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, buttonValue)
'                buttonValue = (buttonValue Or &H800)
'                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, buttonValue)
'            ElseIf (RightNeedle.Value = True) Then
'                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, buttonValue)
'                buttonValue = (buttonValue Or &H100)
'                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, buttonValue)
'            End If
            
'        '    skip = True
'        Else
'            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, buttonValue)
'            buttonValue = buttonValue And &HF7FF
'            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, buttonValue)
            
'            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, buttonValue)
'            buttonValue = buttonValue And &HFEFF
'            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, buttonValue)
            
'            SignalCounter = 0
'            PurgeTimer
'            cmdPurgeButton.Enabled = True
'        End If
    
'        'If skip = False Then
'            purgeButtonTimer.Enabled = True
'            startButtonTimer.Enabled = True
'            pauseButtonTimer.Enabled = True
'            abortButtonTimer.Enabled = True
'            resumeTimer.Enabled = True
'        'End If
'    End If
'End Sub

'Private Sub PurgeTimer()
'    Dim xpos, ypos, zpos As Long
        
'    xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
'    ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
'    'To get the actual direction        'XW
'    ypos = ypos * (-1)
'    zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
'    zpos = zpos * (-1)
    
'    'XW
'    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
'    systemMoveHeight = systemMoveHeight * (-1)
    
'    If (SignalCounter = 1) Then
'        checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
        
'        'To travel the right way (XW)
'        If (systemMoveHeight > zpos) Then
'            doPtp xpos, ypos, systemMoveHeight, zSystemTravelSpeed
'        Else
'            ClickPurgePosition = True
'        End If
    
'        doPtp xpos, ypos, zpos, xySystemTravelSpeed
        
'        DisablePurgeSignal = True
'    ElseIf (SignalCounter = 0) Then
'        If (DisablePurgeSignal = True) Then
'            readyStatus = False
'            busyStatus = True
'            doPtp PrePositionX, PrePositionY, systemMoveHeight, zSystemTravelSpeed
'            doPtp PrePositionX, PrePositionY, PrePositionZ, xySystemTravelSpeed
'            DisablePurgeSignal = False
'            readyStatus = True
'            busyStatus = False
'        End If
'    End If
'End Sub

Private Sub pauseButtonTimer_Timer()

    Dim buttonValue, startbuttonValue1, startbuttonValue2 As Long
    Dim skip As Boolean
    
    'XW
    If (CloseBoard = False) Then
        pauseButtonTimer.Enabled = False
        
        If (Inter_Lock = True) Then
            'Check Door Lock Sensor
            If (Door_Lock = True) Then
                Dim ValveClose As Long
                
                checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
                ValveClose = (ValveClose And &HF7FF)
                checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
            
                checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &H200))
            
                Call Sleep(0.02)
                MsgBox "Please lock the door first when running the application!"
                
                Dim dor As Boolean
                dor = True
                Do While (dor)
                    If (Inter_Lock = True) Then
                        If (Door_Lock = True) Then
                            MsgBox "Please lock the door first before running the application!"
                        Else
                            dor = False
                            Exit Do
                        End If
                    Else
                        dor = False
                        Exit Do
                    End If
                Loop
                
                cmdStopButton_Click
                Exit Sub
            End If
        End If
    
        returncode = P1240MotRdReg(boardNum, X_axis, RR4, startbuttonValue1)
        startbuttonValue1 = startbuttonValue1 And &H2
        returncode = P1240MotRdReg(boardNum, Z_axis, RR5, startbuttonValue2)
        startbuttonValue2 = startbuttonValue2 And &H2
    
        returncode = P1240MotRdReg(boardNum, X_axis, RR4, buttonValue)
    
        buttonValue = buttonValue And &H4
        skip = False
    
        If (buttonValue = 0) And (startbuttonValue1 <> 0) And (startbuttonValue2 <> 0) Then
            PrintParseTree ("Dispensing Pause!")
            stopValue = True
            cmdStopButton.Enabled = False       'Should not be pressed more than one time (XW)
            cmdResumeButton.Enabled = True
            resumeTimer.Enabled = True
            skip = True
        End If
    
        If skip = False Then
            pauseButtonTimer.Enabled = True
        End If
        
    End If
End Sub

Private Sub Timer1_Timer()
    Dim xlimit, ylimit, zlimit, ulimit As Long
    
    'XW
    If (CloseBoard = False) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) = Success) And (proceed = False) Then
            proceed = True
        End If
    
        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, RR2, xlimit, ylimit, zlimit, ulimit))
        xlimit = xlimit And &HC
        ylimit = ylimit And &HC
        zlimit = zlimit And &HC
        ulimit = ulimit And &H8
    
        'If (((xlimit <> 0) Or (ylimit <> 0) Or (zlimit <> 0)) And printErrorLimit = True) Then     'origin
        If (((xlimit <> 0) Or (ylimit <> 0) Or (zlimit <> 0) Or (ulimit <> 0)) And printErrorLimit = True) And home_limit_flag = False Then
            PrintParseTree ("Error limit reach!")
            printErrorLimit = False
            Call abortButton_Click
        End If
        
        If (((xlimit = 0) And (ylimit = 0) And (zlimit = 0) And (ulimit = 0)) And printErrorLimit = False) Then
            printErrorLimit = True
        End If
    End If
End Sub

Private Sub doSoftHome()

    If (systemHomeX <> 0 Or systemHomeY <> 0 Or systemHomeZ <> 0) Then

        Dim rememberToTurnOffTimer1 As Boolean
    
        rememberToTurnOffTimer1 = False

        If (Timer1.Enabled = False) Then
            Timer1.Enabled = True
            rememberToTurnOffTimer1 = True
    
        End If
        
        systemTrackMoveHeight = systemMoveHeight

        If GetStringSetting("EpoxyDispenser", "Setup", "DirectSoftHome", "0") = "1" Then
    
            'returncode = P1240InitialContiBuf(0, 100)
        
            'contiPathArray.PathType = IPO_L3
            'contiPathArray.EndPoint_1 = systemHomeX - convertToPulses(editorForm.xCoOrd.Text, X_axis) - needleOffsetX
            'contiPathArray.EndPoint_2 = systemHomeY - convertToPulses(editorForm.yCoOrd.Text, Y_axis) - needleOffsetY
            'contiPathArray.EndPoint_3 = systemHomeZ - convertToPulses(editorForm.zCoOrd.Text, Z_axis)
            'returncode = P1240SetContiData(0, contiPathArray, 1)
        
            'contiPathArray.EndPoint_1 = 0
            'contiPathArray.EndPoint_2 = 0
            'contiPathArray.EndPoint_3 = 0
        
            'returncode = P1240SetContiData(0, contiPathArray, 2)
            returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate)
            
            'returncode = P1240StartContiDrive(boardNum, X_axis Or Y_axis Or Z_axis, 0)
            returncode = P1240MotLine(0, X_axis Or Y_axis Or Z_axis, 1, systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, 0)
    
            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
                DoEvents
            Loop
        
            'returncode = P1240FreeContiBuf(0)
        Else
            'Call doPtp(0, 0, systemMoveHeight, 50)
            If systemMoveHeight > systemHomeZ Then
                Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            End If
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
        End If

        If (rememberToTurnOffTimer1 = True) Then
            Timer1.Enabled = False
        End If

    End If

End Sub

Private Sub doHome()
    
    Timer1.Enabled = False
    
    cmdStartButton.Enabled = False
    
    Dim Value, valueX, valueY, valueZ, ValueU As Long
    
    Dim speed123 As Long
     
    If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, 0, 801, convertSpeed(30, Z_axis), 2000000, 1200000, 5000000))) Then
        If (checkSuccess(P1240MotCmove(boardNum, Z_axis, 0))) Then 'Move Z in clockwise direction
            Value = 0
            Do While ((Value And &H4) <> &H4) 'Do loop if Z Limit switch still not reached
                checkSuccess (P1240MotRdReg(boardNum, Z_axis, RR2, Value))
                DoEvents
            Loop
            If ((Value And &H4) = &H4) Then 'Do immediate stop on Z axis
                checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            End If
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
            Loop
            If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
                Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
                Loop
                If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, 0, StartVelocity, convertSpeed(30, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
                    If (checkSuccess(P1240MotCmove(boardNum, X_axis, 1))) Or (checkSuccess(P1240MotCmove(boardNum, Y_axis, 0))) Then
                    'Move X and Y motors in clockwise direction and anti-clockwise direction
                        valueX = 0
                        valueY = 0
                        Do While (((valueX And &H8) <> &H8) Or ((valueY And &H4) <> &H4))
                            checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, valueX, valueY, valueZ, ValueU))
                            If ((valueY And &H4) = &H4) Then
                                checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
                            End If
                            If ((valueX And &H8) = &H8) Then
                                checkSuccess (P1240MotStop(boardNum, X_axis, 1))
                            End If
                            DoEvents
                        Loop
                        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                        Loop
                        If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
                            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                            Loop
                        End If
                    End If
                End If
            End If
        End If
    End If
     
    Call doSoftHome
    
    cmdStartButton.Enabled = True
    
End Sub

Private Sub TimerDrawStatus_Timer()
    'XW
    If (CloseBoard = False) Then
        drawStatus
    End If
End Sub

Private Sub txtPurgeTime_Validate(cancel As Boolean)
    Call validateNumber(executionForm.txtPurgeTime.Text, executionForm.PurgeTime.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        executionForm.txtPurgeTime.Text = ""
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(executionForm.txtPurgeTime.Text) <= 0) Then
            executionForm.txtPurgeTime.Text = "1.00"
        ElseIf (CDbl(executionForm.txtPurgeTime.Text) >= 100) Then
            executionForm.txtPurgeTime.Text = "99.99"
        Else
            executionForm.txtPurgeTime.Text = Format(executionForm.txtPurgeTime.Text, "##0.00")
        End If
    End If
End Sub

Private Sub UpDownLighting_DownClick()

    If LightingIntensity.Text <> 0 Then
        'LightingIntensity.Text = LightingIntensity.Text - 1
    End If
End Sub

Private Sub UpDownlighting_UpClick()
    If LightingIntensity.Text <> 150 Then
        'LightingIntensity.Text = LightingIntensity.Text + 1
    End If
End Sub
Private Sub resetTimer_Timer()
    Dim resetX, resetY, resetZ As Long
    
    'XW
    If (CloseBoard = False) Then
        resetTimer.Enabled = False
        Call P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, resetX, resetY, resetZ, 0)
        resetX = resetX And &H10
        resetY = resetY And &H10
        resetZ = resetZ And &H10
        If (resetX = &H10) Or (resetY = &H10) Or (resetZ = &H10) Then
            'frmReset.Show (vbModal)
            
            MsgBox "Driver Error. Please check the Driver!"
            
            PicImage.Enabled = False
            
            Servo_Off
         
            Stop_MotAndValve
            
            StopTimer
            EstopLimit
            
            Close_TowerLight
    
            returncode = P1240FreeContiBuf(0)
            Call Sleep(0.03)
        
            VdeCameraLive False
            VdeReleaseVision
            Call Sleep(0.8)
            
            ResetDriver
            
            unInitializeBoard
            Close_PCI1750
            End
        
            Exit Sub
        End If
        resetTimer.Enabled = True
    End If
End Sub

'''''''''''''''''''''
'   Level Sensing   '
'''''''''''''''''''''

'Private Sub SensorTimer_Timer()
'    Dim sensorValue As Long
'    Dim a As Long
   
    'XW
'    If (CloseBoard = False) Then
'        SensorTimer.Enabled = False
'        Call P1240MotRdReg(boardNum, U_axis, RR5, sensorValue)
'        sensorValue = sensorValue And &H200
'        Call P1240MotRdReg(boardNum, X_axis, WR3, a)
     
        'If(sensorValue = &H200) Then
'        If (sensorValue = 0) Then       'Have to change
'            If a < &H800 Then
'                a = a Or &H800
'                Call P1240MotWrReg(boardNum, X_axis, WR3, a)
'            End If
'            frmAlarm.Show (vbModal)
'        Else
'            Call P1240MotWrReg(boardNum, X_axis, WR3, a)
'        End If
'        SensorTimer.Enabled = True
'    End If
'End Sub

Private Sub StopTimer()
    displayCoOrdsTimer.Enabled = False
    TimerDrawStatus.Enabled = False
    startButtonTimer.Enabled = False
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    resumeTimer.Enabled = False
    redrawValveOnOff.Enabled = False
    OnlineTimer.Enabled = False
    'SensorTimer.Enabled = False
    resetTimer.Enabled = False
End Sub

Private Sub OnTimer()
    displayCoOrdsTimer.Enabled = True
    TimerDrawStatus.Enabled = True
    startButtonTimer.Enabled = True
    pauseButtonTimer.Enabled = True
    abortButtonTimer.Enabled = True
    purgeButtonTimer.Enabled = True
    resumeTimer.Enabled = True
    redrawValveOnOff.Enabled = True
    OnlineTimer.Enabled = True
    'SensorTimer.Enabled = True
    resetTimer.Enabled = True
End Sub

Public Sub EstopLimit()
    Timer1.Enabled = False
    eStopTimer.Enabled = False
End Sub

Private Sub Disable_Button()
    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    closeButton.Enabled = False
    cmdCycleStop.Enabled = False
    PurgePosition.Enabled = False
    homeButton.Enabled = False
    NeedleMode.Enabled = False
    startButtonTimer.Enabled = False
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    resumeTimer.Enabled = False
    'OnlineTimer.Enabled = False
    resetTimer.Enabled = False
End Sub

Private Sub Enable_Button()
    cmdStartButton.Enabled = True
    cmdStopButton.Enabled = True
    cmdResumeButton.Enabled = True
    abortButton.Enabled = True
    closeButton.Enabled = True
    cmdCycleStop.Enabled = True
    PurgePosition.Enabled = True
    homeButton.Enabled = True
    NeedleMode.Enabled = True
    startButtonTimer.Enabled = True
    'pauseButtonTimer.Enabled = True
    'abortButtonTimer.Enabled = True
    'resumeTimer.Enabled = True
    
    'This one is isolated with purging.
    'OnlineTimer.Enabled = True
    resetTimer.Enabled = True
End Sub

Private Function Door_Lock() As Boolean
    Dim Door_Lock_Value As Byte
    
    'Check whether the door is opened or not (ILS)
    'Call P1240MotRdReg(boardNum, U_axis, RR5, Door_Lock_Value)
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, Door_Lock_Value)
    Door_Lock_Value = Door_Lock_Value And &H8
    
    '"&H8" means the door will not be locked.
    If (Door_Lock_Value = &H8) Then
    'If (Door_Lock_Value = 0) Then
        Door_Lock = True
    Else
        Door_Lock = False
    End If
End Function

Private Function Inter_Lock() As Boolean
    Dim Inter_Value As Byte
 
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, Inter_Value)
    Inter_Value = Inter_Value And &H20
    
    '"&H20/&H0" means inter_lock is off/on
    'If (Inter_Value = &H20) Then
    If (Inter_Value = &H0) Then
        Inter_Lock = True
    Else
        Inter_Lock = False
    End If
End Function

Private Function Low_Level() As Boolean
    Dim LowLevel_Value As Byte
 
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, LowLevel_Value)
    LowLevel_Value = LowLevel_Value And &H40
    
    '"&H10" doesn't mean low material
    'If (LowLevel_Value = &H10) Then
    If (LowLevel_Value = 0) Then
       
        MsgBox "The metarial is too low, please refill it."
       
        Yellow_Light = True
        Low_Level = True
    Else
        Yellow_Light = False
        Low_Level = False
    End If
End Function

Private Sub Always_Purge()
    Dim ReadValue As Long
    Dim Start
    Dim xpos, ypos, zpos As Long

    purgeButtonTimer.Enabled = False
    
    If (purgePartArray = False) Then
        Disable_Button
    End If
    
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
       
    xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
    ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
    'To get the actual direction        'XW
    ypos = ypos * (-1)
    zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
    zpos = zpos * (-1)
    
    'XW
    systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    systemMoveHeight = systemMoveHeight * (-1)
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.2)
    
    'To travel the right way (XW)
    If (systemMoveHeight > zpos) Then
        doPtp xpos, ypos, systemMoveHeight, zSystemTravelSpeed
    Else
        ClickPurgePosition = True
    End If
    
    doPtp xpos, ypos, zpos, xySystemTravelSpeed
        
    'Left slider go down
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.2)
    
    Start = Timer
    Do While (Timer < Start + CDbl(txtPurgeTime.Text))
        If leftside = False And rightside = True Then
            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H100
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
        ElseIf leftside = True And rightside = False Then
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H800
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
        Else
            'If (executionForm.LeftNeedle.Value = True) Then
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H800
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
            'ElseIf (executionForm.RightNeedle.Value = True) Then
            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H100
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
            'End If
        End If
        DoEvents
    Loop
    
    returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HF7FF
    returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
    
    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HFEFF
    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
    
'    'Left-needle will be gone up.
'    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
'    ReadValue = ReadValue And &HF7FF
'    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
'    Call Sleep(0.2)
    
    If (purgePartArray = False) Then
        Enable_Button
    End If
    
    purgeButtonTimer.Enabled = True
End Sub

Private Sub Always_Purge_Part_Array()

    Leftslider_go_up
    
'    If systemMoveHeight > systemHomeZ Then
'        Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "100")))
'    End If
'
'    Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "100")))
'
'    Sleep (0.5)
    
    Always_Purge
End Sub

'''''''''''''''''''''''''''''
'   Move to System Height   '
'''''''''''''''''''''''''''''
Private Sub Move_System_Height()
    Dim xpos  As Long, ypos As Long, zpos As Long, upos As Long
    
    Call P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, Lcnt, xpos, ypos, zpos, upos)
    doPtp xpos, ypos, systemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50"))
End Sub

'''''''''''''''''''''''''''''''''''''''''
'   New procedure for "LinksLinePoint"  '
'''''''''''''''''''''''''''''''''''''''''
Private Function doContiLine3D_XW(endx As Long, endy As Long, endz As Long)
    Position(0) = endx
    Position(1) = endy
    Position(2) = endz
End Function

Private Function doSegmentProperty3D_XW(Speed As Long, dispenser As Integer, sequenceNum As Long)
    returncode = P1240MotAxisParaSet(boardNum, 0, &H7, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)
    'returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)
    
    If (dispenser = False Or estopValue = True Or abortValue = True) Then
        Call doDispenseOff
    Else
        Call doDispenseOn
    End If
     
    returncode = P1240MotLine(boardNum, X_axis Or Y_axis Or Z_axis, 0, Position(0), Position(1), Position(2), 0)
    Do While ((P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success) And (estopValue = False) And (abortValue = False))
        If (CloseBoard = True) Then
            Exit Function
        End If
        DoEvents
    Loop
End Function


Private Sub wetRun_Click()
    SetLightIntensity (0)
End Sub
