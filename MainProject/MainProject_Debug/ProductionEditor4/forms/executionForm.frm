VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form executionForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pattern Execution Form"
   ClientHeight    =   4395
   ClientLeft      =   5850
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   9360
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat"
      Height          =   255
      Left            =   9600
      TabIndex        =   21
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtPurgeTime 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   8760
      TabIndex        =   18
      Text            =   "1.00"
      Top             =   2280
      Width           =   495
   End
   Begin VB.CheckBox AlwaysPurge 
      Caption         =   "Always Purge"
      Height          =   250
      Left            =   6480
      TabIndex        =   17
      Top             =   2280
      Width           =   1258
   End
   Begin VB.Frame NeedleMode 
      Caption         =   "Needle Mode"
      Height          =   560
      Left            =   9600
      TabIndex        =   14
      Top             =   1560
      Width           =   2775
      Begin VB.OptionButton RightNeedle 
         Caption         =   "Right "
         Height          =   250
         Left            =   1440
         TabIndex        =   16
         Top             =   230
         Width           =   975
      End
      Begin VB.OptionButton LeftNeedle 
         Caption         =   "Left "
         Height          =   250
         Left            =   240
         TabIndex        =   15
         Top             =   230
         Width           =   855
      End
   End
   Begin VB.Timer purgeButtonTimer 
      Interval        =   200
      Left            =   4800
      Top             =   0
   End
   Begin VB.CommandButton PurgePosition 
      Caption         =   "Purge Position"
      Height          =   375
      Left            =   9720
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton abortButton 
      Caption         =   "Abort"
      Height          =   375
      Left            =   4920
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Timer redrawValveOnOff 
      Interval        =   250
      Left            =   5520
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "executionForm.frx":0000
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin VB.PictureBox PictureValveOnOff 
      Height          =   735
      Left            =   6480
      ScaleHeight     =   675
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   3000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Test Run"
      Height          =   800
      Left            =   6480
      TabIndex        =   7
      Top             =   600
      Width           =   2775
      Begin VB.OptionButton Camera 
         Caption         =   "Camera"
         Height          =   250
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton wetRun 
         Caption         =   "Wet"
         Height          =   250
         Left            =   1440
         TabIndex        =   9
         Top             =   230
         Width           =   735
      End
      Begin VB.OptionButton dryRun 
         Caption         =   "Dry"
         Height          =   250
         Left            =   240
         TabIndex        =   8
         Top             =   230
         Width           =   1095
      End
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
      Left            =   4200
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "executionForm.frx":0078
      Top             =   0
   End
   Begin VB.CommandButton closeButton 
      Caption         =   "Close"
      Height          =   375
      Left            =   7920
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txtMessage 
      Height          =   3615
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   6015
   End
   Begin VB.CommandButton cmdResumeButton 
      Caption         =   "Resume"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdStopButton 
      Caption         =   "Pause"
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdStartButton 
      Caption         =   "Run"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdPurgeButton 
      Caption         =   "Purge"
      Height          =   375
      Left            =   7920
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.CommandButton homeButton 
      Caption         =   "Home"
      Height          =   375
      Left            =   7920
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel PurgeTime 
      Height          =   255
      Left            =   7770
      OleObjectBlob   =   "executionForm.frx":02AC
      TabIndex        =   19
      Top             =   2295
      Width           =   945
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
Dim SignalCounter As Long               'Count enable signal for purge timer
Dim PrePositionX As Long, PrePositionY As Long, PrePositionZ As Long, PrePositionU   'Save PrePosition
Dim DisablePurgeSignal As Boolean
Dim ClickPurgePosition As Boolean       'To recover the removal first point in translation (XW)
Dim LeftValve As Boolean                'Just a flag for choosing left-valve.
Dim RightValve As Boolean               'Just a flag for choosing right-valve.
Dim L_Dispense As Boolean, R_Dispense As Boolean    'Just flag which valve should be dispensed
Dim Button_Purge As Boolean
Dim Rotation_Angle_U As Long            'Save data for rotation
Dim Rotation_Flag As Boolean, Retilt As Boolean, Spray_Valve As Boolean
Dim Position(0 To 2) As Long            'Save X,Y and Z position
Dim Previous_Angle As Long              'Save the angle and compare with the curren one
Dim firstTimeLeft As Boolean            'Flag for part array to avoid Up/Down many times
Dim firstTimeRight As Boolean           'Flag for part array to avoid Up/Down many times
Dim purgePartArray As Boolean           'Flag for auto purging not to do "Enable/Disable GUI" in partArray

Dim countUp As Integer

Private Sub PrintParseTree(Text As String)
    txtMessage.Text = txtMessage.Text & Text & vbNewLine
    txtMessage.SelStart = 65535
End Sub

'These two procedures may not be used
'Private Sub RightNeedle_Click()
'    'Do it when changing from the right valve.
'    If (LeftValve = False) And (RightValve = True) Then
'        LeftNeedleValve
'    End If
'
'    LeftValve = True
'    RightValve = False
'End Sub
'
'Private Sub LeftNeedle_Click()
'    'Do it when changing from the left valve.
'    If (LeftValve = True) And (RightValve = False) Then
'        RightNeedleValve
'    End If
'
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
    End If
    PrintParseTree ("Dispensing Abort!")
    abortButton.Enabled = False
    
    prevDispenserValue = False
End Sub

Private Sub Camera_Click()
    SetLightIntensity Val((editorForm.LightingIntensity.Text))
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub

Private Sub cmdPurgeButton_Mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Original (NYP)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H0)
    'purgeButtonTimer.Enabled = True
    
    Button_Purge = False
 
End Sub

Private Sub cmdPurgeButton_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Original (NYP)
    'purgeButtonTimer.Enabled = False
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H100)
    
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
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.3)
    'To travel the right way (XW)
    If (SystemMoveHeight > zpos) Then
        doPtp xpos, ypos, SystemMoveHeight, zSystemTravelSpeed
    Else
        doPtp xpos, ypos, zpos, zSystemTravelSpeed
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
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H100
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        ElseIf leftside = True And rightside = False Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H800
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
        Else
            'If (executionForm.LeftNeedle.value = True) Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H800
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
            'ElseIf (executionForm.RightNeedle.value = True) Then
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H100
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
            'End If
        End If
        DoEvents
    Loop
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HFEFF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    Call Sleep(0.3)
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    doPtp PrePositionX, PrePositionY, SystemMoveHeight, zSystemTravelSpeed
    doPtp PrePositionX, PrePositionY, PrePositionZ, xySystemTravelSpeed
    
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
    
    'Go and Purge before starting spray
    If (AlwaysPurge.value = 1) Then
        Always_Purge
    End If
    
    abortValue = False
    
    closeButton.Enabled = False
    homeButton.Enabled = False
    
    cmdStopButton.Enabled = True
    cmdResumeButton.Enabled = True
    abortButton.Enabled = True
    
    'To prevent startbutton flicker
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    
    Call doNormalDispense
    
    'To prevent startbutton flicker
    startButtonTimer.Enabled = True
    purgeButtonTimer.Enabled = True
    
    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    
    'doHome
    
    cmdStartButton.Enabled = True
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
End Sub

Private Sub cmdResumeButton_Click()
    PrintParseTree ("Dispensing Resume...")
    stopValue = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = True        '(XW)
End Sub

Private Sub cmdStopButton_Click()
    PrintParseTree ("Dispensing Pause!")
    stopValue = True
    cmdStopButton.Enabled = False       'Should not be pressed more than one time (XW)
    cmdResumeButton.Enabled = True
    resumeTimer.Enabled = True
End Sub

Private Sub dryRun_Click()
    Call SetLightIntensity(0)
End Sub

Private Sub Form_Load()
    
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    
    'End all of the program after pressing the E-Stop button    'XW
    If Close_Emg = True Then
        End
    End If
    
    editorForm.CheckEmergencyStop.Enabled = False       'XW
       
    abortValue = False
    estopValue = False
    stopValue = False
    
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    startButtonTimer.Enabled = False
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
    startButtonTimer.Enabled = True
    
    
    'directSoftHomeOption = False
    prevDispenserValue = False
    printErrorLimit = True
        
    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    doingHome = False
    
    'doHome
    'To provide a shortcut for machine home 3 May 2005
    SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    SystemMoveHeight = SystemMoveHeight * (-1)
    
    txtPurgeTime.Text = Format(GetStringSetting("EpoxyDispenser", "Setup", "DispensingTime", "10"), "##0.00")
    
    
    'doPtp 0, 0, 0, 50
    doPtp 0, 0, 0, 100
        
    doSoftHome
    Leftslider_go_down
    
    'LeftNeedleValve                         'Set Default as Let_Cylinder
    'LeftValve = True                        'Just a default flag
    'RightValve = False
    
    cmdStartButton.Enabled = True
    'cmdStopButton.Enabled = False
    'cmdResumeButton.Enabled = False
    'abortButton.Enabled = False
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    purgeButtonTimer.Enabled = True
    
    deltaX = 0
    deltaY = 0
    deltaAngle = 0
    xOrgFid = 0
    yOrgFid = 0
    
End Sub

Private Sub doNormalDispense()

    cmdStartButton.Enabled = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = True
    pauseButtonTimer.Enabled = True
    PurgePosition.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    abortButton.Enabled = True
    NeedleMode.Enabled = False
    abortButtonTimer.Enabled = True
    
    firstTimeLeft = False
    firstTimeRight = False
    
    'Check the metarial before starting the program
    If rightside = True Then
        Low_Level
    End If
    
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
    
    PrintParseTree ("Dispensing Complete!")
    
    'If the user presses "Abort" button, move "System Height" for safety.
    If (abortValue = True) Then
        Dim ValveClose As Long
        
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
    'executionForm.LeftNeedle.value = True
    cmdStartButton.Refresh
    
    systemTrackMoveHeight = SystemMoveHeight
    
    If (estopValue = False) Then
        cmdStopButton.Enabled = False
        pauseButtonTimer.Enabled = False
        purgeButtonTimer.Enabled = False
        'startButtonTimer.Enabled = True
        'cmdStopButton.Enabled = False
                
        'doSoftHome
        If GetStringSetting("EpoxyDispenser", "Setup", "AlwaysRobotHome", "0") = "1" Then
            doPtp 0, 0, 0, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50"))
            doSoftHome
            Leftslider_go_down
        Else
            'Call doPtp(0, 0, systemMoveHeight, 50)
                        
            'Origin(XW)
            'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
        
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
            Leftslider_go_down
        End If
        
        PurgePosition.Enabled = True
        cmdPurgeButton.Enabled = True
        closeButton.Enabled = True

        startButtonTimer.Enabled = True
        purgeButtonTimer.Enabled = True

        cmdStartButton.Enabled = True
        'cmdStopButton.Enabled = True
        abortButton.Enabled = False
        NeedleMode.Enabled = True
    Else
        cmdStopButton.Enabled = False
        pauseButtonTimer.Enabled = False
        cmdStartButton.Enabled = False
        PurgePosition.Enabled = False
        cmdPurgeButton.Enabled = False
        closeButton.Enabled = False
    End If
    
    'Check the metarial after finishing the program
    If rightside = True Then
        Low_Level
    End If
    Previous_Angle = 0
    Rotation_Flag = False
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
                    
                    LeftNeedle.value = False
                    L_Dispense = True
                    R_Dispense = False
                    
                    RightNeedleValve
                    
                    RightNeedle.value = True
                    L_Dispense = False
                    R_Dispense = True
                    
'                    Tilt_Off
'                    'Need to go back to "angle 0" because we don't know all offset value
'                    If (Spray_Valve = True) Then
'                        Rotation_U (0)
'                    End If
'
'                    Spray_Valve = False
                    Spray_Valve = True            '@$K


                    firstTimeLeft = True
                    firstTimeRight = False
                End If
            ElseIf (66666 = CLng(endx + needleOffsetX)) And (66666 = CLng(endy - needleOffsetY)) And (endz = 66666) And (Speed = 66666) Then
                If (firstTimeRight = False) Then
                    Move_System_Height
                
                    RightNeedle.value = False
                    R_Dispense = True
                    L_Dispense = False
                    
                    LeftNeedleValve
                    
                    LeftNeedle.value = True
                    R_Dispense = False
                    L_Dispense = True
                    
                    'Spray_Valve = True
                    Spray_Valve = False    '@$K
                    
                    Tilt_Off
                    'Need to go back to "angle 0" because we don't know all offset value
                    If (Spray_Valve = True) Then
                        Rotation_U (0)
                    End If
                    
                    firstTimeRight = True
                    firstTimeLeft = False
                End If
            ElseIf (endx = 0) And (endy = 0) And (endz = 0) And (Speed = 77777) Then
                'Go to the system home and do the purging for part array
                If (AlwaysPurge.value = 1) Then
                    purgePartArray = True
                    Always_Purge_Part_Array
                    purgePartArray = False
                End If
            ElseIf (77777 = CLng(endx + needleOffsetX)) And (77777 = CLng(endy - needleOffsetY)) And (endz = 77777) And (Speed = 77777) Then
                'Go to the system home and do the purging for part array
                If (AlwaysPurge.value = 1) Then
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
'                If (Previous_Angle <> 36363) And (Camera.value = False) Then
'                    Tilt_ON
'                    Rotation_U (0)
'                    Previous_Angle = 36363
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 99999) Then
'                If (Previous_Angle <> 99999) And (Camera.value = False) Then
'                    Tilt_ON
'                    'Rotation_U (2500)
'                    Rotation_U (900)
'                    Previous_Angle = 99999
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 18181) Then
'                If (Previous_Angle <> 18181) And (Camera.value = False) Then
'                    Tilt_ON
'                    'Rotation_U (5000)
'                    Rotation_U (1800)
'                    Previous_Angle = 18181
'                    Rotation_Flag = True
'                End If
'            ElseIf (0 = CLng(endx + needleOffsetX)) And (0 = CLng(endy - needleOffsetY)) And (endz = 0) And (Speed = 27272) Then
'                If (Previous_Angle <> 27272) And (Camera.value = False) Then
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
                If (Camera.value = True) Then
                    endx = CLng(Trim(word3(0)))
                    endy = CLng(Trim(word4(0)))
                    endz = 0
                    
                    origX = endx
                    origY = endy
                                
                    detXYfromFudicial endx, endy, endx, endy, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
                
                    presentX = endx
                    presentY = endy
                   
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
                    origX = endx
                    origY = endy
                                
                    detXYfromFudicial endx, endy, endx, endy, convertToPulses(xOrgFid, X_axis), convertToPulses(yOrgFid, Y_axis), convertToPulses(deltaX, X_axis), convertToPulses(deltaY, Y_axis), deltaAngle
                
                    presentX = endx
                    presentY = endy
                
                    'If "Purge Position" is higher than "systemMoveHeight", the robort should take z-Height of "Purge Position".
                    If (Spray_Valve = False) Then
                        If (ClickPurgePosition = True) Then
                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                            ClickPurgePosition = False
                        Else
                            Call doPtp(endx + Offset_DistanceX_Camera_L_Needle, endy + Offset_DistanceY_Camera_L_Needle, endz - needleOffsetZ_L, Speed)
                        End If
                    Else
                        If (Rotation_Flag = False) Then
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, endz - needleOffsetZ_R, Speed)
                            End If
                        Else
                            If (ClickPurgePosition = True) Then
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis)) * (-1), Speed)
                                ClickPurgePosition = False
                            Else
                                Call doPtp(endx + Offset_DistanceX_Camera_R_Needle, endy + Offset_DistanceY_Camera_R_Needle, endz - needleOffsetZ_R, Speed)
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
            'If (RightNeedle.value = True) Then
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
            If Camera.value = True Then
                SetLightIntensity (Val(editorForm.LightingIntensity.Text))
            Else
                SetLightIntensity (0)
            End If
    End Select
End Sub

Private Sub doFudicial(x1 As Long, y1 As Long, x2 As Long, y2 As Long, FileName As String, LightingAmt As Long)
    
    Dim patFile As String
    Dim Length As Long
    
    'Move z_axis to "zero" before doing fiducial (XW)
    checkSuccess (P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(20, Z_axis), 2000000, 1200000, 9158400))
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
    Loop
    
    xOrgFid = convertToMM(x1, X_axis)
    yOrgFid = convertToMM(y1, Y_axis)
    'yOrgFid = convertToMM(x2, Y_axis)
    
    Length = Len(FileName)

    patFile = Mid(FileName, 2, Length - 2)
    
    'Use "S-curve" (XW)
    'returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis), MaxVelocity, AccelSpeed, AccelRate)
    returncode = P1240MotAxisParaSet(boardNum, 0, X_axis Or Y_axis Or Z_axis, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis), MaxVelocity, AccelSpeed, AccelRate)

    visionRetry = True
    
    SetLightIntensity (LightingAmt)
    
    Do While (visionRetry)
        returncode = VdeFindRefPt(patFile, convertToMM(x1, X_axis), convertToMM(y1, Y_axis), convertToMM(x2, X_axis), convertToMM(y2, Y_axis), deltaX, deltaY, deltaAngle)
        If returncode <> 1 Then
            RetryFudicial.Show (vbModal)
        Else
            visionRetry = False
            doneFudicial = True
        End If
    Loop

    VdeSelectCamera 2
    VdeCameraLive 1
End Sub


Private Function doContiStart(Speed As Long, dispense As Integer, axisMovement As Long)
    
    returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)

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

        Do While ((PresentSequenceNum < sequenceNum) And (estopValue = False) And (abortValue = False))
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

Private Function Rotate(ByVal Rotate_Angle As Long)
    '''''''''''''''''''''''''''''''''''''
    '   Close Valve and do rotation     '
    '''''''''''''''''''''''''''''''''''''
    If (Rotate_Angle = 10101) Then
        'Tilt_Off
        'doDelay (0.2)      =>speed=10
        'doDelay (0.15)     =>speed=20
        'doDelay (0.12)     =>speed=30
        'doDelay (0.08)     =>speed=40
        'doDelay (0.06)     =>speed=50,60
        'doDelay (0.04)     =>speed=70
        'doDelay (0.03)     =>speed=80,90
        doDelay (0.01)
        Retilt = True
        Rotation_Flag = False
        Exit Function
    End If
    
    If (Retilt = True) Then
        'Tilt_on
        Retilt = False
    End If
    
    'Call doDispenseOff
    If (Rotate_Angle = 36363) Then
        Rotation_U (0)
    ElseIf (Rotate_Angle = 99999) Then
        Rotation_U (1250)
    ElseIf (Rotate_Angle = 18181) Then
        Rotation_U (2500)
    ElseIf (Rotate_Angle = 27272) Then
        Rotation_U (3750)
    'ElseIf (Rotate_Angle = 10101) Then
    '    'Tilt_Off
    End If
    
    'doDelay (0.22)     =>speed= 1 to 90
    'doDelay (0.2)      =>speed=100
    'doDelay (0.18)     =>speed=110
    'doDelay (0.17)     =>speed=120
    'doDelay (0.15)     =>speed=130
    'doDelay (0.14)     =>speed=140,150
    'doDelay (0.12)
    Rotation_Flag = False
End Function

Private Function doContiArc(endx As Long, endy As Long, CenX As Long, ceny As Long, ccw)
        
    If (ccw = 1) Then
        contiPathArray.PathType = IPO_CCW
    Else
        contiPathArray.PathType = IPO_CW
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
    returncode = P1240InitialContiBuf(0, 400)
    
    Debug.Print "Initialize contibuffer " & returncode
    
    contiPathArrayIndex = 1
End Function

Private Function doDispenseOn()
    Dim ReadValve As Long
    
    If wetRun.value = True Then
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
            
            If (L_Dispense = True) Then
            'If (R_Dispense = True) Then
                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValve)
                ReadValve = ReadValve Or &H800
                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValve)
            ElseIf (R_Dispense = True) Then
            'ElseIf (L_Dispense = True) Then
                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValve)
                ReadValve = ReadValve Or &H100
                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValve)
            End If
        End If
    
        prevDispenserValue = True
        
    End If
    
End Function

Private Function doDispenseOff()
    Dim ReadValue As Long
    
    If (prevDispenserValue = True) Then
    
        Debug.Print "Dispense Off"
    
        'Origin (NYP)
        'returncode = P1240MotWrReg(0, 4, WR3, &H0)
        
        returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HF7FF
        returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
    
        returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HFEFF
        returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
    End If
    
    prevDispenserValue = False
    
End Function

Private Function doPtp(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal Speed As Long)
     
    Debug.Print "DoPtp " & X & " " & y & " " & Z & " " & Speed
    
    Dim AccelSpeed, AccelSpeedZ As Double
    Dim AccelRate, AccelRateZ As Double
    Dim factor, factorZ As Double
    Dim value, ValueX, ValueY, ValueZ, valueU As Long
        
    If ballScrew = 1 Then
        
        '0.12G
        AccelSpeedZ = 1200000
        AccelRateZ = 9158400
        
        '0.2G
        'AccelSpeed = 2000000
        'AccelRate = 15264000
        
        AccelSpeed = 5000000
        AccelRate = 30000000
        
'        AccelSpeedZ = 260000       'origin
'        AccelRateZ = 500000
'        AccelSpeed = 260000
'        AccelRate = 500000
     
'        If Speed <= 10 Then
'            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
'            AccelRateZ = AccelSpeedZ / 0.05
'
'            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.1
'            AccelRate = AccelSpeed / 0.03
'        ElseIf Speed <= 90 Then
'            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
'            AccelRateZ = AccelSpeedZ / 0.05
'
'            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.2
'            AccelRate = AccelSpeed / 0.06
'        Else
'            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.8
'            AccelRateZ = AccelSpeedZ / 0.08
'
'            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
'            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.7
'            AccelRate = AccelSpeed / 0.2
'
'        End If
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
    returncode = P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(Speed, Z_axis), 200000, AccelSpeedZ, AccelRateZ)
    
    Do While (stopValue = True And estopValue = False And abortValue = False)
        DoEvents
        pauseButtonTimer.Enabled = True
    Loop
    
    If (X = 0 And y = 0 And Z = 0) Then
        'returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, -1000, 0)
        returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, -100, 0)
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
            DoEvents
        Loop
        'returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, 1000, -1000, 0, 0)
        returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, 100, -100, 0, 0)
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
            DoEvents
        Loop
    Else
        'XW (Can't give 0 pulse)
        'If (Z < 2) And (Z > -2) Then
        '    returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, -10, 0)
        'Else
        '    If (Z <> Pt_Z) Then
                returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, Z, 0)
        '    End If
        'End If
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
            DoEvents
        Loop
        'If ((x < 2) And (x > -2)) And ((y < 2) And (y > -2)) Then
        '    returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, 10, -10, 0, 0)
        'Else
        '    If (x <> Pt_X) And (y <> Pt_Y) Then
                returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, X, y, 0, 0)
        '    ElseIf (x = Pt_X) And (y <> Pt_Y) Then
        '        returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, x + 10, y, 0, 0)
        '    ElseIf (x <> Pt_X) And (y = Pt_Y) Then
        '        returncode = P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, x, y + 10, 0, 0)
        '    End If
        'End If
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
            DoEvents
        Loop
    End If
    proceed = False
    checkProceed
            
    cmdStartButton.Enabled = False
           
    If (X = 0 And y = 0 And Z = 0) Then
        home_limit_flag = True  'xu long
        'Change T_curve to S_curve
        If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(20, Z_axis), 200000, 1200000, 9158400))) Then
            If (checkSuccess(P1240MotCmove(boardNum, Z_axis, 0))) Then   'Move Z in clockwise direction
                value = 0
                valueU = 0
                Do While ((value And &H4) <> &H4)  'Do loop if Z Limit switch still not reached
                    checkSuccess (P1240MotRdReg(boardNum, Z_axis, RR2, value))
                    DoEvents
                Loop
                If ((value And &H4) = &H4) Then 'Do immediate stop on Z axis
                    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
                End If
              
                Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
                Loop
                If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
                    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
                        DoEvents        'XW
                    Loop
                    
                    '''''''''''''''''''''
                    '   U_axis (Homing) '
                    '''''''''''''''''''''
                    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 300, 300, 8300, 53000, 9000000))
                    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 600, 600, 16600, 106000, 18000000))
                    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 500, 5000, 50000, 9000000))
    
                    checkSuccess (P1240MotCmove(boardNum, U_axis, 8))
                    valueU = 0
                
                    Do While ((valueU And &H8) <> &H8)
                        checkSuccess (P1240MotRdReg(boardNum, U_axis, RR2, valueU))
            
                        If ((valueU And &H8) = &H8) Then 'Do immediate stop on U axis
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
                            ValueX = 0
                            ValueY = 0
                            Do While (((ValueX And &H8) <> &H8) Or ((ValueY And &H4) <> &H4))
                                checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, ValueX, ValueY, ValueZ, valueU))
                                If ((ValueY And &H4) = &H4) Then
                                    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
                                End If
                                If ((ValueX And &H8) = &H8) Then
                                    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
                                End If
                                DoEvents
                            Loop
                            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                            Loop
                            If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
                                Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                                    DoEvents        'XW
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
    
    '''''''''''''''''
    '   Origin      '
    '''''''''''''''''
    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 3000, 10000, 50000, 9000000))
    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 500, 10000, 50000, 9000000))
    
    'Save previous position
    Pt_X = X
    Pt_Y = y
    Pt_Z = Z
    
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

    'returncode = P1240MotRdReg(boardNum, X_axis, RR0, buttonValue) 'origin
    'buttonValue = buttonValue And &HF0                             'origin
    
    returncode = P1240MotRdReg(boardNum, X_axis, RR2, buttonValue)
    buttonValue = buttonValue And &H20
    
    'If (buttonValue = &HF0) Then       'origin
    
    If (buttonValue = &H20) Then
        abortButtonTimer.Enabled = False
        startButtonTimer.Enabled = False
        pauseButtonTimer.Enabled = False
        purgeButtonTimer.Enabled = False
        
        Servo_Off
         
        'Close valve first before the Emergency Form come out
        'Call P1240MotRdReg(boardNum, Z_axis, WR3, ValveClose)
        'If (ValveClose = &H100) Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
            ValveClose = (ValveClose And &HF7FF)
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
            
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, &H0))
            
'            'Right-needle will be gone up.
'            checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
'            ReadValue = ReadValue And &HFEFF
'            checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
    
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
            Exit Sub
        End If
        
    End If
    eStopTimer.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim Tower_Light_Value As Long
    
    'To do the closed board and release the buffer after pressing the E-stop button
    'XW
    If Close_Emg = True Then
        editorForm.displayCoOrdsTimer.Enabled = False
        editorForm.CheckEmergencyStop.Enabled = False
        editorForm.TimerDrawStatus.Enabled = False
        
        StopTimer
        EstopLimit
        
        'Disable Red_Light,Yellow_light and Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Tower_Light_Value = Tower_Light_Value And &HF1FF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
        
        WindowFree editorForm.lstPattern.hWnd
        Call Sleep(0.5)
        returncode = P1240FreeContiBuf(0)
        VdeCameraLive False
        VdeReleaseVision
        
        Close_PCI1750
        unInitializeBoard
        
        End
        Exit Sub
    Else
        'Disable Red_Light,Yellow_light and Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Tower_Light_Value = Tower_Light_Value And &H400
        If (Tower_Light_Value <> &H0) Then
            Yellow_Light = False
        End If
        
        VdeCameraLive 1
        'Origin (NYP)
        'setSpeed (editorForm.jogSpeedSlider.value - 1)
        setSpeed (28)
    End If
    
    SaveStringSetting "EpoxyDispenser", "Setup", "DispensingTime", txtPurgeTime.Text
    editorForm.CheckEmergencyStop.Enabled = True
    
    SetLightIntensity (Val(editorForm.LightingIntensity.Text))
End Sub

Private Sub homeButton_Click()
    purgeButtonTimer.Enabled = False
    cmdStartButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdResumeButton.Enabled = False
    abortButton.Enabled = False
    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    PurgePosition.Enabled = False
    doingHome = True

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
    'cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    PurgePosition.Enabled = True
    purgeButtonTimer.Enabled = True
    doingHome = False
End Sub

Private Sub PurgePosition_Click()
    
    Dim xpos, ypos, zpos As Long
    
    xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
    ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
    'To get the actual direction        'XW
    ypos = ypos * (-1)
    zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
    zpos = zpos * (-1)
    
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False

    homeButton.Enabled = False
    cmdPurgeButton.Enabled = False
    closeButton.Enabled = False
    abortButton.Enabled = False
    cmdResumeButton.Enabled = False
    cmdStopButton.Enabled = False
    cmdStartButton.Enabled = False
        
    'To travel the right way (XW)
    If (SystemMoveHeight > zpos) Then
        doPtp xpos, ypos, SystemMoveHeight, zSystemTravelSpeed
    Else
        doPtp xpos, ypos, zpos, zSystemTravelSpeed
        ClickPurgePosition = True
    End If
    
    doPtp xpos, ypos, zpos, xySystemTravelSpeed
    
    startButtonTimer.Enabled = True
    purgeButtonTimer.Enabled = True
            
    homeButton.Enabled = True
    cmdPurgeButton.Enabled = True
    closeButton.Enabled = True
    'abortButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'cmdStopButton.Enabled = True
    cmdStartButton.Enabled = True
    
End Sub

Private Sub redrawValveOnOff_Timer()
    Dim PW, PH
    Dim ValveStatus As Long, ValveStatus2 As Long
   
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
End Sub

Private Sub abortButtonTimer_Timer()
    
    Dim buttonValue As Long
    Dim skip As Boolean
    
    abortButtonTimer.Enabled = False
    
    'Check Door Lock Sensor (Chang as "Pause")
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
    
End Sub

Private Sub resumeTimer_Timer()

    Dim buttonValue, pauseButtonValue As Long
    Dim skip As Boolean

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
    
End Sub

Private Sub startButtonTimer_Timer()

    Dim buttonValue1, buttonValue2, abortButtonValue As Long
    Dim skip As Boolean
    
    If (Inter_Lock = True) Then
        'Check Door Lock Sensor
        If (Door_Lock = True) Then
            MsgBox "Please lock the door first before running the application!"
            Exit Sub
        End If
    End If
    
    startButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    'Move to "doNormalDispense" procedure
    'abortButtonTimer.Enabled = True
   
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
        abortValue = False
        
        'Do the purging first
        If (AlwaysPurge.value = 1) Then
            Always_Purge
        End If
                
        Call doNormalDispense
        'If (estopValue = False) Then
            'doHome
            'doPtp 0, 0, 0, 50
            'doSoftHome 'See comments of 300605
        'End If
    Else
        If (Check1.value = 1) Then
            countUp = countUp + 1
            If (countUp = 5) Then
                stopValue = False
                pauseButtonTimer.Enabled = True
                skip = True
                abortValue = False
        
                'Do the purging first
                If (AlwaysPurge.value = 1) Then
                    Always_Purge
                End If
                
                Call doNormalDispense
                
                countUp = 0
            End If
        Else
            countUp = 0
        End If
    End If
    
    If skip = False Then
        startButtonTimer.Enabled = True
    End If
    
    cmdStartButton.Enabled = True
    cmdPurgeButton.Enabled = True
    
    purgeButtonTimer.Enabled = True
    
    If doingHome = False Then
        closeButton.Enabled = True
    End If
    'cmdStopButton.Enabled = True
    'cmdResumeButton.Enabled = True
    'abortButton.Enabled = True
    homeButton.Enabled = True

End Sub

Private Sub purgeButtonTimer_Timer()        'XW
    
    Dim buttonValue As Long, ReadValue As Long
    
    purgeButtonTimer.Enabled = False
    
    'Purge_Button input
    returncode = P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue)
    buttonValue = buttonValue And &H400
    
    If (buttonValue = 0) Then
        'PrintParseTree ("Purge...")
        
        Disable_Button
        cmdPurgeButton.Enabled = False
        
        Dim xpos, ypos, zpos As Long
    
        xpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0"), X_axis))
        ypos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0"), Y_axis))
        'To get the actual direction        'XW
        ypos = ypos * (-1)
        zpos = CLng(convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0"), Z_axis))
        zpos = zpos * (-1)
        
        DisablePurgeSignal = True
        checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
          
        'Left-needle will be gone up.
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
        ReadValue = ReadValue And &HF7FF
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        Call Sleep(0.3)
        
        'To travel the right way (XW)
        If (SystemMoveHeight > zpos) Then
            doPtp xpos, ypos, SystemMoveHeight, zSystemTravelSpeed
        Else
            doPtp xpos, ypos, zpos, zSystemTravelSpeed
            ClickPurgePosition = True
        End If
    
        doPtp xpos, ypos, zpos, xySystemTravelSpeed
        
        'Left slider go down
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
        ReadValue = ReadValue Or &H800
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        Call Sleep(0.3)
        
        Do While (buttonValue = 0)
            If leftside = False And rightside = True Then
                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
                ReadValue = ReadValue Or &H100
                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
            ElseIf leftside = True And rightside = False Then
                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
                ReadValue = ReadValue Or &H800
                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
            Else
                'If (executionForm.LeftNeedle.value = True) Then
                returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
                ReadValue = ReadValue Or &H800
                returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
                'ElseIf (executionForm.RightNeedle.value = True) Then
                returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
                ReadValue = ReadValue Or &H100
                returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
                'End If
            End If
            'Purge_Button input
            returncode = P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue)
            buttonValue = buttonValue And &H400
        Loop
    
        returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HF7FF
        returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
        
        returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HFEFF
        returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
            
        Call Sleep(0.3)
        'Left-needle will be gone up.
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
        ReadValue = ReadValue And &HF7FF
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        
        doPtp PrePositionX, PrePositionY, SystemMoveHeight, zSystemTravelSpeed
        doPtp PrePositionX, PrePositionY, PrePositionZ, xySystemTravelSpeed
        'PTPToXYZ PrePositionX, PrePositionY, PrePositionZ
        
        'Left slider go down
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
        ReadValue = ReadValue Or &H800
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
        cmdPurgeButton.Enabled = True
        Enable_Button
        DisablePurgeSignal = False
    End If
    
    purgeButtonTimer.Enabled = True
   
End Sub

Private Sub pauseButtonTimer_Timer()

    Dim buttonValue, startbuttonValue1, startbuttonValue2 As Long
    Dim skip As Boolean

    pauseButtonTimer.Enabled = False
    
    If (Inter_Lock = True) Then
        'Check Door Lock Sensor
        If (Door_Lock = True) Then
            Dim ValveClose As Long
                
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose)
            ValveClose = (ValveClose And &HF7FF)
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose)
            
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H0)
            
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
        resumeTimer.Enabled = True
        cmdStopButton.Enabled = False       'Should not be pressed more than one time (XW)
        cmdResumeButton.Enabled = True
        skip = True
    End If
    
    If skip = False Then
        pauseButtonTimer.Enabled = True
    End If
    
End Sub

Private Sub Timer1_Timer()
    Dim xlimit, ylimit, zlimit, ulimit As Long
    
    If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) = Success) And (proceed = False) Then
        proceed = True
    End If
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, RR2, xlimit, ylimit, zlimit, ulimit))
    xlimit = xlimit And &HC
    ylimit = ylimit And &HC
    zlimit = zlimit And &HC
    ulimit = ulimit And &H8
    
    If (((xlimit <> 0) Or (ylimit <> 0) Or (zlimit <> 0) Or (ulimit <> 0)) And printErrorLimit = True) And home_limit_flag = False Then
    'If (((xlimit <> 0) Or (ylimit <> 0) Or (zlimit <> 0)) And printErrorLimit = True) Then     'origin
        PrintParseTree ("Error limit reach!")
        printErrorLimit = False
        Call abortButton_Click
    End If
        
    If (((xlimit = 0) And (ylimit = 0) And (zlimit = 0) And (ulimit = 0)) And printErrorLimit = False) Then
        printErrorLimit = True
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
        
        systemTrackMoveHeight = SystemMoveHeight

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
            
            'Origin(XW)
            'returncode = P1240MotLine(boardNum, X_axis Or Y_axis Or Z_axis, 1, systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, 0)
            returncode = P1240MotLine(boardNum, X_axis Or Y_axis Or Z_axis, 1, systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemHomeZ, 0)
            
            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
                DoEvents
            Loop
        
            'returncode = P1240FreeContiBuf(0)
        
        Else

            'Call doPtp(0, 0, systemMoveHeight, 50)
            If SystemMoveHeight > systemHomeZ Then
                'Origin (XW)
                'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
                Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            End If
            'Origin
            'Call doPtp(systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
            Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))

        End If

        If (rememberToTurnOffTimer1 = True) Then
            Timer1.Enabled = False
        End If
    End If
End Sub

Private Sub doHome()
    
    Timer1.Enabled = False
    
    cmdStartButton.Enabled = False
    
    Dim value, ValueX, ValueY, ValueZ, valueU As Long
    
    Dim speed123 As Long
    
    If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, 0, 1000, convertSpeed(20, Z_axis), 2000000, 1200000, 9158400))) Then
        If (checkSuccess(P1240MotCmove(boardNum, Z_axis, 0))) Then 'Move Z in clockwise direction
            value = 0
            Do While ((value And &H4) <> &H4) 'Do loop if Z Limit switch still not reached
                checkSuccess (P1240MotRdReg(boardNum, Z_axis, RR2, value))
                DoEvents
            Loop
            If ((value And &H4) = &H4) Then 'Do immediate stop on Z axis
                checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            End If
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
            Loop
            If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
                Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
                    DoEvents        'XW
                Loop
                If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, 0, StartVelocity, convertSpeed(20, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
                    If (checkSuccess(P1240MotCmove(boardNum, X_axis, 1))) Or (checkSuccess(P1240MotCmove(boardNum, Y_axis, 0))) Then
                    'Move X and Y motors in clockwise direction and anti-clockwise direction
                        ValueX = 0
                        ValueY = 0
                        Do While (((ValueX And &H8) <> &H8) Or ((ValueY And &H4) <> &H4))
                            checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, ValueX, ValueY, ValueZ, valueU))
                            If ((ValueY And &H4) = &H4) Then
                                checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
                            End If
                            If ((ValueX And &H8) = &H8) Then
                                checkSuccess (P1240MotStop(boardNum, X_axis, 1))
                            End If
                            DoEvents
                        Loop
                        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                        Loop
                        If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
                            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
                                DoEvents        'XW
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

Private Sub StopTimer()
    startButtonTimer.Enabled = False
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    resumeTimer.Enabled = False
    redrawValveOnOff.Enabled = False
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
    PurgePosition.Enabled = False
    homeButton.Enabled = False
    Frame1.Enabled = False
    NeedleMode.Enabled = False
    startButtonTimer.Enabled = False
    pauseButtonTimer.Enabled = False
    abortButtonTimer.Enabled = False
    resumeTimer.Enabled = False
End Sub

Private Sub Enable_Button()
    cmdStartButton.Enabled = True
    cmdStopButton.Enabled = True
    cmdResumeButton.Enabled = True
    abortButton.Enabled = True
    closeButton.Enabled = True
    PurgePosition.Enabled = True
    homeButton.Enabled = True
    Frame1.Enabled = True
    NeedleMode.Enabled = True
    startButtonTimer.Enabled = True
End Sub

Private Function Door_Lock() As Boolean
    Dim Door_Lock_Value As Byte
    
    'Check whether the door is opened or not (ILS)
    'Call P1240MotRdReg(boardNum, U_axis, RR5, Door_Lock_Value)
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, Door_Lock_Value)
    Door_Lock_Value = Door_Lock_Value And &H8
    
    '"&H8" means the door will not be locked.
    If (Door_Lock_Value = &H8) Then
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

    If (LowLevel_Value = 0) Then
       
        MsgBox "The metarial is too low, please refill it."
       
        Yellow_Light = True
        Low_Level = True
    Else
        Yellow_Light = False
        Low_Level = False
    End If
End Function

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
    SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
    SystemMoveHeight = SystemMoveHeight * (-1)
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    Call Sleep(0.2)
    
    'To travel the right way (XW)
    If (SystemMoveHeight > zpos) Then
        doPtp xpos, ypos, SystemMoveHeight, zSystemTravelSpeed
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
        If (Timer < Start) Then
            Start = (86400 - Start)
        End If
        
        'Check for dry run
        If (wetRun.value = True) Then
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
        End If
        DoEvents
    Loop
    
    If (wetRun.value = True) Then
        returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HF7FF
        returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
    
        returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
        ReadValue = ReadValue And &HFEFF
        returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
    End If
    
    Call Sleep(0.3)
'    'Left-needle will be gone up.
'    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
'    ReadValue = ReadValue And &HF7FF
'    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
'    Call Sleep(0.1)
'
    If (purgePartArray = False) Then
        Enable_Button
    End If
    
    purgeButtonTimer.Enabled = True
End Sub

Private Sub Always_Purge_Part_Array()
    
'    If SystemMoveHeight > systemHomeZ Then
'        Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "100")))
'    End If
'
'    Call doPtp(systemHomeX - needleOffsetX, systemHomeY - (needleOffsetY * (-1)), systemHomeZ, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "100")))
'
'    Sleep (0.5)
    Leftslider_go_up
    
    Always_Purge
End Sub

'''''''''''''''''''''''''''''
'   Move to System Height   '
'''''''''''''''''''''''''''''
Private Sub Move_System_Height()
    Dim xpos  As Long, ypos As Long, zpos As Long, upos As Long
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, Lcnt, xpos, ypos, zpos, upos))
    'doPtp xpos, ypos, SystemMoveHeight, CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50"))
    
    'Testing (Because ove function will take longer processing time.
    Call setSpeed(100)
    checkSuccess (P1240MotPtp(boardNum, X_axis Or Y_axis Or Z_axis, X_axis Or Y_axis Or Z_axis, xpos, ypos, SystemMoveHeight, 0))
    
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
        DoEvents
    Loop
End Sub

'''''''''''''''''''''''''''''''''''''''''
'   New procedure for "LinksLinePoint"  '
'''''''''''''''''''''''''''''''''''''''''
Private Function doContiLine3D_XW(endx As Long, endy As Long, endz As Long)
    Position(0) = endx
    Position(1) = endy
    Position(2) = endz
    
    'Save previous position
    'Pt_X = Pt_X + endx
    'Pt_Y = Pt_Y + endy
    'Pt_Z = Pt_Z + endz
End Function

Private Function doSegmentProperty3D_XW(Speed As Long, dispenser As Integer, sequenceNum As Long)
    'checkSuccess (P1240MotAxisParaSet(boardNum, 0, &H7, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate))
    returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(Speed, X_axis), MaxVelocity, AccelSpeed, AccelRate)
    
    
    If (dispenser = 0 Or dispenser = 10 Or dispenser = 20 Or dispenser = 30 Or dispenser = 40 Or estopValue = True Or abortValue = True) Then
        Call doDispenseOff
    Else
        If (dispenser = 1 Or dispenser = 11 Or dispenser = 21 Or dispenser = 31 Or dispenser = 41) Then
            Call doDispenseOn
        End If
    End If
    
'    If (dispenser = False Or estopValue = True Or abortValue = True) Then
'        Call doDispenseOff
'    Else
'        Call doDispenseOn
'    End If
    
    returncode = P1240MotLine(boardNum, X_axis Or Y_axis Or Z_axis, 0, Position(0), Position(1), Position(2), 0)
    Do While ((P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success) And (estopValue = False) And (abortValue = False))
        DoEvents
    Loop
End Function


Private Sub wetRun_Click()
    Call SetLightIntensity(0)
End Sub
