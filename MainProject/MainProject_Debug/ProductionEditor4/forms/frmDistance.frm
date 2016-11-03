VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmDistance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculate the distance between two points"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   6480
   Begin ACTIVESKINLibCtl.SkinLabel CalculateDistance 
      Height          =   285
      Left            =   120
      OleObjectBlob   =   "frmDistance.frx":0000
      TabIndex        =   27
      Top             =   900
      Width           =   1710
   End
   Begin VB.CommandButton Calculate 
      Caption         =   "Teach First Point"
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton zMinus 
      Height          =   615
      Left            =   5160
      Picture         =   "frmDistance.frx":008C
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton zPlus 
      Height          =   615
      Left            =   5160
      Picture         =   "frmDistance.frx":0496
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton xMinus 
      Height          =   615
      Left            =   3360
      Picture         =   "frmDistance.frx":0880
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3000
      Width           =   735
   End
   Begin VB.CommandButton yPlus 
      Height          =   615
      Left            =   2640
      Picture         =   "frmDistance.frx":0C4C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton yMinus 
      Height          =   615
      Left            =   2640
      Picture         =   "frmDistance.frx":1036
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.CommandButton xPlus 
      Height          =   615
      Left            =   1920
      Picture         =   "frmDistance.frx":1440
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox StepDistance 
      Height          =   285
      Left            =   4920
      TabIndex        =   8
      Text            =   "1.000"
      Top             =   1680
      Width           =   1200
   End
   Begin VB.Frame JoggingMode 
      Caption         =   "Jogging Mode"
      Height          =   560
      Left            =   1920
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
      Begin VB.OptionButton Jogging 
         Caption         =   "Jog"
         Height          =   240
         Left            =   1200
         TabIndex        =   7
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton JoggingStep 
         Caption         =   "Step"
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.TextBox Distance 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   2800
   End
   Begin VB.TextBox SecondPoint 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2800
   End
   Begin ACTIVESKINLibCtl.SkinLabel Point2 
      Height          =   285
      Left            =   120
      OleObjectBlob   =   "frmDistance.frx":181B
      TabIndex        =   2
      Top             =   560
      Width           =   1455
   End
   Begin VB.TextBox FirstPoint 
      Height          =   375
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   2800
   End
   Begin ACTIVESKINLibCtl.SkinLabel Point1 
      Height          =   285
      Left            =   120
      OleObjectBlob   =   "frmDistance.frx":189B
      TabIndex        =   0
      Top             =   165
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4560
      OleObjectBlob   =   "frmDistance.frx":1919
      Top             =   3600
   End
   Begin ACTIVESKINLibCtl.SkinLabel LabelDistance 
      Height          =   255
      Left            =   4920
      OleObjectBlob   =   "frmDistance.frx":1B4D
      TabIndex        =   9
      Top             =   1440
      Width           =   1095
   End
   Begin ComCtl2.UpDown UpDownStep 
      Height          =   255
      Left            =   6120
      TabIndex        =   10
      Top             =   1680
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      OrigLeft        =   2520
      OrigTop         =   3120
      OrigRight       =   2775
      OrigBottom      =   3375
      Enabled         =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Index           =   0
      Left            =   2880
      OleObjectBlob   =   "frmDistance.frx":1BC5
      TabIndex        =   15
      Top             =   2160
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Index           =   0
      Left            =   2160
      OleObjectBlob   =   "frmDistance.frx":1C25
      TabIndex        =   16
      Top             =   2760
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Index           =   1
      Left            =   2880
      OleObjectBlob   =   "frmDistance.frx":1C85
      TabIndex        =   17
      Top             =   3360
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Index           =   1
      Left            =   3720
      OleObjectBlob   =   "frmDistance.frx":1CE5
      TabIndex        =   18
      Top             =   2760
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   5400
      OleObjectBlob   =   "frmDistance.frx":1D45
      TabIndex        =   21
      Top             =   3240
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   405
      OleObjectBlob   =   "frmDistance.frx":1DA5
      TabIndex        =   23
      Top             =   4080
      Width           =   300
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   410
      OleObjectBlob   =   "frmDistance.frx":1E09
      TabIndex        =   24
      Top             =   1440
      Width           =   285
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   420
      Left            =   240
      OleObjectBlob   =   "frmDistance.frx":1E6D
      TabIndex        =   25
      Top             =   2640
      Width           =   465
   End
   Begin ComctlLib.Slider jogSpeedSlider 
      Height          =   3015
      Left            =   840
      TabIndex        =   26
      Top             =   1320
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   5318
      _Version        =   327682
      Orientation     =   1
      Min             =   2
      Max             =   101
      SelStart        =   28
      TickStyle       =   2
      Value           =   28
   End
End
Attribute VB_Name = "frmDistance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim First_Pt_X As Double, First_Pt_Y As Double, First_Pt_Z As Double                    'First position
Dim Second_Pt_X As Double, Second_Pt_Y As Double, Second_Pt_Z As Double                 'Second position
Dim OneSevenKey As Integer                                                              'Use in movement with key-board

Private Sub Calculate_Click()
    Dim xValue As Long, yValue As Long, zValue As Long, uValue As Long

    If (Calculate.Caption = "Teach First Point") Or (Calculate.Caption = "Teach Second Point") Then
        checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))

        yValue = yValue * (-1)
        zValue = zValue * (-1)
    ElseIf (Calculate.Caption = "Calculate Distance") Then
        Second_Pt_X = Second_Pt_X - First_Pt_X
        Second_Pt_Y = Second_Pt_Y - First_Pt_Y
        Second_Pt_Z = Second_Pt_Z - First_Pt_Z
        
        Distance.Text = "X=" & Second_Pt_X & ", Y=" & Second_Pt_Y & ", Z=" & Second_Pt_Z
        
        Calculate.Caption = "&Restart"
        Exit Sub
    ElseIf (Calculate.Caption = "&Restart") Then
        FirstPoint.Text = ""
        SecondPoint.Text = ""
        Distance.Text = ""
        
        Calculate.Caption = "Teach First Point"
        Exit Sub
    End If
    
    If (Calculate.Caption = "Teach First Point") Then
        First_Pt_X = convertToMM(xValue, X_axis)
        First_Pt_Y = convertToMM(yValue, Y_axis)
        First_Pt_Z = convertToMM(zValue, Z_axis)
    
        FirstPoint.Text = "X=" & First_Pt_X & ", Y=" & First_Pt_Y & ", Z=" & First_Pt_Z
        
        Calculate.Caption = "Teach Second Point"
    ElseIf (Calculate.Caption = "Teach Second Point") Then
        Second_Pt_X = convertToMM(xValue, X_axis)
        Second_Pt_Y = convertToMM(yValue, Y_axis)
        Second_Pt_Z = convertToMM(zValue, Z_axis)
    
        SecondPoint.Text = "X=" & Second_Pt_X & ", Y=" & Second_Pt_Y & ", Z=" & Second_Pt_Z
        
        Calculate.Caption = "Calculate Distance"
    End If
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    
    UpDownStep.Enabled = False
    StepDistance.Enabled = False
    LabelDistance.Enabled = False
    frmDistance.KeyPreview = True
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
    Call validateNumber(frmDistance.StepDistance.Text, frmDistance.LabelDistance.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        frmDistance.StepDistance.Text = ""
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(frmDistance.StepDistance.Text) <= 0) Then
            frmDistance.StepDistance.Text = "0.001"
        ElseIf (CDbl(frmDistance.StepDistance.Text) >= 10) Then
            frmDistance.StepDistance.Text = "10.000"
        Else
            frmDistance.StepDistance.Text = Format(frmDistance.StepDistance.Text, "#0.000")
        End If
    End If
End Sub

Private Sub UpDownStep_DownClick()
    If (CDbl(StepDistance.Text) <> 0) Then
        StepDistance.Text = CDbl(StepDistance.Text) - 0.001
    End If
End Sub

Private Sub UpDownStep_UpClick()
    StepDistance.Text = CDbl(StepDistance.Text) + 0.001
End Sub

Private Sub xMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_xMinus_mouseDown
End Sub

Private Sub xPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_xPlus_mouseDown
End Sub

Private Sub xMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_xPlusMinus_mouseUp
End Sub

Private Sub xPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_xPlusMinus_mouseUp
End Sub

Private Sub yMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_yMinus_mouseDown
End Sub

Private Sub yPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_yPlus_mouseDown
End Sub

Private Sub yMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_yPlusMinus_mouseUp
End Sub

Private Sub yPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_yPlusMinus_mouseUp
End Sub

Private Sub zMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_zMinus_mouseDown
End Sub

Private Sub zPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_zPlus_mouseDown
End Sub

Private Sub zMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_zPlusMinus_mouseUp
End Sub

Private Sub zPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    Move_zPlusMinus_mouseUp
End Sub

Private Sub jogSpeedSlider_Change()
    setSpeed (jogSpeedSlider.value - 1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      
    If (KeyCode = vbKeyNumlock) Then
        NumLock = True
    End If
        
    If ((KeyCode = vbKeyRight) Or (KeyCode = 102)) And (Indicator = True) Then
        xMinus.SetFocus
       
        Move_xMinus_mouseDown
    ElseIf ((KeyCode = vbKeyLeft) Or (KeyCode = 100)) And (Indicator = True) Then
        xPlus.SetFocus
       
        Move_xPlus_mouseDown
    ElseIf ((KeyCode = vbKeyUp) Or (KeyCode = 104)) And (Indicator = True) Then
        yPlus.SetFocus
        
        Move_yPlus_mouseDown
    ElseIf ((KeyCode = vbKeyDown) Or (KeyCode = 98)) And (Indicator = True) Then
        yMinus.SetFocus
       
        Move_yMinus_mouseDown
    ElseIf ((KeyCode = vbKeyUp) Or (KeyCode = 104)) And (Reflector = True) Then
        zPlus.SetFocus
        
        Move_zPlus_mouseDown
    ElseIf ((KeyCode = vbKeyDown) Or (KeyCode = 98)) And (Reflector = True) Then
        zMinus.SetFocus
        
        Move_zMinus_mouseDown
    End If
    
    'Detect the pressing key at the same time
    If (KeyCode = 97) Then
        If (KeySeven = True) Then
            MsgBox "Don't press these two buttons at the same time."
            Disable
            Exit Sub
        End If
        
        KeyOne = True
    End If
    
    If (KeyCode = 103) Then
        If (KeyOne = True) Then
            MsgBox "Don't press these two buttons at the same time."
            Disable
            Exit Sub
        End If
        
        KeySeven = True
    End If
    
    If Shift = vbShiftMask + vbCtrlMask Then
        MsgBox ("Please don't press the two keys at the same time!")
        Exit Sub
    ElseIf (KeyCode = 17) Or (KeyCode = 97) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success) Then
            Exit Sub
        End If
        Reflector = False
        Indicator = True
        If (OneSevenKey > 0) Then
            xPlus.SetFocus          'Just for testing or change the focus       'XW
        End If
    ElseIf (KeyCode = 16) Or (KeyCode = 103) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success) Then
            Exit Sub
        End If
        Indicator = False
        Reflector = True
        If (OneSevenKey > 0) Then
            zPlus.SetFocus          'Just for testing or change the focus       'XW
        End If
    ElseIf (KeyCode = 18) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success) Then
            Exit Sub
        End If
        MovingMouse = True
    End If
    
    If (KeyCode = 16) Or (KeyCode = 17) Or (KeyCode = 97) Or (KeyCode = 103) Then
        OneSevenKey = OneSevenKey + 1
    End If
End Sub

Private Sub Disable()
    KeyOne = False
    KeySeven = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (NumLock = True) Then
        '"0" if the key is off and "1" if it is on
        If GetKeyState(vbKeyNumlock) = 0 Then
            NumLock = False
            MsgBox ("Please don't disable 'Num Lock' key.")
            
            Exit Sub
        End If
    End If
    
    If (KeyCode = 97) Or (KeyCode = 103) Then
        Disable
        OneSevenKey = 0
    End If
    
    If (Indicator = True) Then
        If (KeyCode = vbKeyRight) Or (KeyCode = 102) Or (KeyCode = vbKeyLeft) Or (KeyCode = 100) Then
            Move_xPlusMinus_mouseUp
            Indicator = True
        ElseIf (KeyCode = vbKeyUp) Or (KeyCode = 104) Or (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
            Move_yPlusMinus_mouseUp
            Indicator = True
        Else
            Move_xPlusMinus_mouseUp
            Move_yPlusMinus_mouseUp
            Indicator = False
        End If
    ElseIf (Reflector = True) Then
        If (KeyCode = vbKeyUp) Or (KeyCode = 104) Or (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
            Move_zPlusMinus_mouseUp
            Reflector = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = 100) Or (KeyCode = 102) Then
            Indicator = False
            Reflector = True
        Else
            Move_zPlusMinus_mouseUp
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
