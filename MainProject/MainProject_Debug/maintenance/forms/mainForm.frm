VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form mainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Desktop Setup Panel"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8370
   Icon            =   "mainForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   633
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer SolventPosTeachTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   2880
   End
   Begin VB.TextBox xSolventPos 
      Height          =   285
      Left            =   5040
      TabIndex        =   96
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox ySolventPos 
      Height          =   285
      Left            =   5040
      TabIndex        =   95
      Top             =   5640
      Width           =   855
   End
   Begin VB.TextBox zSolventPos 
      Height          =   285
      Left            =   5040
      TabIndex        =   94
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton SolventPosTeach 
      Caption         =   "Teach Off"
      Height          =   375
      Left            =   6960
      TabIndex        =   93
      Top             =   5640
      Width           =   1215
   End
   Begin VB.CheckBox chkSolventPos 
      Caption         =   "Always Park"
      Height          =   210
      Left            =   6960
      TabIndex        =   92
      Top             =   6240
      Width           =   1335
   End
   Begin VB.CommandButton cmdSolventPos 
      Caption         =   "Park Position"
      Height          =   375
      Left            =   6960
      TabIndex        =   91
      Top             =   5160
      Width           =   1215
   End
   Begin VB.Frame NeedleMode 
      Caption         =   "Needle Mode"
      Height          =   615
      Left            =   600
      TabIndex        =   88
      Top             =   3360
      Width           =   2535
      Begin VB.OptionButton RightNeedle 
         Caption         =   "Right "
         Height          =   255
         Left            =   1440
         TabIndex        =   90
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton LeftNeedle 
         Caption         =   "Left "
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "mainForm.frx":08CA
      TabIndex        =   87
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton SaveParameters 
      Caption         =   "Save "
      Height          =   375
      Left            =   2160
      TabIndex        =   86
      Top             =   2760
      Width           =   975
   End
   Begin VB.Timer MouseMovement 
      Interval        =   200
      Left            =   0
      Top             =   3360
   End
   Begin VB.Frame JoggingMode 
      Caption         =   "Jogging Mode"
      Height          =   615
      Left            =   600
      TabIndex        =   83
      Top             =   6840
      Width           =   2535
      Begin VB.OptionButton Jogging 
         Caption         =   "Jog"
         Height          =   255
         Left            =   1440
         TabIndex        =   85
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton JoggingStep 
         Caption         =   "Step"
         Height          =   255
         Left            =   240
         TabIndex        =   84
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox StepDistance 
      Height          =   285
      Left            =   1680
      TabIndex        =   82
      Text            =   "1.000"
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Timer SetFocusTimer 
      Interval        =   1000
      Left            =   0
      Top             =   3840
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
      Height          =   255
      Index           =   0
      Left            =   4920
      OleObjectBlob   =   "mainForm.frx":0946
      TabIndex        =   77
      Top             =   6480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
      Height          =   255
      Index           =   0
      Left            =   3600
      OleObjectBlob   =   "mainForm.frx":09A6
      TabIndex        =   75
      Top             =   7560
      Width           =   135
   End
   Begin VB.Timer CheckEmergencyStop 
      Interval        =   100
      Left            =   0
      Top             =   8160
   End
   Begin VB.Timer purgeButtonTimer 
      Interval        =   100
      Left            =   0
      Top             =   7680
   End
   Begin VB.Timer resetTimer 
      Interval        =   1000
      Left            =   0
      Top             =   7200
   End
   Begin VB.Timer SensorTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   0
      Top             =   4320
   End
   Begin ACTIVESKINLibCtl.SkinLabel LimitReachedLabel 
      Height          =   375
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0A06
      TabIndex        =   70
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Timer datumTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   6240
   End
   Begin VB.CommandButton NeedleDatumCmd 
      Caption         =   "Needle Datum Position"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11640
      TabIndex        =   69
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox xDatum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9840
      TabIndex        =   62
      Top             =   3480
      Width           =   855
   End
   Begin VB.TextBox yDatum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9840
      TabIndex        =   61
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox zDatum 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9840
      TabIndex        =   60
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton teachDatum 
      Caption         =   "Teach Off"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11640
      TabIndex        =   58
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Timer DisplayCoOrdsTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   4800
   End
   Begin VB.TextBox zCoOrd 
      Height          =   285
      Left            =   4440
      TabIndex        =   57
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox yCoOrd 
      Height          =   285
      Left            =   3120
      TabIndex        =   56
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox xCoOrd 
      Height          =   285
      Left            =   1800
      TabIndex        =   55
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton PurgePositionCmd 
      Caption         =   "Purge Position"
      Height          =   375
      Left            =   6960
      TabIndex        =   54
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton RobotHomeCmd 
      Caption         =   "Robot Home"
      Height          =   375
      Left            =   6960
      TabIndex        =   53
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton SystemHomeCmd 
      Caption         =   "System Home"
      Height          =   375
      Left            =   6960
      TabIndex        =   52
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Timer updatePurgePosition 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   6720
   End
   Begin VB.Timer updateSystemHomeTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   5760
   End
   Begin VB.Timer updateMoveHeightTimer 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   5280
   End
   Begin VB.CommandButton teachMoveHeight 
      Caption         =   "Teach Off"
      Height          =   375
      Left            =   11040
      TabIndex        =   45
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton systemHomeTeach 
      Caption         =   "Teach Off"
      Height          =   375
      Left            =   6960
      TabIndex        =   44
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton purgeTeach 
      Caption         =   "Teach Off"
      Height          =   375
      Left            =   6960
      TabIndex        =   43
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPurgeButton 
      Caption         =   "Purge"
      Height          =   375
      Left            =   6960
      TabIndex        =   42
      Top             =   3840
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0AC2
      TabIndex        =   41
      Top             =   4200
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0B24
      TabIndex        =   40
      Top             =   3840
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel25 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0B86
      TabIndex        =   39
      Top             =   3480
      Width           =   255
   End
   Begin VB.TextBox zPurgePosition 
      Height          =   285
      Left            =   5040
      TabIndex        =   38
      Top             =   4200
      Width           =   855
   End
   Begin VB.TextBox yPurgePosition 
      Height          =   285
      Left            =   5040
      TabIndex        =   37
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox xPurgePosition 
      Height          =   285
      Left            =   5040
      TabIndex        =   36
      Top             =   3480
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0BE8
      TabIndex        =   35
      Top             =   4200
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel17 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0C4A
      TabIndex        =   34
      Top             =   3840
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0CAC
      TabIndex        =   33
      Top             =   3480
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   4680
      OleObjectBlob   =   "mainForm.frx":0D0E
      TabIndex        =   32
      Top             =   3120
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0D88
      TabIndex        =   31
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel11 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0DEA
      TabIndex        =   30
      Top             =   2160
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":0E4C
      TabIndex        =   29
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox zSystemHome 
      Height          =   285
      Left            =   5040
      TabIndex        =   28
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox ySystemHome 
      Height          =   285
      Left            =   5040
      TabIndex        =   27
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox xSystemHome 
      Height          =   285
      Left            =   5040
      TabIndex        =   26
      Top             =   1800
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0EAE
      TabIndex        =   25
      Top             =   2520
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0F10
      TabIndex        =   24
      Top             =   2160
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":0F72
      TabIndex        =   23
      Top             =   1800
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   255
      Left            =   4680
      OleObjectBlob   =   "mainForm.frx":0FD4
      TabIndex        =   22
      Top             =   1320
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Vision Calibration"
      Height          =   375
      Left            =   1080
      TabIndex        =   21
      Top             =   6000
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "mainForm.frx":1054
      TabIndex        =   20
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox systemMoveHeight 
      Height          =   285
      Left            =   9720
      TabIndex        =   19
      Text            =   "0"
      Top             =   2040
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel23 
      Height          =   375
      Left            =   9360
      OleObjectBlob   =   "mainForm.frx":10B6
      TabIndex        =   18
      Top             =   2040
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel22 
      Height          =   255
      Left            =   9600
      OleObjectBlob   =   "mainForm.frx":1118
      TabIndex        =   17
      Top             =   1560
      Width           =   2175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   375
      Left            =   2760
      OleObjectBlob   =   "mainForm.frx":11AA
      TabIndex        =   16
      Top             =   2160
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   2760
      OleObjectBlob   =   "mainForm.frx":1210
      TabIndex        =   15
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox zDefaultSpeed 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox xyDefaultSpeed 
      Height          =   285
      Left            =   2160
      TabIndex        =   13
      Top             =   1800
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   495
      Left            =   480
      OleObjectBlob   =   "mainForm.frx":1276
      TabIndex        =   12
      Top             =   2160
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   480
      OleObjectBlob   =   "mainForm.frx":12F0
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   1200
      OleObjectBlob   =   "mainForm.frx":1366
      TabIndex        =   10
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton xMinus 
      Height          =   615
      Left            =   5520
      Picture         =   "mainForm.frx":13EA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton yPlus 
      Height          =   615
      Left            =   4680
      Picture         =   "mainForm.frx":17B6
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6720
      Width           =   735
   End
   Begin VB.CommandButton yMinus 
      Height          =   615
      Left            =   4680
      Picture         =   "mainForm.frx":1BA0
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton xPlus 
      Height          =   615
      Left            =   3840
      Picture         =   "mainForm.frx":1FAA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7320
      Width           =   735
   End
   Begin VB.CommandButton zMinus 
      Height          =   615
      Left            =   6960
      Picture         =   "mainForm.frx":2385
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7920
      Width           =   735
   End
   Begin VB.CommandButton zPlus 
      Height          =   615
      Left            =   6960
      Picture         =   "mainForm.frx":278F
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6720
      Width           =   735
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "mainForm.frx":2B79
      Top             =   8760
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   6720
      OleObjectBlob   =   "mainForm.frx":2DAD
      TabIndex        =   0
      Top             =   8640
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "mainForm.frx":2E11
      TabIndex        =   1
      Top             =   8640
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   960
      OleObjectBlob   =   "mainForm.frx":2E75
      TabIndex        =   2
      Top             =   8880
      Width           =   615
   End
   Begin ComctlLib.Slider jogSpeedSlider 
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   8880
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
   Begin VB.Frame Frame1 
      Caption         =   "HOME Travelling Profile"
      Height          =   615
      Left            =   600
      TabIndex        =   46
      Top             =   4200
      Width           =   2535
      Begin VB.OptionButton indirectOption 
         Caption         =   "Indirect"
         Height          =   255
         Left            =   1440
         TabIndex        =   49
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton directOption 
         Caption         =   "Direct"
         Height          =   255
         Left            =   240
         TabIndex        =   48
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Always Robot Home"
      Height          =   615
      Left            =   600
      TabIndex        =   47
      Top             =   5040
      Width           =   2535
      Begin VB.OptionButton alwaysHomeRobotOff 
         Caption         =   "OFF"
         Height          =   255
         Left            =   1440
         TabIndex        =   51
         Top             =   240
         Width           =   615
      End
      Begin VB.OptionButton alwaysHomeRobotOn 
         Caption         =   "ON"
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   240
         Width           =   855
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel13 
      Height          =   255
      Left            =   10800
      OleObjectBlob   =   "mainForm.frx":2EE5
      TabIndex        =   59
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   9360
      OleObjectBlob   =   "mainForm.frx":2F47
      TabIndex        =   63
      Top             =   4200
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel28 
      Height          =   255
      Left            =   9360
      OleObjectBlob   =   "mainForm.frx":2FA9
      TabIndex        =   64
      Top             =   3840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
      Height          =   255
      Left            =   9360
      OleObjectBlob   =   "mainForm.frx":300B
      TabIndex        =   65
      Top             =   3480
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
      Height          =   375
      Left            =   9480
      OleObjectBlob   =   "mainForm.frx":306D
      TabIndex        =   66
      Top             =   3000
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel31 
      Height          =   255
      Left            =   10800
      OleObjectBlob   =   "mainForm.frx":30F5
      TabIndex        =   67
      Top             =   3840
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
      Height          =   255
      Left            =   10800
      OleObjectBlob   =   "mainForm.frx":3157
      TabIndex        =   68
      Top             =   4200
      Width           =   375
   End
   Begin VB.Frame Frame3 
      Caption         =   "Robot Position"
      Height          =   735
      Left            =   1200
      TabIndex        =   71
      Top             =   240
      Width           =   4455
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel35 
         Height          =   255
         Left            =   3120
         OleObjectBlob   =   "mainForm.frx":31B9
         TabIndex        =   74
         Top             =   240
         Width           =   135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
         Height          =   255
         Left            =   1800
         OleObjectBlob   =   "mainForm.frx":3219
         TabIndex        =   73
         Top             =   240
         Width           =   135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "mainForm.frx":3279
         TabIndex        =   72
         Top             =   240
         Width           =   255
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
      Height          =   255
      Index           =   1
      Left            =   6360
      OleObjectBlob   =   "mainForm.frx":32D9
      TabIndex        =   76
      Top             =   7560
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
      Height          =   255
      Index           =   1
      Left            =   4920
      OleObjectBlob   =   "mainForm.frx":3339
      TabIndex        =   78
      Top             =   8520
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
      Height          =   255
      Index           =   1
      Left            =   7320
      OleObjectBlob   =   "mainForm.frx":3399
      TabIndex        =   79
      Top             =   7560
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel LabelDistance 
      Height          =   255
      Left            =   600
      OleObjectBlob   =   "mainForm.frx":33F9
      TabIndex        =   80
      Top             =   7800
      Width           =   1095
   End
   Begin ComCtl2.UpDown UpDownStep 
      Height          =   255
      Left            =   2760
      TabIndex        =   81
      Top             =   7800
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel SolventPosition 
      Height          =   255
      Left            =   4680
      OleObjectBlob   =   "mainForm.frx":3471
      TabIndex        =   97
      Top             =   4920
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":34E9
      TabIndex        =   98
      Top             =   6030
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":354B
      TabIndex        =   99
      Top             =   5700
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
      Height          =   255
      Left            =   4560
      OleObjectBlob   =   "mainForm.frx":35AD
      TabIndex        =   100
      Top             =   5325
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":360F
      TabIndex        =   101
      Top             =   5280
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":3671
      TabIndex        =   102
      Top             =   5640
      Width           =   255
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "mainForm.frx":36D3
      TabIndex        =   103
      Top             =   6000
      Width           =   255
   End
   Begin MSCommLib.MSComm mscomLighIntensity 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "mainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable for mouse's movement
Dim lngReturn As Long, SaveParameter As Long
Dim lngScreenX As Long, lngScreenY As Long
Dim OneSevenKey As Integer, TabNumber As Integer
Const MaxSpeed As Long = 150
Const MaxMouseRangeX As Long = 1023
Const MaxMouseRangeY As Long = 767
Dim SignalCounter As Long           'Count enable signal for purge timer
Dim PrePositionX As Long, PrePositionY As Long, PrePositionZ As Long, PrePositionU   'Save PrePosition
Dim DisablePurgeSignal As Boolean
Dim LeftValve As Boolean            'Just a flag for choosing left-valve.
Dim RightValve As Boolean           'Just a flag for choosing right-valve.
    
Private Sub alwaysHomeRobotOff_Click()
    alwaysHomeRobotOn.value = False
    alwaysHomeRobotOff.value = True
End Sub

Private Sub alwaysHomeRobotOn_Click()
    alwaysHomeRobotOn.value = True
    alwaysHomeRobotOff.value = False
End Sub

Private Sub cmdSolventPos_Click()
    'Left slider go up first before move park position
    Call Leftslider_go_up
    
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
    
    PTPToXYZ convertToPulses(xSolventPos.Text, X_axis), convertToPulses(ySolventPos.Text, Y_axis), convertToPulses(zSolventPos.Text, Z_axis)
    
    'Left slider go down
    Call Leftslider_go_down
End Sub

Private Sub Command3_Click()
    'XW
    'Save them first before we change the vision form
    SaveStringSetting "EpoxyDispenser", "Setup", "xyDefaultSpeed", xyDefaultSpeed.Text
    SaveStringSetting "EpoxyDispenser", "Setup", "zDefaultSpeed", zDefaultSpeed.Text
    CheckEmergencyStop.Enabled = False
    purgeButtonTimer.Enabled = False
    
    Command3.Enabled = False
    
    mainForm.KeyPreview = False
    
    visionCalibration.Show (vbModal)
    
    'Show last statge of sliders
    If (LeftNeedle.value = True) Then
        LeftValve = False
        RightValve = True
        LeftNeedle_Click
    Else
        LeftValve = True
        RightValve = False
        RightNeedle_Click
    End If
                
    mainForm.KeyPreview = True
    Command3.Enabled = True
    SetFocusTimer.Enabled = True
    purgeButtonTimer.Enabled = True
    CheckEmergencyStop.Enabled = True
    
End Sub

Private Sub datumTimer_Timer()

    Dim xValue, yValue, zValue, uValue As Long

    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
        
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
        
    xDatum.Text = convertToMM(xValue, X_axis)
    yDatum.Text = convertToMM(yValue, Y_axis)
    zDatum.Text = convertToMM(zValue, Z_axis)
        
End Sub

Private Sub directOption_Click()
    directOption.value = True
    indirectOption.value = False
End Sub

Private Sub displayCoOrdsTimer_Timer()
    displayCoOrds
    
    'Check and show tower light
    Tower_Light
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
      
    If (KeyCode = vbKeyNumlock) Then
        NumLock = True
    End If
        
    If ((KeyCode = vbKeyRight) Or (KeyCode = 102)) And (Indicator = True) Then
        xMinus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, X_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) + CDbl(StepDistance.Text), X_axis), 0, 0, 0)
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
            Loop
        End If
    ElseIf ((KeyCode = vbKeyLeft) Or (KeyCode = 100)) And (Indicator = True) Then
        xPlus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, X_axis, 1)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) - CDbl(StepDistance.Text), X_axis), 0, 0, 0)
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)  'Loop while X motor is still spinning
            Loop
        End If
    ElseIf ((KeyCode = vbKeyUp) Or (KeyCode = 104)) And (Indicator = True) Then
        yPlus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Y_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) - CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0)
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS) 'Loop while Y motor is still spinning
            Loop
        End If
    ElseIf ((KeyCode = vbKeyDown) Or (KeyCode = 98)) And (Indicator = True) Then
        yMinus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Y_axis, 2)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) + CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0)
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS)  'Loop while Y motor is still spinning
            Loop
        End If
    ElseIf ((KeyCode = vbKeyUp) Or (KeyCode = 104)) And (Reflector = True) Then
        zPlus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Z_axis, 0)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) - CDbl(StepDistance.Text), Z_axis)) * (-1), 0)
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
            Loop
        End If
    ElseIf ((KeyCode = vbKeyDown) Or (KeyCode = 98)) And (Reflector = True) Then
        zMinus.SetFocus
        readyStatus = False
        busyStatus = True
        MouseMovement.Enabled = False
        If Jogging.value = True Then
            Call setSpeed(jogSpeedSlider.value - 1)
            Call P1240MotCmove(boardNum, Z_axis, 4)
        ElseIf JoggingStep.value = True Then
            Call P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) + CDbl(StepDistance.Text), Z_axis)) * (-1), 0)
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
            Loop
        End If
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
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS) Then
            Exit Sub
        End If
        Reflector = False
        Indicator = True
        If (OneSevenKey > 0) Then
            xPlus.SetFocus          'Just for testing or change the focus       'XW
        End If
    ElseIf (KeyCode = 16) Or (KeyCode = 103) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS) Then
            Exit Sub
        End If
        Indicator = False
        Reflector = True
        If (OneSevenKey > 0) Then
            zPlus.SetFocus          'Just for testing or change the focus       'XW
        End If
    ElseIf (KeyCode = 18) Then
        If (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> SUCCESS) Then
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

Private Sub MouseMovement_Timer()
    MouseMovement.Enabled = False
    
    Dim Point As POINTAPI
    Dim angle As Integer
    Dim jogSpeed As Long
    Dim xValue As Long, yValue As Long
    Static StorageValueX, StorageValueY As Long
    Static referencePointX As Long, referencePointY As Long
    
    If MovingMouse = True Then
        
        Call MousePosition(Point)
        
        If (StartingPoint <> 0) Then
            If (referencePointX = lngScreenX) And (referencePointY = lngScreenY) Then
                busyStatus = False
                readyStatus = True
                checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis, 1 Or 2))
                Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
                Loop
                MovingMouse = False
                StartingPoint = 0
                referencePointX = lngScreenX
                referencePointY = lngScreenY
                MouseMovement.Enabled = True
                Exit Sub
            Else
                readyStatus = False
                busyStatus = True
            
                Call setSpeed(jogSpeedSlider.value - 1)
                
                If (referencePointX <> lngScreenX) And (referencePointY <> lngScreenY) Then
                    angle = MouseAngle(referencePointX, lngScreenX, referencePointY, lngScreenY)
                ElseIf (referencePointX = lngScreenX) Then
                    ' Just for indication
                    angle = 75
                ElseIf (referencePointY = lngScreenY) Then
                    angle = 1
                End If
                                
                If (RightDirection = True) Then
                    
                    If (angle <= 15) Then
                        If (referencePointX > lngScreenX) Then
                            XaxisStop
                        End If
                    Else
                        XaxisStop
                    End If
                    
                    RightDirection = False
                ElseIf (LeftDirection = True) Then
                                       
                    If (angle <= 15) Then
                        If (referencePointX < lngScreenX) Then
                            XaxisStop
                        End If
                    Else
                        XaxisStop
                    End If
                    
                    LeftDirection = False
                ElseIf (UpDirection = True) Then
                                      
                    If (angle >= 75) Then
                        If (referencePointY < lngScreenY) Then
                            YaxisStop
                        End If
                    Else
                        YaxisStop
                    End If
                    
                    UpDirection = False
                ElseIf (DownDirection = True) Then
                                       
                    If (angle >= 75) Then
                        If (referencePointY > lngScreenY) Then
                            YaxisStop
                        End If
                    Else
                        YaxisStop
                    End If
                    
                    DownDirection = False
                ElseIf (UpRightDirection = True) Then
                    
                    If (angle > 15) And (angle < 75) Then
                        If Not ((referencePointX < lngScreenX) And (referencePointY > lngScreenY)) Then
                            XYaxisStop
                        End If
                    Else
                        XYaxisStop
                    End If
                    
                    UpRightDirection = False
                ElseIf (UpLeftDirection = True) Then
                    
                    If (angle > 15) And (angle < 75) Then
                        If Not ((referencePointX > lngScreenX) And (referencePointY > lngScreenY)) Then
                            XYaxisStop
                        End If
                    Else
                        XYaxisStop
                    End If
                    
                    UpLeftDirection = False
                ElseIf (DownRightDirection = True) Then
                    
                    If (angle > 15) And (angle < 75) Then
                        If Not ((referencePointX < lngScreenX) And (referencePointY < lngScreenY)) Then
                            XYaxisStop
                        End If
                    Else
                        XYaxisStop
                    End If
                    
                    DownRightDirection = False
                ElseIf (DownLeftDirection = True) Then
                    
                    If (angle > 15) And (angle < 75) Then
                        If Not ((referencePointX > lngScreenX) And (referencePointY < lngScreenY)) Then
                            XYaxisStop
                        End If
                    Else
                        XYaxisStop
                    End If
                    
                    DownLeftDirection = False
                End If
            
                If (angle <= 15) Then
                    If (referencePointX < lngScreenX) Then
                        'If (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS) Then
                        '    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
                        '    Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS)
                        '    Loop
                        'End If
                        RightDirection = True
                        Call P1240MotCmove(boardNum, X_axis, 0)
                        'Do While (P1240MotCmove(boardNum, X_axis, 0))
                        '    DoEvents
                            Call MousePosition(Point)
                            If (referencePointX > lngScreenX) Then
                                XaxisStop
                        '        Exit Do
                            End If
                        '    referencePointX = lngScreenX
                        '    referencePointY = lngScreenY
                        'Loop
                    Else
                    'If (referencePointX > lngScreenX) Then
                        LeftDirection = True
                        Call P1240MotCmove(boardNum, X_axis, 1)
                           
                        Call MousePosition(Point)
                        If (referencePointX < lngScreenX) Then
                            XaxisStop
                        End If
                    End If
                ElseIf (angle >= 75) Then
                    If (referencePointY < lngScreenY) Then
                        DownDirection = True
                        Call P1240MotCmove(boardNum, Y_axis, 2)
                        Call MousePosition(Point)
                        If (referencePointY > lngScreenY) Then
                            YaxisStop
                        End If
                    Else
                    'If (referencePointY > lngScreenY) Then
                        UpDirection = True
                        Call P1240MotCmove(boardNum, Y_axis, 0)
                        Call MousePosition(Point)
                        If (referencePointY < lngScreenY) Then
                            YaxisStop
                        End If
                    End If
                Else
                    If (referencePointX > lngScreenX) And (referencePointY > lngScreenY) Then
                        UpLeftDirection = True
                        Call P1240MotCmove(boardNum, X_axis Or Y_axis, 1 Or 0)
                        Call MousePosition(Point)
                        If ((referencePointX < lngScreenX) And (referencePointY < lngScreenY)) Then
                            XYaxisStop
                        End If
                    ElseIf (referencePointX < lngScreenX) And (referencePointY < lngScreenY) Then
                        DownRightDirection = True
                        Call P1240MotCmove(boardNum, X_axis Or Y_axis, 0 Or 2)
                        Call MousePosition(Point)
                        If ((referencePointX > lngScreenX) And (referencePointY > lngScreenY)) Then
                            XYaxisStop
                        End If
                    ElseIf (referencePointX > lngScreenX) And (referencePointY < lngScreenY) Then
                        DownLeftDirection = True
                        Call P1240MotCmove(boardNum, X_axis Or Y_axis, 1 Or 2)
                        Call MousePosition(Point)
                        If ((referencePointX < lngScreenX) And (referencePointY > lngScreenY)) Then
                            XYaxisStop
                        End If
                    ElseIf (referencePointX < lngScreenX) And (referencePointY > lngScreenY) Then
                        UpRightDirection = True
                        Call P1240MotCmove(boardNum, X_axis Or Y_axis, 0 Or 0)
                        Call MousePosition(Point)
                        If ((referencePointX > lngScreenX) And (referencePointY < lngScreenY)) Then
                            XYaxisStop
                        End If
                    End If
                End If
            End If
        End If
        
        StartingPoint = StartingPoint + 1
        
        If (lngScreenX = 0) Then
            lngScreenX = 1023
            StorageValueX = referencePointX
            lngReturn = SetCursorPos(lngScreenX, lngScreenY)
        ElseIf (lngScreenX = 1023) Then
            lngScreenX = 0
            StorageValueX = 1023 - referencePointX
            lngReturn = SetCursorPos(lngScreenX, lngScreenY)
        End If
                
        If (lngScreenY = 0) Then
            lngScreenY = 767
            StorageValueY = referencePointY
            lngReturn = SetCursorPos(lngScreenX, lngScreenY)
        ElseIf (lngScreenY = 767) Then
            lngScreenY = 0
            StorageValueY = 767 - referencePointY
            lngReturn = SetCursorPos(lngScreenX, lngScreenY)
        End If
                    
        referencePointX = lngScreenX
        referencePointY = lngScreenY
    End If
    
    MouseMovement.Enabled = True
End Sub

Private Sub MousePosition(Point As POINTAPI)
    lngReturn = GetCursorPos(Point)
    lngScreenX = Point.X
    lngScreenY = Point.Y
End Sub

Private Function MouseAngle(ByVal x1 As Integer, ByVal x2 As Integer, ByVal y1 As Integer, ByVal y2 As Integer) As Integer
    Dim xValue, yValue As Integer
    Dim NewValue As Double
    
    xValue = x2 - x1
    yValue = y2 - y1
    
    If (xValue < 0) Then
        xValue = xValue * (-1)
    End If
    
    If (yValue < 0) Then
        yValue = yValue * (-1)
    End If
    
    NewValue = Atn(CDbl(yValue / xValue))
    
    MouseAngle = NewValue * 180 / (Atn(1) * 4)
End Function

Private Sub XaxisStop()
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    Do While (P1240MotAxisBusy(boardNum, X_axis) <> SUCCESS)
    Loop
End Sub

Private Sub YaxisStop()
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    Do While (P1240MotAxisBusy(boardNum, Y_axis) <> SUCCESS)
    Loop
End Sub

Private Sub XYaxisStop()
    checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis, 1 Or 2))
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
    Loop
End Sub

Private Sub SaveParameters_Click()
    SaveParameter = 0
    If (xyDefaultSpeed.Text <> "") And (zDefaultSpeed.Text <> "") And (systemMoveHeight.Text <> "") _
        And (xSystemHome.Text <> "") And (ySystemHome.Text <> "") And (zSystemHome.Text <> "") _
        And (xPurgePosition.Text <> "") And (yPurgePosition.Text <> "") And (zPurgePosition.Text <> "") _
        And (xDatum.Text <> "") And (yDatum.Text <> "") And (zDatum.Text <> "") Then
        
        SaveParameterSetting
        
        SaveStringSetting "EpoxyDispenser", "Setup", "xyDefaultSpeed", xyDefaultSpeed.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "zDefaultSpeed", zDefaultSpeed.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "SystemMoveHeight", systemMoveHeight.Text
            
        SaveStringSetting "EpoxyDispenser", "Setup", "xSystemHome", xSystemHome.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "ySystemHome", ySystemHome.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "zSystemHome", zSystemHome.Text
            
        SaveStringSetting "EpoxyDispenser", "Setup", "xPurgePosition", xPurgePosition.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "yPurgePosition", yPurgePosition.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "zPurgePosition", zPurgePosition.Text
    
        SaveStringSetting "EpoxyDispenser", "Setup", "xDatum", xDatum.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "yDatum", yDatum.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "zDatum", zDatum.Text
        
        SaveStringSetting "EpoxyDispenser", "Setup", "xSolventPos", xSolventPos.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "ySolventPos", ySolventPos.Text
        SaveStringSetting "EpoxyDispenser", "Setup", "zSolventPos", zSolventPos.Text
        
    Else
        MsgBox "Please check the datas whether some of them are blank or not!"
        Exit Sub
    End If
    SaveParameter = SaveParameter + 1
End Sub

Private Sub SetFocusTimer_Timer()
    SetFocusTimer.Enabled = False
    
    Dim xValue, yValue, zValue, uValue As Long
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, xValue, yValue, zValue, uValue))
    xValue = xValue And &HC
    yValue = yValue And &HC
    zValue = zValue And &HC
        
    If (xValue <> 0) Or (yValue <> 0) Or (zValue <> 0) Then
        busyStatus = False
        readyStatus = True
        Me.SetFocus
    End If
    SetFocusTimer.Enabled = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (NumLock = True) Then
        '"0" if the key is off and "1" if it is on
        If GetKeyState(vbKeyNumlock) = 0 Then
            NumLock = False
            MsgBox ("Please don't disable 'Num Lock' key.")
            'MsgBox ("Please lock number key first to teach the robot's position.")
            Exit Sub
        End If
    End If
    
    If (KeyCode = 97) Or (KeyCode = 103) Then
    'If (KeyCode = 35) Or (KeyCode = 36) Then
        Disable
        OneSevenKey = 0
    End If
    
    If (Indicator = True) Then
        If (KeyCode = vbKeyRight) Or (KeyCode = 102) Then
        'If (KeyCode = vbKeyRight) Or (KeyCode = 39) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = 100) Then
        'ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = 37) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf (KeyCode = vbKeyUp) Or (KeyCode = 104) Then
        'ElseIf (KeyCode = vbKeyUp) Or (KeyCode = 38) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = True
        ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
        'ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 40) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = True
        Else
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = False
        End If
    ElseIf (Reflector = True) Then
        If (KeyCode = vbKeyUp) Or (KeyCode = 104) Then
        'If (KeyCode = vbKeyUp) Or (KeyCode = 38) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
        'ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 40) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = 100) Or (KeyCode = 102) Then
        'ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = 37) Or (KeyCode = 39) Then
            Indicator = False
            Reflector = True
        Else
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = False
        End If
    ElseIf (MovingMouse = True) Then
        busyStatus = False
        readyStatus = True
        checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis, 1 Or 2))
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
        Loop
        MovingMouse = False
        StartingPoint = 0
    Else
        busyStatus = False
        readyStatus = True
        checkSuccess (P1240MotStop(boardNum, X_axis, 1))
        checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
        checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
        Indicator = False
        Reflector = False
    End If
    MouseMovement.Enabled = True
End Sub

Private Sub Form_Load()

    Dim hWnd, hWnd2 As Long

    hWnd = FindWindow(vbNullString, "Profile Editor")
    hWnd2 = FindWindow(vbNullString, "File Load")
    
    'SensorTimer.Enabled = False
    resetTimer.Enabled = False
    SetFocusTimer.Enabled = False
    CheckEmergencyStop.Enabled = False
    
    If App.PrevInstance Or hWnd <> 0 Or hWnd2 <> 0 Then
        MsgBox ("Another conflicting process has been detected! This process will abort")
        Unload Me
    Else
        
        purgeButtonTimer.Enabled = False
        'Skin1.LoadSkin (".\skin\epoxySkin.skn")
        Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn") 'for login (NNO)
        Skin1.ApplySkin Me.hWnd
    
        'NNO
        loginsuccessful = False
        mCancel = False
        confirmreset = False
        frmlogin.Show vbModal
        If loginsuccessful = False Or mCancel = True Or confirmreset = True Then
            Unload Me
        Else
        
        'If Get_User_Name <> "Techno Digm" Then
        '    AccessDenied.Show vbModal
        '    Unload Me
        'Else
        
            'Tower_Light
            Red_Light = False
            Yellow_Light = False
            Green_Light = True
            
            readRegistryOptions
            determineProfile
            
            If (InitializePCI1750 = False) Then
                MsgBox "IO card cannot be opened. Please check the card."
                
                End
                Exit Sub
            End If
            
            If (initializeBoard = False) Then
                Close_PCI1750
                
                MsgBox "Motion card has some problems. Please check the card."
                
                End
                Exit Sub
            End If
            
            displayCoOrdsTimer.Enabled = True
            
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
                
                Servo_Off
                Close_TowerLight
                ResetDriver
                unInitializeBoard
                Close_PCI1750
                End
                Exit Sub
            End If
            
            SetFocusTimer.Enabled = False
            purgeButtonTimer.Enabled = False
            
            xyDefaultSpeed.Text = GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")
            zDefaultSpeed.Text = GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "100")
            systemMoveHeight.Text = GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0")
    
            xSystemHome.Text = GetStringSetting("EpoxyDispenser", "Setup", "xSystemHome", "0")
            ySystemHome.Text = GetStringSetting("EpoxyDispenser", "Setup", "ySystemHome", "0")
            zSystemHome.Text = GetStringSetting("EpoxyDispenser", "Setup", "zSystemHome", "0")
    
            xPurgePosition.Text = GetStringSetting("EpoxyDispenser", "Setup", "xPurgePosition", "0")
            yPurgePosition.Text = GetStringSetting("EpoxyDispenser", "Setup", "yPurgePosition", "0")
            zPurgePosition.Text = GetStringSetting("EpoxyDispenser", "Setup", "zPurgePosition", "0")
            
            xSolventPos.Text = GetStringSetting("EpoxyDispenser", "Setup", "xSolventPos", "0")
            ySolventPos.Text = GetStringSetting("EpoxyDispenser", "Setup", "ySolventPos", "0")
            zSolventPos.Text = GetStringSetting("EpoxyDispenser", "Setup", "zSolventPos", "0")
            
            If GetStringSetting("EpoxyDispenser", "Setup", "DirectSoftHome", "0") = "1" Then
                directOption.value = True
                indirectOption.value = False
            Else
                directOption.value = False
                indirectOption.value = True
            End If
            
            needleOffsetX = convertToPulses(GetStringSetting("EpoxyDispenser", "NeedleOffset", "XOff", "0"), X_axis)
            needleOffsetY = convertToPulses(GetStringSetting("EpoxyDispenser", "NeedleOffset", "YOff", "0"), Y_axis)
            needleOffsetY = needleOffsetY * (-1)        'XW
            
            xDatum.Text = GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0")
            yDatum.Text = GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0")
            zDatum.Text = GetStringSetting("EpoxyDispenser", "Setup", "zDatum", "0")
  
            If GetStringSetting("EpoxyDispenser", "Setup", "AlwaysRobotHome", "0") = "1" Then
                alwaysHomeRobotOn.value = True
                alwaysHomeRobotOff.value = False
            Else
                alwaysHomeRobotOn.value = False
                alwaysHomeRobotOff.value = True
            End If
            
            If GetStringSetting("EpoxyDispenser", "Setup", "EnableSolventPosition", "0") = "1" Then
                chkSolventPos.value = 1
            Else
                chkSolventPos.value = 0
            End If
            
            'Set the value for some parameters  'XW
            If (xyDefaultSpeed.Text = "") Or (zDefaultSpeed.Text = "") Or (systemMoveHeight.Text = "") _
                Or (xSystemHome.Text = "") Or (ySystemHome.Text = "") Or (zSystemHome.Text = "") _
                Or (xPurgePosition.Text = "") Or (yPurgePosition.Text = "") Or (zPurgePosition.Text = "") _
                Or (xDatum.Text = "") Or (yDatum.Text = "") Or (zDatum.Text = "") Then
                
                MsgBox ("System will set or generate the previous value because of the system error!")
                                
                ReadParameters
            End If
            
            SaveParameterSetting
            
            Jogging.value = True                    'XW
            JoggingStep.value = False               'XW
            UpDownStep.Enabled = False              'XW
            StepDistance.Enabled = False            'XW
            LabelDistance.Enabled = False           'XW
            ErrorKeyIn = False
            LeftNeedleValve                         'Set Default as Let_Cylinder
            LeftValve = True                        'Just a default flag
            'SensorTimer.Enabled = True
            resetTimer.Enabled = True
            SetFocusTimer.Enabled = True
            CheckEmergencyStop.Enabled = True
            purgeButtonTimer.Enabled = True
            
            SetWindowOnTop Me, True    '@$K
        End If
    End If
End Sub

Private Sub Form_Unload(cancel As Integer)

    Dim hWnd, hWnd2 As Long
    Dim ReadValue As Long, DriverXYZ As Long, Tower_Light_Value As Long
    Dim line, path, directory As String
    Dim StrLenght As Integer

    hWnd = FindWindow(vbNullString, "Profile Editor")
    hWnd2 = FindWindow(vbNullString, "File Load")

    If App.PrevInstance <> True And hWnd = 0 And hWnd2 = 0 Then
    
        'If Get_User_Name = "Techno Digm" Then
        If loginsuccessful = False Or mCancel = True Then 'NNO
            Exit Sub
        Else
    
            If (chkSolventPos.value = 1) Then
                'Move to Solving Position
                cmdSolventPos_Click
            Else
                'Move to System Home Position
                SystemHomeCmd_Click
            End If
            
            Close_AllTimer
            
            Servo_Off
            
            'Remove this part because of the customer's spect.
            'Right-needle will be gone up.
            'Call P1240MotRdReg(boardNum, U_axis, WR3, ReadValue)
            'ReadValue = ReadValue And &HFEFF
            'Call P1240MotWrReg(boardNum, U_axis, WR3, ReadValue)
    
            'Left-needle will be gone up.
            'Call P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            'ReadValue = ReadValue And &HF7FF
            'Call P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
            
            'Disable Red_Light,Yellow_light and Green_Light
            checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
            Tower_Light_Value = Tower_Light_Value And &HF1FF
            checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
            
            Sleep (0.6)
            unInitializeBoard
            Close_PCI1750
            
            If (SaveParameter = 0) Then
                If (xyDefaultSpeed.Text <> "") And (zDefaultSpeed.Text <> "") And (systemMoveHeight.Text <> "") _
                    And (xSystemHome.Text <> "") And (ySystemHome.Text <> "") And (zSystemHome.Text <> "") _
                    And (xPurgePosition.Text <> "") And (yPurgePosition.Text <> "") And (zPurgePosition.Text <> "") _
                    And (xDatum.Text <> "") And (yDatum.Text <> "") And (zDatum.Text <> "") Then
        
                    SaveParameterSetting
        
                    SaveStringSetting "EpoxyDispenser", "Setup", "xyDefaultSpeed", xyDefaultSpeed.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "zDefaultSpeed", zDefaultSpeed.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "SystemMoveHeight", systemMoveHeight.Text
            
                    SaveStringSetting "EpoxyDispenser", "Setup", "xSystemHome", xSystemHome.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "ySystemHome", ySystemHome.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "zSystemHome", zSystemHome.Text
            
                    SaveStringSetting "EpoxyDispenser", "Setup", "xPurgePosition", xPurgePosition.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "yPurgePosition", yPurgePosition.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "zPurgePosition", zPurgePosition.Text
    
                    SaveStringSetting "EpoxyDispenser", "Setup", "xDatum", xDatum.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "yDatum", yDatum.Text
                    SaveStringSetting "EpoxyDispenser", "Setup", "zDatum", zDatum.Text
                End If
            End If
    
            If directOption.value = True Then
                SaveStringSetting "EpoxyDispenser", "Setup", "DirectSoftHome", "1"
            Else
                SaveStringSetting "EpoxyDispenser", "Setup", "DirectSoftHome", "0"
            End If
    
            If alwaysHomeRobotOn.value = True Then
                SaveStringSetting "EpoxyDispenser", "Setup", "AlwaysRobotHome", "1"
            Else
                SaveStringSetting "EpoxyDispenser", "Setup", "AlwaysRobotHome", "0"
            End If
            
            If chkSolventPos.value = 1 Then
                SaveStringSetting "EpoxyDispenser", "Setup", "EnableSolventPosition", "1"
            Else
                SaveStringSetting "EpoxyDispenser", "Setup", "EnableSolventPosition", "0"
            End If
            
            'Copy the "VisionSetup.txt" File to another two folder before closing
            'the window     'XW
            
            line = ""
            path = ""
            directory = ""
            StrLenght = 0
            
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set A = fs.OpenTextFile(App.path & "\VisionSetup.txt", 1, False)
            
            path = CStr(App.path)
            StrLenght = Len(path)
            directory = Left(path, StrLenght - 12)
    
            Set fss = CreateObject("Scripting.FileSystemObject")
            Set aa = fss.createtextfile(directory & "\ProductionEditor4\VisionSetup.txt", True)
            
            Set fsss = CreateObject("Scripting.FileSystemObject")
            Set aaa = fss.createtextfile(directory & "\ProductionRunEngine4\VisionSetup.txt", True)
    
            Do While A.AtEndOfStream <> True
                line = A.ReadLine
                aa.writeline (line)
                aaa.writeline (line)
            Loop
            A.Close
            aa.Close
            aaa.Close
            
            SetWindowOnTop Me, False    '@$K
        End If
    End If
End Sub

Private Sub LeftNeedle_Click()
    Move_To_Zero
    
    'Do it when changing from the right valve.
    If (LeftValve = False) And (RightValve = True) Then
        LeftNeedleValve
    End If
    
    LeftValve = True
    RightValve = False
End Sub

Private Sub RightNeedle_Click()
    Move_To_Zero
    
    'Do it when changing from the left valve.
    If (LeftValve = True) And (RightValve = False) Then
        RightNeedleValve
    End If
    
    LeftValve = False
    RightValve = True
End Sub

Private Sub Move_To_Zero()
    jogSpeedSlider.value = 28
    setSpeed (28)
        
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS)
        DoEvents
    Loop
End Sub

Private Sub inDirectOption_Click()
    directOption.value = False
    indirectOption.value = True
End Sub

Private Sub NeedleDatumCmd_Click()
    purgeButtonTimer.Enabled = False
    systemTrackMoveHeight = convertToPulses(CLng(mainForm.systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(mainForm.xyDefaultSpeed.Text))
    PTPToXYZ convertToPulses(xDatum.Text, X_axis), convertToPulses(yDatum.Text, Y_axis), convertToPulses(zDatum.Text, Z_axis)
    purgeButtonTimer.Enabled = True
End Sub

Private Sub cmdPurgeButton_Mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Original (NYP)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H0)
        
    DisablePurgeSignal = False
End Sub

Private Sub cmdPurgeButton_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Original (NYP)
    'returncode = P1240MotWrReg(boardNum, Z_axis, WR3, &H100)
    
    purgeButtonTimer.Enabled = False
    DisablePurgeSignal = True
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
    
    'To get the actual direcion
    PrePositionY = PrePositionY * (-1)
    PrePositionZ = PrePositionZ * (-1)
               
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
    PTPToXYZ convertToPulses(xPurgePosition.Text, X_axis), convertToPulses(yPurgePosition.Text, Y_axis), convertToPulses(zPurgePosition.Text, Z_axis)
    
    Dispensing
   
End Sub

Private Sub Dispensing()
    Dim ReadValue As Long
    
    Do While (DisablePurgeSignal = True)
        If (LeftNeedle.value = True) Then
            returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H800
            returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
        ElseIf (RightNeedle.value = True) Then
            returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
            ReadValue = ReadValue Or &H100
            returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
        End If
        
        DoEvents
    
    Loop
    
    returncode = P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HF7FF
    returncode = P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue)
    
    returncode = P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue)
    ReadValue = ReadValue And &HFEFF
    returncode = P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue)
    
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
    PTPToXYZ PrePositionX, PrePositionY, PrePositionZ
    
    purgeButtonTimer.Enabled = True
End Sub

Private Sub PurgePositionCmd_Click()
    'Left slider go up first before move park position
    Call Leftslider_go_up
    
    purgeButtonTimer.Enabled = False
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
    PTPToXYZ convertToPulses(xPurgePosition.Text, X_axis), convertToPulses(yPurgePosition.Text, Y_axis), convertToPulses(zPurgePosition.Text, Z_axis)
    purgeButtonTimer.Enabled = True
    
    'Left slider go down
    Call Leftslider_go_down
End Sub

Private Sub purgeTeach_Click()
    If purgeTeach.Caption = "Teach Off" Then
        teachMoveHeight.Caption = "Teach Off"
        systemHomeTeach.Caption = "Teach Off"
        purgeTeach.Caption = "Teach On"
        teachDatum.Caption = "Teach Off"
        SolventPosTeach.Caption = "Teach Off"

        teachMoveHeight.Refresh
        systemHomeTeach.Refresh
        teachDatum.Refresh
        SolventPosTeach.Refresh
        
        updateMoveHeightTimer.Enabled = False
        updateSystemHomeTimer.Enabled = False
        updatePurgePosition.Enabled = True
        datumTimer.Enabled = False
        SolventPosTeachTimer.Enabled = False
    Else
        purgeTeach.Caption = "Teach Off"
        updatePurgePosition.Enabled = False
    End If
End Sub

Private Sub RobotHomeCmd_Click()
    'Think about for e-stop timer???
    purgeButtonTimer.Enabled = False
    moveToHome
    purgeButtonTimer.Enabled = True
End Sub

Private Sub resetTimer_Timer()
    Dim resetX, resetY, resetZ As Long
    
    resetTimer.Enabled = False
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, resetX, resetY, resetZ, 0))
    resetX = resetX And &H10
    resetY = resetY And &H10
    resetZ = resetZ And &H10
    If (resetX = &H10) Or (resetY = &H10) Or (resetZ = &H10) Then
        MsgBox "Driver Error. Please check the Driver!"
        
        Close_AllTimer
            
        Servo_Off
            
        Close_TowerLight
        
        If (loadVisionCalibration = True) Then
            With visionCalibration
                .picImage.Enabled = False
    
                .setfocusTimer2.Enabled = False
                .displayCoOrdsTimer.Enabled = False
        
                Call Sleep(0.6)
        
                VdeCameraLive False
                VdeReleaseVision
            End With
        End If
        
        ResetDriver
        
        unInitializeBoard
        Close_PCI1750
        End
        Exit Sub
    End If
    resetTimer.Enabled = True
End Sub

Private Sub SolventPosTeach_Click()
    If SolventPosTeach.Caption = "Teach Off" Then
    
        teachMoveHeight.Caption = "Teach Off"
        systemHomeTeach.Caption = "Teach Off"
        purgeTeach.Caption = "Teach Off"
        teachDatum.Caption = "Teach Off"
        SolventPosTeach.Caption = "Teach On"
        
        teachMoveHeight.Refresh
        systemHomeTeach.Refresh
        purgeTeach.Refresh
        teachDatum.Refresh
                
        updateMoveHeightTimer.Enabled = False
        updateSystemHomeTimer.Enabled = False
        updatePurgePosition.Enabled = False
        datumTimer.Enabled = False
        SolventPosTeachTimer.Enabled = True
    Else
        SolventPosTeach.Caption = "Teach Off"
        SolventPosTeachTimer.Enabled = False
    End If
End Sub

Private Sub SolventPosTeachTimer_Timer()
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
        
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
        
    xSolventPos.Text = convertToMM(xValue, X_axis)
    ySolventPos.Text = convertToMM(yValue, Y_axis)
    zSolventPos.Text = convertToMM(zValue, Z_axis)
End Sub

'''''''''''''''''''''
'   Level Sensing   '
'''''''''''''''''''''

'Remove it first because our system is not enough IO (For Spray_Two_Head)
'Private Sub SensorTimer_Timer()
'    Dim sensorValue As Long
'    Dim a As Long
'
'    SensorTimer.Enabled = False
'    Call P1240MotRdReg(boardNum, U_axis, RR5, sensorValue)
'    sensorValue = sensorValue And &H200
    
'    Call P1240MotRdReg(boardNum, X_axis, WR3, a)
     
'    'If (sensorValue = &H200) Then
'    If (sensorValue = 0) Then
'        If a < &H800 Then
'            a = a Or &H800
'            Call P1240MotWrReg(boardNum, X_axis, WR3, a)
'        End If
'        frmAlarm.Show (vbModal)
'    Else
'        Call P1240MotWrReg(boardNum, X_axis, WR3, a)
'    End If
'    SensorTimer.Enabled = True
'End Sub

Private Sub SystemHomeCmd_Click()
    'Left slider go up first before move park position
    Call Leftslider_go_up
    
    purgeButtonTimer.Enabled = False
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
    PTPToXYZ convertToPulses(xSystemHome.Text, X_axis), convertToPulses(ySystemHome.Text, Y_axis), convertToPulses(zSystemHome.Text, Z_axis)
    purgeButtonTimer.Enabled = True
    
    'Left slider go down
    Call Leftslider_go_down
End Sub

Private Sub systemHomeTeach_Click()
    If systemHomeTeach.Caption = "Teach Off" Then
        teachMoveHeight.Caption = "Teach Off"
        systemHomeTeach.Caption = "Teach On"
        purgeTeach.Caption = "Teach Off"
        teachDatum.Caption = "Teach Off"
        SolventPosTeach.Caption = "Teach Off"
        
        teachMoveHeight.Refresh
        purgeTeach.Refresh
        teachDatum.Refresh
        SolventPosTeach.Refresh
        
        updateMoveHeightTimer.Enabled = False
        updateSystemHomeTimer.Enabled = True
        updatePurgePosition.Enabled = False
        datumTimer.Enabled = False
        SolventPosTeach.Enabled = False
    Else
        systemHomeTeach.Caption = "Teach Off"
        updateSystemHomeTimer.Enabled = False
    End If
End Sub

Private Sub teachDatum_Click()
    If teachDatum.Caption = "Teach Off" Then
        teachMoveHeight.Caption = "Teach Off"
        systemHomeTeach.Caption = "Teach Off"
        purgeTeach.Caption = "Teach Off"
        teachDatum.Caption = "Teach On"
        teachMoveHeight.Refresh
        purgeTeach.Refresh
        systemHomeTeach.Refresh
        updateMoveHeightTimer.Enabled = False
        updateSystemHomeTimer.Enabled = False
        updatePurgePosition.Enabled = False
        datumTimer.Enabled = True
    Else
        teachDatum.Caption = "Teach Off"
        datumTimer.Enabled = False
    End If
End Sub

Private Sub teachMoveHeight_Click()
    If teachMoveHeight.Caption = "Teach Off" Then
        teachMoveHeight.Caption = "Teach On"
        systemHomeTeach.Caption = "Teach Off"
        purgeTeach.Caption = "Teach Off"
        teachDatum.Caption = "Teach Off"
        systemHomeTeach.Refresh
        purgeTeach.Refresh
        teachDatum.Refresh
        updateMoveHeightTimer.Enabled = True
        updateSystemHomeTimer.Enabled = False
        updatePurgePosition.Enabled = False
        datumTimer.Enabled = False
    Else
        teachMoveHeight.Caption = "Teach Off"
        updateMoveHeightTimer.Enabled = False
    End If
End Sub

Private Sub updateMoveHeightTimer_Timer()
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
    systemMoveHeight.Text = convertToMM(zValue, Z_axis)
    
    If (CDbl(systemMoveHeight.Text) < 0) Then
        systemMoveHeight.Text = CDbl(systemMoveHeight.Text) * (-1)
    End If
End Sub

Private Sub updatePurgePosition_Timer()
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
        
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
        
    xPurgePosition.Text = convertToMM(xValue, X_axis)
    yPurgePosition.Text = convertToMM(yValue, Y_axis)
    zPurgePosition.Text = convertToMM(zValue, Z_axis)
  
End Sub

Private Sub updateSystemHomeTimer_Timer()
    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
        
    xValue = xValue + needleOffsetX
    yValue = yValue + needleOffsetY
        
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
        
    xSystemHome.Text = convertToMM(xValue, X_axis)
    ySystemHome.Text = convertToMM(yValue, Y_axis)
    zSystemHome.Text = convertToMM(zValue, Z_axis)
        
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
    Call validateNumber(mainForm.StepDistance.Text, mainForm.LabelDistance.Caption)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If mainForm.StepDistance.Text <= "0" Then
            mainForm.StepDistance.Text = "0.001"
        ElseIf mainForm.StepDistance.Text > "10" Then
            mainForm.StepDistance.Text = "10.000"
        Else
            mainForm.StepDistance.Text = Format(mainForm.StepDistance.Text, "#0.000")
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

Private Sub xMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = True
End Sub

Private Sub xPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    purgeButtonTimer.Enabled = True
End Sub

Private Sub yMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = True
End Sub

Private Sub yPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    purgeButtonTimer.Enabled = True
End Sub

Private Sub zMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = False
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
    purgeButtonTimer.Enabled = True
End Sub

Private Sub zPlus_mouseup(Button As Integer, Shift As Integer, X As Single, Y As Single)
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
    purgeButtonTimer.Enabled = True
End Sub

Private Sub jogSpeedSlider_Change()
    purgeButtonTimer.Enabled = False
    setSpeed (jogSpeedSlider.value - 1)
    purgeButtonTimer.Enabled = True
End Sub

Private Sub systemMoveHeight_Validate(cancel As Boolean)    'XW
    'Do the limitation in the text box
    Call validateNumber(mainForm.systemMoveHeight.Text, mainForm.SkinLabel23)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (mainForm.systemMoveHeight <> "") Then
            If CLng(mainForm.systemMoveHeight.Text) > 130 Then
                mainForm.systemMoveHeight.Text = "130"
            'ElseIf CLng(mainForm.systemMoveHeight.Text) < -14 Then
            '    mainForm.systemMoveHeight.Text = "-14"
            End If
            
            'If (CDbl(systemMoveHeight.Text) < 0) Then
            '    MsgBox "There is no nevative value in systemMoveHeight."
            '    systemMoveHeight.Text = ""
            '    Cancel = True
            'End If
        End If
    End If
End Sub

Private Sub purgeButtonTimer_Timer()        'XW (Purge Button)
    Dim buttonValue As Long, ReadValue As Long
    
    purgeButtonTimer.Enabled = False
    
    'Purge_Button input
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, RR4, buttonValue))
    buttonValue = buttonValue And &H400
    
    If (buttonValue = 0) Then
        cmdPurgeButton.Enabled = False
        SignalCounter = SignalCounter + 1
        PurgeTimer
        
        'If (LeftNeedle.value = True) Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H800
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
        'ElseIf (RightNeedle.value = True) Then
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue Or &H100
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        'End If
    Else
        If (DisablePurgeSignal = True) Then
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HF7FF
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
        
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HFEFF
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
            SignalCounter = 0
            PurgeTimer
            cmdPurgeButton.Enabled = True
        End If
    End If
    
    purgeButtonTimer.Enabled = True
End Sub

Private Sub PurgeTimer()
    systemTrackMoveHeight = convertToPulses(CLng(systemMoveHeight.Text), Z_axis)
    systemTrackMoveHeight = systemTrackMoveHeight * (-1)
    setSpeed (CLng(xyDefaultSpeed.Text))
            
    If (SignalCounter = 1) Then
        checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, PrePositionX, PrePositionY, PrePositionZ, PrePositionU))
    
        'To get the actual direcion
        PrePositionY = PrePositionY * (-1)
        PrePositionZ = PrePositionZ * (-1)
        
        PTPToXYZ convertToPulses(xPurgePosition.Text, X_axis), convertToPulses(yPurgePosition.Text, Y_axis), convertToPulses(zPurgePosition.Text, Z_axis)
        
        DisablePurgeSignal = True
    ElseIf (SignalCounter = 0) Then
        If (DisablePurgeSignal = True) Then
            PTPToXYZ PrePositionX, PrePositionY, PrePositionZ
            DisablePurgeSignal = False
        End If
    End If
End Sub

Private Sub CheckEmergencyStop_Timer()
    Dim CheckValueX, CheckValueY, CheckValueZ As Long
    Dim ValveClose As Long, DriverXYZ As Long
    
    Emergency_Stop = False
    CheckEmergencyStop.Enabled = False
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, CheckValueX, CheckValueY, CheckValueZ, 0))
    
    CheckValueX = (CheckValueX And &H20)
    CheckValueY = (CheckValueY And &H20)
    CheckValueZ = (CheckValueZ And &H20)
    
    If ((CheckValueX <> 0) Or (CheckValueY <> 0) Or (CheckValueZ <> 0)) Then
        Dim A As Long
        
        'Servo OFF (XW)
        checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, DriverXYZ))
        DriverXYZ = (DriverXYZ And &HF8FF)
        checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, DriverXYZ))
                    
        'Close both valves
        checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ValveClose))
        ValveClose = (ValveClose And &HF7FF)
        checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ValveClose))
            
        checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ValveClose))
        ValveClose = ValveClose And &HFEFF
        checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ValveClose))
            
        checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, A))
        If A < &H800 Then
            A = A Or &H800
            checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, A))
            
            Green_Light = False
            Red_Light = True
        End If
        frmEmergencyStopForm.Show (vbModal)
    End If
    
    If Emergency_Stop = True Then
        moveToHome
        Ext = True
    End If

    CheckEmergencyStop.Enabled = True
End Sub

Private Sub SaveParameterSetting()
    'Can't use and set the default value

    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.createtextfile(App.path & "\ParameterSetting.txt", True)
   
    A.writeline ("xyDefaultSpeed=" & mainForm.xyDefaultSpeed.Text)
    A.writeline ("zDefaultSpeed=" & mainForm.zDefaultSpeed.Text)
    A.writeline ("systemMoveHeight=" & mainForm.systemMoveHeight.Text)
    
    A.writeline ("xSystemHome=" & mainForm.xSystemHome.Text)
    A.writeline ("ySystemHome=" & mainForm.ySystemHome.Text)
    A.writeline ("zSystemHome=" & mainForm.zSystemHome.Text)
    
    A.writeline ("xPurgePosition=" & mainForm.xPurgePosition.Text)
    A.writeline ("yPurgePosition=" & mainForm.yPurgePosition.Text)
    A.writeline ("zPurgePosition=" & mainForm.zPurgePosition.Text)
    
    A.writeline ("xDatum=" & mainForm.xDatum.Text)
    A.writeline ("yDatum=" & mainForm.yDatum.Text)
    A.writeline ("zDatum=" & mainForm.zDatum.Text)
    
    A.writeline ("xSolventPos=" & mainForm.xSolventPos.Text)
    A.writeline ("ySolventPos=" & mainForm.ySolventPos.Text)
    A.writeline ("zSolventPos=" & mainForm.zSolventPos.Text)
   
    A.Close
End Sub

Public Sub ReadParameters()
    Dim VariableName As String
    Dim words() As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.OpenTextFile(App.path & "\ParameterSetting.txt", 1, False)
                
    Do While A.AtEndOfStream <> True
        For lines = 1 To 100
            VariableName = A.ReadLine
            
            words() = Split(VariableName, "=")
                        
            Select Case words(0)
                Case "xyDefaultSpeed"
                    mainForm.xyDefaultSpeed.Text = words(1)
                Case "zDefaultSpeed"
                    mainForm.zDefaultSpeed.Text = words(1)
                Case "systemMoveHeight"
                    mainForm.systemMoveHeight.Text = words(1)
                Case "xSystemHome"
                    mainForm.xSystemHome.Text = words(1)
                Case "ySystemHome"
                    mainForm.ySystemHome.Text = words(1)
                Case "zSystemHome"
                    mainForm.zSystemHome.Text = words(1)
                Case "xPurgePosition"
                    mainForm.xPurgePosition.Text = words(1)
                Case "yPurgePosition"
                    mainForm.yPurgePosition.Text = words(1)
                Case "zPurgePosition"
                    mainForm.zPurgePosition.Text = words(1)
                Case "xDatum"
                    mainForm.xDatum.Text = words(1)
                Case "yDatum"
                    mainForm.yDatum.Text = words(1)
                Case "zDatum"
                    mainForm.zDatum.Text = words(1)
            End Select
            
            If A.AtEndOfStream Then
                Exit For
            End If
        Next
    Loop
    A.Close
End Sub

Private Sub xyDefaultSpeed_GotFocus()
    TabNumber = 13
End Sub

Private Sub xyDefaultSpeed_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 13) Then
        xyDefaultSpeed.Text = TextString(xyDefaultSpeed.Text, OneSevenKey)
    End If
End Sub

Private Sub zDefaultSpeed_GotFocus()
    TabNumber = 14
End Sub

Private Sub zDefaultSpeed_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 14) Then
        zDefaultSpeed.Text = TextString(zDefaultSpeed.Text, OneSevenKey)
    End If
End Sub

Private Sub systemMoveHeight_GotFocus()
    TabNumber = 19
End Sub

Private Sub systemMoveHeight_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 19) Then
        systemMoveHeight.Text = TextString(systemMoveHeight.Text, OneSevenKey)
    End If
End Sub

Private Sub StepDistance_GotFocus()
    TabNumber = 83
End Sub

Private Sub StepDistance_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 83) Then
        StepDistance.Text = TextString(StepDistance.Text, OneSevenKey)
    End If
End Sub

Private Sub xSystemHome_GotFocus()
    TabNumber = 26
End Sub

Private Sub xSystemHome_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 26) Then
        xSystemHome.Text = TextString(xSystemHome.Text, OneSevenKey)
    End If
End Sub

Private Sub ySystemHome_GotFocus()
    TabNumber = 27
End Sub

Private Sub ySystemHome_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 27) Then
        ySystemHome.Text = TextString(ySystemHome.Text, OneSevenKey)
    End If
End Sub

Private Sub zSystemHome_GotFocus()
    TabNumber = 28
End Sub

Private Sub zSystemHome_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 28) Then
        zSystemHome.Text = TextString(zSystemHome.Text, OneSevenKey)
    End If
End Sub

Private Sub xDatum_GotFocus()
    TabNumber = 62
End Sub

Private Sub xDatum_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 62) Then
        xDatum.Text = TextString(xDatum.Text, OneSevenKey)
    End If
End Sub

Private Sub yDatum_GotFocus()
    TabNumber = 61
End Sub

Private Sub yDatum_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 61) Then
        yDatum.Text = TextString(yDatum.Text, OneSevenKey)
    End If
End Sub

Private Sub zDatum_GotFocus()
    TabNumber = 60
End Sub

Private Sub zDatum_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 60) Then
        zDatum.Text = TextString(zDatum.Text, OneSevenKey)
    End If
End Sub

Private Sub xPurgePosition_GotFocus()
    TabNumber = 36
End Sub

Private Sub xPurgePosition_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 36) Then
        xPurgePosition.Text = TextString(xPurgePosition.Text, OneSevenKey)
    End If
End Sub

Private Sub yPurgePosition_GotFocus()
    TabNumber = 37
End Sub

Private Sub yPurgePosition_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 37) Then
        yPurgePosition.Text = TextString(yPurgePosition.Text, OneSevenKey)
    End If
End Sub

Private Sub zPurgePosition_GotFocus()
    TabNumber = 38
End Sub

Private Sub zPurgePosition_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 38) Then
        zPurgePosition.Text = TextString(zPurgePosition.Text, OneSevenKey)
    End If
End Sub

Private Function TextString(ByVal StringVal As String, ByVal counter As Integer) As String
    Dim StringLenght As Integer
    
    StringLenght = Len(StringVal)
    If (StringLenght <= counter) Then
        TextString = Left(StringVal, counter - StringLenght)
    Else
        TextString = Left(StringVal, StringLenght - counter)
    End If
    
    TabNumber = 777
    OneSevenKey = 0
End Function

Private Sub xyDefaultSpeed_Validate(cancel As Boolean)
    Call validateNumber(mainForm.xyDefaultSpeed.Text, mainForm.SkinLabel2)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
    
    If (CDbl(xyDefaultSpeed.Text) <= 0) Then
        MsgBox "There is no zero and negative value in xyDefaultSpeed."
        xyDefaultSpeed.Text = ""
        cancel = True
    End If
End Sub

Private Sub zDefaultSpeed_Validate(cancel As Boolean)
    Call validateNumber(mainForm.zDefaultSpeed.Text, mainForm.SkinLabel3)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
    
    If (CDbl(zDefaultSpeed.Text) <= 0) Then
        MsgBox "There is no zero and negative value in zDefaultSpeed."
        zDefaultSpeed.Text = ""
        cancel = True
    End If
End Sub

Private Sub xSystemHome_Validate(cancel As Boolean)
    Call validateNumber(mainForm.xSystemHome.Text, mainForm.SkinLabel7)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub ySystemHome_Validate(cancel As Boolean)
    Call validateNumber(mainForm.ySystemHome.Text, mainForm.SkinLabel8)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub zSystemHome_Validate(cancel As Boolean)
    Call validateNumber(mainForm.zSystemHome.Text, mainForm.SkinLabel9)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub xDatum_Validate(cancel As Boolean)
    Call validateNumber(mainForm.xDatum.Text, mainForm.SkinLabel29)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub yDatum_Validate(cancel As Boolean)
    Call validateNumber(mainForm.yDatum.Text, mainForm.SkinLabel28)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub zDatum_Validate(cancel As Boolean)
    Call validateNumber(mainForm.zDatum.Text, mainForm.SkinLabel14)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub xPurgePosition_Validate(cancel As Boolean)
    Call validateNumber(mainForm.xPurgePosition.Text, mainForm.SkinLabel16)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub yPurgePosition_Validate(cancel As Boolean)
    Call validateNumber(mainForm.yPurgePosition.Text, mainForm.SkinLabel17)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub zPurgePosition_Validate(cancel As Boolean)
    Call validateNumber(mainForm.zPurgePosition.Text, mainForm.SkinLabel21)
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub Close_AllTimer()
    SetFocusTimer.Enabled = False
    SensorTimer.Enabled = False
    displayCoOrdsTimer.Enabled = False
    updateMoveHeightTimer.Enabled = False
    updateSystemHomeTimer.Enabled = False
    datumTimer.Enabled = False
    updatePurgePosition.Enabled = False
    resetTimer.Enabled = False
    purgeButtonTimer.Enabled = False
    MouseMovement.Enabled = False
    CheckEmergencyStop.Enabled = False
End Sub
