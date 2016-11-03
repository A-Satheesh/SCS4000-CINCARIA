VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form editorForm 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Profile Editor"
   ClientHeight    =   11010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "editorForm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   734
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1018
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtPitch 
      Height          =   285
      Left            =   1800
      TabIndex        =   136
      Text            =   "10"
      Top             =   9840
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblChooseFunction 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":08CA
      TabIndex        =   134
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdReferenceHigh 
      Caption         =   "Save Height"
      Height          =   495
      Left            =   6000
      TabIndex        =   133
      Top             =   1080
      Width           =   1455
   End
   Begin VB.PictureBox PicImage 
      FillColor       =   &H0000FF00&
      Height          =   5130
      Left            =   8280
      ScaleHeight     =   338
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   449
      TabIndex        =   77
      Top             =   4620
      Width           =   6795
   End
   Begin VB.TextBox txtOffsetY 
      Height          =   285
      Left            =   12000
      TabIndex        =   130
      Text            =   "0"
      Top             =   8400
      Width           =   1335
   End
   Begin VB.TextBox txtOffsetZ 
      Height          =   285
      Left            =   12000
      TabIndex        =   129
      Text            =   "0"
      Top             =   8640
      Width           =   1335
   End
   Begin VB.TextBox txtOffsetX 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   12000
      TabIndex        =   128
      Text            =   "0"
      Top             =   8160
      Width           =   1335
   End
   Begin VB.TextBox potDepth 
      Height          =   285
      Left            =   12000
      TabIndex        =   125
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox depthSpeed 
      Height          =   285
      Left            =   12000
      TabIndex        =   124
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox endDispenseHeight 
      Height          =   285
      Left            =   12000
      TabIndex        =   123
      Top             =   7440
      Width           =   1335
   End
   Begin VB.ListBox NodeType 
      Height          =   1815
      Left            =   1800
      TabIndex        =   74
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "Offset"
      Height          =   375
      Left            =   16920
      TabIndex        =   106
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdGDsiTimeSpeed 
      Caption         =   "GTimeSpeed"
      Height          =   375
      Left            =   16920
      TabIndex        =   108
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "editorForm.frx":094A
      Left            =   1800
      List            =   "editorForm.frx":095D
      Style           =   2  'Dropdown List
      TabIndex        =   117
      Top             =   3840
      Width           =   2655
   End
   Begin VB.CheckBox VisionTeach 
      Caption         =   "Vision Teach (Not for rotation)"
      Height          =   255
      Left            =   6000
      TabIndex        =   116
      Top             =   4200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.ComboBox RotationAngle 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "editorForm.frx":09C7
      Left            =   1800
      List            =   "editorForm.frx":09DD
      TabIndex        =   114
      Text            =   "None"
      Top             =   6240
      Width           =   1575
   End
   Begin VB.Frame NeedleMode 
      Caption         =   "Needle Mode"
      Height          =   615
      Left            =   6000
      TabIndex        =   110
      Top             =   4560
      Width           =   2190
      Begin VB.OptionButton LeftNeedle 
         Caption         =   "Left "
         Height          =   255
         Left            =   240
         TabIndex        =   112
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton RightNeedle 
         Caption         =   "Right "
         Height          =   255
         Left            =   1320
         TabIndex        =   111
         Top             =   240
         Width           =   800
      End
   End
   Begin VB.CommandButton cmdGoffset 
      Caption         =   "Goffset"
      Height          =   375
      Left            =   16920
      TabIndex        =   109
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel ProgramStep 
      Height          =   255
      Left            =   12720
      OleObjectBlob   =   "editorForm.frx":0A01
      TabIndex        =   107
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton TeachFudicialPt 
      Caption         =   "Teach Fiducial Points"
      Height          =   495
      Left            =   13200
      TabIndex        =   104
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton AbortNeedleOffset 
      Caption         =   "Abort"
      Enabled         =   0   'False
      Height          =   495
      Left            =   13200
      TabIndex        =   103
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton CancelFud 
      Caption         =   "Cancel Fiducial"
      Height          =   495
      Left            =   13200
      TabIndex        =   102
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
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
      Left            =   7440
      TabIndex        =   97
      Text            =   "20"
      Top             =   10080
      Width           =   495
   End
   Begin VB.ListBox lstPattern 
      Height          =   2985
      ItemData        =   "editorForm.frx":0A5F
      Left            =   7725
      List            =   "editorForm.frx":0A61
      MultiSelect     =   2  'Extended
      TabIndex        =   93
      Top             =   1080
      Width           =   7455
   End
   Begin VB.Timer MouseMovement 
      Interval        =   200
      Left            =   1200
      Top             =   2520
   End
   Begin VB.CommandButton Expand 
      Caption         =   "Expand"
      Height          =   375
      Left            =   16920
      TabIndex        =   92
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel FocusLstBox 
      Height          =   255
      Left            =   8880
      OleObjectBlob   =   "editorForm.frx":0A63
      TabIndex        =   91
      Top             =   4200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Timer SetFocusTimer 
      Interval        =   1000
      Left            =   720
      Top             =   2520
   End
   Begin ACTIVESKINLibCtl.SkinLabel LabelDistance 
      Height          =   255
      Left            =   6000
      OleObjectBlob   =   "editorForm.frx":0AFD
      TabIndex        =   88
      Top             =   6000
      Width           =   1020
   End
   Begin ComCtl2.UpDown UpDownStep 
      Height          =   255
      Left            =   7920
      TabIndex        =   87
      Top             =   6000
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      Enabled         =   -1  'True
   End
   Begin VB.TextBox StepDistance 
      Height          =   285
      Left            =   7035
      TabIndex        =   86
      Text            =   "1.000"
      Top             =   6000
      Width           =   870
   End
   Begin VB.Frame JoggingMode 
      Caption         =   "Jogging Mode"
      Height          =   615
      Left            =   6000
      TabIndex        =   85
      Top             =   5280
      Width           =   2190
      Begin VB.OptionButton Jogging 
         Caption         =   "Jog"
         Height          =   255
         Left            =   1320
         TabIndex        =   90
         Top             =   240
         Width           =   800
      End
      Begin VB.OptionButton JoggingStep 
         Caption         =   "Step"
         Height          =   255
         Left            =   240
         TabIndex        =   89
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Timer ClickTimer 
      Interval        =   450
      Left            =   240
      Top             =   2520
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   255
      Left            =   6480
      OleObjectBlob   =   "editorForm.frx":0B75
      TabIndex        =   82
      Top             =   8880
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Index           =   0
      Left            =   6960
      OleObjectBlob   =   "editorForm.frx":0BD5
      TabIndex        =   81
      Top             =   6360
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Index           =   0
      Left            =   6240
      OleObjectBlob   =   "editorForm.frx":0C35
      TabIndex        =   80
      Top             =   6960
      Width           =   135
   End
   Begin VB.Timer CheckEmergencyStop 
      Interval        =   100
      Left            =   1200
      Top             =   3000
   End
   Begin VB.Timer resetTimer 
      Interval        =   1000
      Left            =   720
      Top             =   3000
   End
   Begin VB.CommandButton xPlus 
      Height          =   615
      Left            =   6000
      Picture         =   "editorForm.frx":0C95
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton yMinus 
      Height          =   615
      Left            =   6720
      Picture         =   "editorForm.frx":1070
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7800
      Width           =   735
   End
   Begin VB.CommandButton yPlus 
      Height          =   615
      Left            =   6720
      Picture         =   "editorForm.frx":147A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6600
      Width           =   735
   End
   Begin VB.CommandButton xMinus 
      Height          =   615
      Left            =   7440
      Picture         =   "editorForm.frx":1864
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7200
      Width           =   735
   End
   Begin VB.CommandButton goCommand 
      Caption         =   "Go"
      Height          =   375
      Left            =   3480
      MaskColor       =   &H00C0C0FF&
      TabIndex        =   79
      Top             =   1200
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LimitReachedLabel 
      Height          =   495
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":1C30
      TabIndex        =   78
      Top             =   1200
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer FudicialTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   600
      Top             =   4680
   End
   Begin VB.CommandButton decideDispensePt 
      Caption         =   "Teach On"
      Height          =   375
      Left            =   3360
      TabIndex        =   76
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton translateButton 
      Caption         =   "Translate"
      Height          =   495
      Left            =   13680
      TabIndex        =   75
      Top             =   360
      Width           =   1455
   End
   Begin VB.Timer displayCoOrdsTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   4680
   End
   Begin VB.Timer TimerDrawStatus 
      Interval        =   1000
      Left            =   -3720
      Top             =   4920
   End
   Begin VB.CheckBox dispenseOnOff 
      Caption         =   "Dispense On"
      Height          =   255
      Left            =   360
      TabIndex        =   73
      Top             =   6960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
      Height          =   255
      Left            =   6240
      OleObjectBlob   =   "editorForm.frx":1CEC
      TabIndex        =   72
      Top             =   120
      Width           =   1215
   End
   Begin VB.PictureBox PictureError 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   7080
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   71
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox PictureBusy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6480
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   70
      Top             =   360
      Width           =   615
   End
   Begin VB.PictureBox PictureReady 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   5880
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   69
      Top             =   360
      Width           =   615
   End
   Begin MSComCtl2.UpDown UpDownYRepeatNum 
      Height          =   285
      Left            =   3120
      TabIndex        =   68
      Top             =   9000
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2
      BuddyControl    =   "yRepeatNum"
      BuddyDispid     =   196683
      OrigLeft        =   3240
      OrigTop         =   9339
      OrigRight       =   3480
      OrigBottom      =   9744
      Max             =   200
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownXRepeatNum 
      Height          =   285
      Left            =   3120
      TabIndex        =   67
      Top             =   8760
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2
      BuddyControl    =   "xRepeatNum"
      BuddyDispid     =   196682
      OrigLeft        =   3240
      OrigTop         =   8958
      OrigRight       =   3480
      OrigBottom      =   9363
      Max             =   200
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownWithDrawalSpeed 
      Height          =   285
      Left            =   3120
      TabIndex        =   66
      Top             =   7800
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "withdrawalSpeed"
      BuddyDispid     =   196669
      OrigLeft        =   3240
      OrigTop         =   7624
      OrigRight       =   3480
      OrigBottom      =   8029
      Max             =   200
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownRetractDelay 
      Height          =   255
      Left            =   3120
      TabIndex        =   65
      Top             =   7560
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDispenseSpeed 
      Height          =   285
      Left            =   3120
      TabIndex        =   64
      Top             =   5760
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "DispenseSpeed"
      BuddyDispid     =   196666
      OrigLeft        =   3240
      OrigTop         =   5336
      OrigRight       =   3480
      OrigBottom      =   5741
      Max             =   200
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDelay 
      Height          =   285
      Left            =   3120
      TabIndex        =   63
      Top             =   5520
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "delay"
      BuddyDispid     =   196668
      OrigLeft        =   3240
      OrigTop         =   4765
      OrigRight       =   3480
      OrigBottom      =   5170
      Increment       =   0
      Max             =   100
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDispenseTime 
      Height          =   255
      Left            =   3120
      TabIndex        =   62
      Top             =   5280
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton loadPartArray 
      Caption         =   "Load"
      Height          =   255
      Left            =   5640
      TabIndex        =   61
      Top             =   10320
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel pathFileNameLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":1D64
      TabIndex        =   60
      Top             =   10320
      Width           =   1215
   End
   Begin VB.TextBox PathFileName 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   10320
      Width           =   3735
   End
   Begin VB.CommandButton decideMoveHeight 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3360
      TabIndex        =   58
      Top             =   8280
      Width           =   1095
   End
   Begin VB.TextBox moveHeight 
      Height          =   285
      Left            =   1800
      TabIndex        =   57
      Top             =   8280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel moveHeightLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":1DDC
      TabIndex        =   56
      Top             =   8280
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   120
      TabIndex        =   55
      Top             =   6720
      Width           =   3255
      Begin VB.CheckBox NoSprayArea 
         Caption         =   "No Spray Area"
         Height          =   255
         Left            =   1680
         TabIndex        =   113
         Top             =   240
         Width           =   1455
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel delayLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":1E50
      TabIndex        =   54
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox dispenseTime 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   53
      Text            =   "1.0"
      Top             =   5280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel DispenseTimeLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":1EC6
      TabIndex        =   52
      Top             =   5280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "editorForm.frx":1F3E
      TabIndex        =   51
      Top             =   9600
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   5040
      OleObjectBlob   =   "editorForm.frx":1FA2
      TabIndex        =   50
      Top             =   4560
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   4530
      OleObjectBlob   =   "editorForm.frx":2006
      TabIndex        =   49
      Top             =   6720
      Width           =   495
   End
   Begin ComctlLib.Slider jogSpeedSlider 
      Height          =   4815
      Left            =   5040
      TabIndex        =   48
      Top             =   4800
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   8493
      _Version        =   327682
      Orientation     =   1
      Min             =   2
      Max             =   151
      SelStart        =   28
      TickStyle       =   2
      Value           =   28
   End
   Begin VB.CommandButton trackButton 
      Caption         =   "Track"
      Height          =   375
      Left            =   16920
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel NodeLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":2076
      TabIndex        =   45
      Top             =   4320
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   3600
      OleObjectBlob   =   "editorForm.frx":20DE
      TabIndex        =   44
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   1920
      OleObjectBlob   =   "editorForm.frx":213E
      TabIndex        =   43
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":219E
      TabIndex        =   42
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel yRepeatNumLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":21FE
      TabIndex        =   41
      Top             =   8760
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel yDevLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":2276
      TabIndex        =   40
      Top             =   9480
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel xRepeatNumLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":22EA
      TabIndex        =   39
      Top             =   9000
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel xDevLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":235C
      TabIndex        =   38
      Top             =   9240
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel retractDelayLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":23D0
      TabIndex        =   37
      Top             =   7560
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel withDrawalSpeedLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":2448
      TabIndex        =   36
      Top             =   7800
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel withDrawalZLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":24C0
      TabIndex        =   35
      Top             =   8040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel dispenseSpeedLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":253A
      TabIndex        =   34
      Top             =   5760
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel dispensePtZLabel 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "editorForm.frx":25B0
      TabIndex        =   33
      Top             =   4800
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel YLabel 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "editorForm.frx":2610
      TabIndex        =   32
      Top             =   4560
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel XLabel 
      Height          =   255
      Left            =   1560
      OleObjectBlob   =   "editorForm.frx":2670
      TabIndex        =   31
      Top             =   4320
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel NodeTypeLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":26D0
      TabIndex        =   30
      Top             =   1920
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":2748
      Top             =   3000
   End
   Begin VB.CommandButton Home 
      Caption         =   "Home"
      Height          =   495
      Left            =   6000
      TabIndex        =   29
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox DispenseSpeed 
      Height          =   285
      Left            =   1800
      TabIndex        =   28
      Text            =   "10"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox retractDelay 
      Height          =   285
      Left            =   1800
      TabIndex        =   27
      Text            =   "1.0"
      Top             =   7560
      Width           =   1335
   End
   Begin VB.TextBox delay 
      Height          =   285
      Left            =   1800
      TabIndex        =   26
      Text            =   "1.0"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox withdrawalSpeed 
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Text            =   "10"
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox WithDrawalZ 
      Height          =   285
      Left            =   1800
      TabIndex        =   24
      Top             =   8040
      Width           =   1335
   End
   Begin VB.TextBox xCoOrd 
      Height          =   375
      Left            =   360
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox yCoOrd 
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox zCoOrd 
      Height          =   375
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton deletePt 
      Caption         =   "Delete Node"
      Height          =   495
      Left            =   6000
      TabIndex        =   20
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton addPt 
      Caption         =   "Add Node"
      Height          =   495
      Left            =   6000
      TabIndex        =   19
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton decideWithdrawalHeight 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3360
      TabIndex        =   18
      Top             =   8040
      Width           =   1095
   End
   Begin VB.TextBox dispensePtY 
      Height          =   285
      Left            =   1800
      TabIndex        =   17
      Top             =   4560
      Width           =   1335
   End
   Begin VB.TextBox dispensePtZ 
      Height          =   285
      Left            =   1800
      TabIndex        =   16
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox dispensePtX 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   15
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox xDev 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Top             =   9240
      Width           =   1335
   End
   Begin VB.TextBox yDev 
      Height          =   285
      Left            =   1800
      TabIndex        =   13
      Top             =   9480
      Width           =   1335
   End
   Begin VB.TextBox xRepeatNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   12
      Text            =   "1"
      Top             =   8760
      Width           =   1335
   End
   Begin VB.TextBox yRepeatNum 
      Height          =   285
      Left            =   1800
      TabIndex        =   11
      Text            =   "1"
      Top             =   9000
      Width           =   1335
   End
   Begin VB.CommandButton decideXYDev 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   9240
      Width           =   1095
   End
   Begin VB.CommandButton modifyPt 
      Caption         =   "Modify Node"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   2880
      Width           =   1455
   End
   Begin VB.CommandButton fileNew 
      Caption         =   "New"
      Height          =   495
      Left            =   7920
      TabIndex        =   8
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton fileLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   9840
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton fileSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   11760
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton zPlus 
      Height          =   615
      Left            =   6720
      Picture         =   "editorForm.frx":297C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8520
      Width           =   735
   End
   Begin VB.CommandButton zMinus 
      Height          =   615
      Left            =   6720
      Picture         =   "editorForm.frx":2D66
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9120
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Robot Position"
      Height          =   735
      Left            =   120
      TabIndex        =   46
      Top             =   240
      Width           =   5175
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   255
      Index           =   1
      Left            =   6960
      OleObjectBlob   =   "editorForm.frx":3170
      TabIndex        =   83
      Top             =   7560
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Index           =   1
      Left            =   7800
      OleObjectBlob   =   "editorForm.frx":31D0
      TabIndex        =   84
      Top             =   6960
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   255
      Left            =   7440
      OleObjectBlob   =   "editorForm.frx":3230
      TabIndex        =   94
      Top             =   9840
      Width           =   735
   End
   Begin ComCtl2.UpDown UpDownLighting 
      Height          =   255
      Left            =   7920
      TabIndex        =   96
      Top             =   10080
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      BuddyControl    =   "LightingIntensity"
      BuddyDispid     =   196631
      OrigLeft        =   504
      OrigTop         =   656
      OrigRight       =   520
      OrigBottom      =   673
      Max             =   255
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel RotationAngleLabel 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":32A2
      TabIndex        =   115
      Top             =   6240
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel FudMsgText 
      Height          =   615
      Left            =   8280
      OleObjectBlob   =   "editorForm.frx":331C
      TabIndex        =   105
      Top             =   9840
      Width           =   2775
   End
   Begin ACTIVESKINLibCtl.SkinLabel NeedleOffsetErrMsg 
      Height          =   495
      Left            =   8400
      OleObjectBlob   =   "editorForm.frx":337A
      TabIndex        =   95
      Top             =   9840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSComCtl2.UpDown UpDownEndDispenseHeight 
      Height          =   255
      Left            =   13320
      TabIndex        =   118
      Top             =   7440
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDepthSpeed 
      Height          =   255
      Left            =   13320
      TabIndex        =   119
      Top             =   7200
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownPotDepth 
      Height          =   255
      Left            =   13320
      TabIndex        =   120
      Top             =   6960
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel endDispenseHeightLabel 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":3452
      TabIndex        =   121
      Top             =   7440
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel depthSpeedLabel 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":34CE
      TabIndex        =   122
      Top             =   7200
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel PotDepthLabel 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":3542
      TabIndex        =   126
      Top             =   6960
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetX 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":35BA
      TabIndex        =   127
      Top             =   8160
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetY 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":3632
      TabIndex        =   131
      Top             =   8400
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetZ 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":36AA
      TabIndex        =   132
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Calibrate 
      Caption         =   "Calibrate"
      Height          =   495
      Left            =   11160
      TabIndex        =   101
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Calibrate1 
      Caption         =   "Calibrate"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11160
      TabIndex        =   100
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton NextFudStep 
      Caption         =   "Next"
      Height          =   495
      Left            =   11160
      TabIndex        =   99
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton FindNeedleOffset 
      Caption         =   "Needle Calibration"
      Enabled         =   0   'False
      Height          =   495
      Left            =   11160
      TabIndex        =   98
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSCommLib.MSComm mscomLighIntensity 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblPitch 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "editorForm.frx":3722
      TabIndex        =   135
      Top             =   9840
      Width           =   1215
   End
   Begin VB.Menu mnuCustom 
      Caption         =   "&Custom"
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "C&opy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuCancel 
         Caption         =   "C&ancel"
      End
   End
End
Attribute VB_Name = "editorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variable for mouse's movement
Dim lngReturn As Long
Dim OneSevenKey As Integer, TabNumber As Integer
Dim lngScreenX As Long, lngScreenY As Long
Const MaxSpeed As Long = 100
Const MaxMouseRangeX As Long = 1023
Const MaxMouseRangeY As Long = 767
Dim GroupOffset As Boolean              'Just a flage for group modify.
Dim AddingElement As Boolean            'Give a flag to remove the old program line after adding a new one in "GlobalDispenseTime" function
Dim LeftValve As Boolean                'Just a flag for choosing left-valve.
Dim RightValve As Boolean               'Just a flag for choosing right-valve.
Dim LeftNeedle_No As Integer, RightNeedle_No As Integer   'To add for one time
Dim Save_Angle As String                'Save rotation value
Dim Save_Index As Integer
Dim Save_Old_zHigh As Long                  'Save the old_reference_z_high

Private Sub lockMoveControls()
    fileNew.Enabled = False
    fileLoad.Enabled = False
    fileSave.Enabled = False
    Home.Enabled = False
    translateButton.Enabled = False
    addPt.Enabled = False
    modifyPt.Enabled = False
    deletePt.Enabled = False
    TeachFudicialPt.Enabled = False
    FindNeedleOffset.Enabled = False
    zPlus.Enabled = False
    zMinus.Enabled = False
    xPlus.Enabled = False
    xMinus.Enabled = False
    yPlus.Enabled = False
    yMinus.Enabled = False
    lstPattern.Enabled = False
    RotationAngle.Enabled = False
End Sub

Private Sub lockSaveMoveControls()      'XW
    fileNew.Enabled = False
    fileLoad.Enabled = False
    fileSave.Enabled = False
    Home.Enabled = False
    addPt.Enabled = False
    modifyPt.Enabled = False
    deletePt.Enabled = False
    TeachFudicialPt.Enabled = False
    FindNeedleOffset.Enabled = False
    zPlus.Enabled = False
    zMinus.Enabled = False
    xPlus.Enabled = False
    xMinus.Enabled = False
    yPlus.Enabled = False
    yMinus.Enabled = False
    goCommand.Enabled = False
    cmdOffset.Enabled = False
    decideDispensePt.Enabled = False
    'cmdGlobalOffset.Enabled = False
    'cmdGlobalDispenseTimeSpeed.Enabled = False
    lstPattern.Enabled = False
    RotationAngle.Enabled = False
    cmdReferenceHigh.Enabled = False
End Sub

Private Sub lockNonMoveControls()
    fileNew.Enabled = False
    fileLoad.Enabled = False
    fileSave.Enabled = False
    Home.Enabled = False
    translateButton.Enabled = False
    addPt.Enabled = False
    modifyPt.Enabled = False
    deletePt.Enabled = False
    lstPattern.Enabled = False
    cmdReferenceHigh.Enabled = False
End Sub

Private Sub unLockNonMoveControls()
    fileNew.Enabled = True
    fileLoad.Enabled = True
    fileSave.Enabled = True
    Home.Enabled = True
    translateButton.Enabled = True
    addPt.Enabled = True
    modifyPt.Enabled = True
    deletePt.Enabled = True
    lstPattern.Enabled = True
    cmdReferenceHigh.Enabled = True
End Sub

Private Sub unLockMoveControls()
    fileNew.Enabled = True
    fileLoad.Enabled = True
    fileSave.Enabled = True
    Home.Enabled = True
    translateButton.Enabled = True
    addPt.Enabled = True
    modifyPt.Enabled = True
    deletePt.Enabled = True
    TeachFudicialPt.Enabled = True
    'FindNeedleOffset.Enabled = True
    zPlus.Enabled = True
    zMinus.Enabled = True
    xPlus.Enabled = True
    xMinus.Enabled = True
    yPlus.Enabled = True
    yMinus.Enabled = True
    lstPattern.Enabled = True
    'RotationAngle.Enabled = True
End Sub

Private Sub unLockSaveMoveControls()
    fileNew.Enabled = True
    fileLoad.Enabled = True
    fileSave.Enabled = True
    Home.Enabled = True
    addPt.Enabled = True
    modifyPt.Enabled = True
    deletePt.Enabled = True
    TeachFudicialPt.Enabled = True
    FindNeedleOffset.Enabled = True
    zPlus.Enabled = True
    zMinus.Enabled = True
    xPlus.Enabled = True
    xMinus.Enabled = True
    yPlus.Enabled = True
    yMinus.Enabled = True
    goCommand.Enabled = True
    cmdOffset.Enabled = True
    decideDispensePt.Enabled = True
    'cmdGlobalOffset.Enabled = True
    'cmdGlobalDispenseTimeSpeed.Enabled = True
    lstPattern.Enabled = True
    RotationAngle.Enabled = True
    cmdReferenceHigh.Enabled = True
End Sub

Private Sub AbortNeedleOffset_Click()
    FudMsgText.Visible = False
    AbortNeedleOffset.Visible = False
    NeedleOffsetErrMsg.Visible = False
    Calibrate1.Visible = False
    'FindNeedleOffset.Visible = True
    TeachFudicialPt.Visible = True
    VdeSelectCamera 2
    SetLightIntensity (camera2LightSetting)
    LightingIntensity.Text = camera2LightSetting
    unLockNonMoveControls
    unLockMoveControls
End Sub

Private Sub Calibrate_Click()
    
    lockNonMoveControls
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    
    PTPToXYZ xDatum, yDatum, SystemMoveHeight
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), SystemMoveHeight
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationZ", "0")
    Calibrate.Visible = False
    'FindNeedleOffset.Visible = False
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

'Private Sub Command1_Click()
'Dim dx As Double
'Dim dy As Double
'Dim da As Double

    'patFile = App.Path & "\pat"

    'VdeFindRefPt patFile, CDbl(txtRef1x), CDbl(txtRef1y), CDbl(txtRef2x), CDbl(txtRef2y), dx, dy, da
    'MsgBox "Ref Pt offset = " & CStr(dx) & ", " & CStr(dy) & ", " & CStr(da)
'End Sub

Private Sub CancelFud_Click()
    FudicialTimer.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgCancel, s)
    FudMsgText.Caption = ""
    CancelFud.Visible = False
    NextFudStep.Visible = False
    TeachFudicialPt.Visible = True
    'FindNeedleOffset.Visible = True

    VdeCameraLive 1
    unLockNonMoveControls
    jogSpeedSlider.value = 28
    Call setSpeed(jogSpeedSlider.value)         'XW
End Sub

Private Sub cmdReferenceHigh_Click()
    If (LeftNeedle.value = True) Then
        Z_High = convertToPulses(dispensePtZ.Text, Z_axis)
        reference_ZHigh = True
    Else
        R_Z_High = convertToPulses(dispensePtZ.Text, Z_axis)
        reference_R_ZHigh = True
    End If
    
    jogSpeedSlider.value = 28
    setSpeed (28)
    Call PTPToXYZ(convertToPulses(xCoOrd.Text, X_axis), convertToPulses(yCoOrd.Text, Y_axis), 0)
    'cmdReferenceHigh.Enabled = False
End Sub

Private Sub Combo1_Click()
    SetFocusTimer.Enabled = False
    
    If (Combo1.Text = "Expand Dot Array") Or (Combo1.Text = "Individual Offset") Then
        If (lstPattern.SelCount = 0) Then
            NodeError.Show (vbModal)
            SetFocusTimer.Enabled = True
            Exit Sub
        End If
    ElseIf (Combo1.Text = "Global Offset") Or (Combo1.Text = "Global Dispense Time/Speed") Then
        If (lstPattern.ListCount = 0) Then
            MsgBox "There is no programming position."
            SetFocusTimer.Enabled = True
            Exit Sub
        End If
    End If
    
    If (Combo1.Text = "Expand Dot Array") Then
        Expand_Click
    ElseIf (Combo1.Text = "Global Offset") Then
        frmGlobalOffset.Show (vbModal)
    
        If (FlagGlobalOffset = True) Then
            cmdGoffset_Click
            editorForm.Caption = "Profile Editor"
        End If
    ElseIf (Combo1.Text = "Global Dispense Time/Speed") Then
        frmGlobalTimeSpeed.Show (vbModal)
    
        If (FlagGlobalTimeSpeed = True) Then
            cmdGDsiTimeSpeed_Click
            editorForm.Caption = "Profile Editor"
        End If
    ElseIf (Combo1.Text = "Individual Offset") Then
        frmOffset.Show (vbModal)
        
        If (FlagGlobalOffset = True) Then
            cmdOffset_Click
            editorForm.Caption = "Profile Editor"
        End If
    ElseIf (Combo1.Text = "Measure the distance") Then
        frmDistance.Show (vbModal)
    End If
    
    GlobalOffsetX = "0"
    GlobalOffsetY = "0"
    GlobalOffsetZ = "0"
    SetFocusTimer.Enabled = True
End Sub

Private Sub LeftNeedle_Click()

    Move_To_Zero
    
    LeftNeedle.value = True
    RightNeedle.value = False
  
    'Do it when changing from the right valve.
    If (LeftValve = False) And (RightValve = True) Then
        LeftNeedleValve
        
        If (VisionTeach.Enabled = False) Then
            VisionTeach.Enabled = True
        End If
    End If
    
    Tilt_Off
    'Go back to "angle 0"
    RotationAngle.Text = "None"
    Call Tilt_Rotate(0)
    
    RotationAngle.Enabled = False
    RotationAngleLabel.Enabled = False
    NoSprayArea.value = 0
    NoSprayArea.Enabled = False
    
    If (NodeType.ListIndex = 16) Then
        decideXYDev.Enabled = False
        yDev.Enabled = False
        yDevLabel.Enabled = False
    End If
    
    If (NodeType.ListIndex = 18) Then
        If (dispenseOnOff.Caption = "Always on") Then
            dispenseOnOff.Caption = "Dispense On"
            dispenseOnOff.Refresh
        End If
    End If
    
    ForLeftNeedle
End Sub

Private Sub Move_To_Zero()
    Call setSpeed(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50"))
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
        DoEvents
    Loop
    
End Sub

Private Sub NoSprayArea_Click()
    If (NoSprayArea.value = 1) And (Click = 0) Then
        MsgBox "The direction for teaching point should be the same as previous rectangle!"
    End If
End Sub

Private Sub RightNeedle_Click()
    
    Dim Ind As Integer

    Move_To_Zero
    
    RightNeedle.value = True
    LeftNeedle.value = False


    'Do it when changing from the left valve.
    If (LeftValve = True) And (RightValve = False) Then
        RightNeedleValve

        If (VisionTeach.Enabled = False) Then
            VisionTeach.Enabled = True
        End If
    End If

    RotationAngle.Enabled = True
    Ind = NodeType.ListIndex

    If (Ind = 0) Or (Ind = 2) Or (Ind = 5) Or (Ind = 7) Or (Ind = 8) Or (Ind = 9) Or (Ind = 12) Or (Ind = 13) Or (Ind = 14) Then
        RotationAngle.Enabled = False
        RotationAngleLabel.Enabled = False
        Check_Node
    ElseIf (Ind = 16) Or (Ind = 17) Or (Ind = 18) Then
        NoSprayArea.Enabled = True
        RotationAngle.Enabled = False
        RotationAngleLabel.Enabled = False
        If (Ind = 16) Then
            decideXYDev.Enabled = True
            yDev.Enabled = True
            yDevLabel.Enabled = True
        End If
    Else
        RotationAngleLabel.Enabled = True
        If (RotationAngle.Text <> "None") And (RotationAngle.Text <> "") Then
            Tilt_ON
            Call Tilt_Rotate(RotationAngle.Text)
            Save_Angle = ""
            Save_Index = RotationAngle.ListIndex
            Save_Angle = RotationAngle.List(RotationAngle.ListIndex)
        End If
    End If

    If (Ind = 18) Then
        If (dispenseOnOff.Caption = "Dispense On") Then
            dispenseOnOff.Caption = "Always on"
            dispenseOnOff.Refresh
        End If
    End If

    ForRightNeedle
End Sub

Private Sub Right_Or_Left()
    Dim i As Integer
    
    If ((lstPattern.SelCount >= 1) And (lstPattern.ListIndex = 0)) Then
        LeftNeedle_No = 0
        RightNeedle_No = 0
        Exit Sub
    End If
    
    For i = 0 To (lstPattern.ListCount - 1)
        If (i >= lstPattern.ListCount) Then
            Exit For
        End If
        
        If (lstPattern.List(i) = "*** Left-Needle ***") Then
            RightNeedle_No = 0
            LeftNeedle_No = 1
        ElseIf (lstPattern.List(i) = "*** Right-Needle ***") Then
            LeftNeedle_No = 0
            RightNeedle_No = 1
        End If
    Next i
End Sub

Private Sub ForLeftNeedle()
    Right_Or_Left
    RightValve = False
    LeftValve = True
End Sub

Private Sub ForRightNeedle()
    Right_Or_Left
    LeftValve = False
    RightValve = True
End Sub

Private Sub FindNeedleOffset_Click()
    
    lockMoveControls
    lockNonMoveControls
    
    TeachFudicialPt.Visible = False
    'camera2LightSetting = LightingIntensity.Text
    
    'VdeSelectCamera 1
    setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
    
    Dim xDatum, yDatum, zDatum As Long
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    zDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zDatum", "0"), Z_axis)

    systemTrackMoveHeight = SystemMoveHeight
    
    PTPToXYZ xDatum, yDatum, SystemMoveHeight
    PTPToXYZ xDatum, yDatum, zDatum
    'FindNeedleOffset.Visible = False
    'Calibrate.Visible = True
    FudMsgText.Visible = True
    FudMsgText.Caption = "Mount syring with needle tip resting on datum, click Calibrate->"
    'AbortNeedleOffset.Visible = True

End Sub

Private Sub FudicialTimer_Timer()
    FudicialTimer.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgOnTimer, s)
    FudicialTimer.Enabled = True
End Sub

Private Sub goCommand_Click()
    If busyStatus = False Then
        goCommand.Enabled = False
        setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
        
        Call PTPToXYZ(convertToPulses(xCoOrd.Text, X_axis), convertToPulses(yCoOrd.Text, Y_axis), convertToPulses(zCoOrd.Text, Z_axis))
       
        displayCoOrdsTimer.Enabled = True
        xCoOrd.Locked = True
        yCoOrd.Locked = True
        zCoOrd.Locked = True
        unLockMoveControls
        goCommand.Enabled = True
        setSpeed (jogSpeedSlider.value - 1)
        'XW
        editorForm.KeyPreview = True
    End If
End Sub

Private Sub jogSpeedSlider_Change()
    setSpeed (jogSpeedSlider.value - 1)
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

Private Function NotArray() As Boolean
    Dim words() As String
    
    words() = Split(editorForm.lstPattern.List(editorForm.lstPattern.ListIndex), "(")
    
    If (words(0) = "StartDotArray") Or (words(0) = "StartDotPottingArray") Or (words(0) = "StartLinePottingArray") Or (words(0) = "EndArray") Then
        MsgBox ("System doesn't allow array to do 'Copy' and 'Paste'!")
        NotArray = True
    Else
        NotArray = False
    End If
End Function

Private Sub ShowingMessageBox()
    If (editorForm.lstPattern.SelCount > 1) Then
        MsgBox ("Only one line is allowed to do 'Copy' and 'Paste'!")
    ElseIf (editorForm.lstPattern.SelCount = 0) Then
        MsgBox ("Please select one programming line first!")
    End If
End Sub

Private Sub mnucut_click()
    'Origin
    'If (editorForm.lstPattern.ListIndex <> -1) Then
    If (editorForm.lstPattern.SelCount = 1) Then
        If (NotArray = False) Then
            tempCutPasteString = editorForm.lstPattern.List(editorForm.lstPattern.ListIndex)
            editorForm.lstPattern.RemoveItem (editorForm.lstPattern.ListIndex)
        End If
    Else
        ShowingMessageBox
    End If
End Sub

Private Sub mnucopy_click()
    'Origin
    'If (editorForm.lstPattern.ListIndex <> -1) Then
    If (editorForm.lstPattern.SelCount = 1) Then
        If (NotArray = False) Then
            tempCutPasteString = editorForm.lstPattern.List(editorForm.lstPattern.ListIndex)
        End If
    Else
        ShowingMessageBox
    End If
End Sub

Private Sub mnupaste_click()
    'Origin
    'If (editorForm.lstPattern.ListIndex <> -1) Then
    If (editorForm.lstPattern.SelCount = 1) Then
        Call lstPattern.AddItem(tempCutPasteString, editorForm.lstPattern.ListIndex)
        'XW
        editorForm.lstPattern.ListIndex = editorForm.lstPattern.ListIndex + 1
    Else
        ShowingMessageBox
    End If
End Sub

Private Function SelectedItems(ByVal Count As Integer, LstIndex As Integer) As Integer
    Dim no As Integer
    Dim counter As Integer
    
    LstIndex = 0
    counter = 0
    
    For no = 0 To Count - 1
        If lstPattern.Selected(no) = True Then
            LstIndex = no
            counter = counter + 1
        End If
    Next no
    
    SelectedItems = counter
End Function

Private Function Different_Part() As Boolean
    Dim i As Integer
    Dim Compare_String As String
    
    Compare_String = ""
    For i = 0 To lstPattern.ListCount - 1
        If (lstPattern.Selected(i) = True) Then
            If (Compare_String = "*** Left-Needle ***") Then
                If (RightNeedle.value = True) Then
                    'MsgBox ("Please choose the corrected neelde before adding the element!")
                    Different_Part = True
                    Exit For
                Else
                    LeftNeedle_No = 1
                End If
            ElseIf (Compare_String = "*** Right-Needle ***") Then
                If (LeftNeedle.value = True) Then
                    'MsgBox ("Please choose the corrected neelde before adding the element!")
                    Different_Part = True
                    Exit For
                Else
                    RightNeedle_No = 1
                End If
            Else
                Different_Part = False
            End If
        End If
        
        If (lstPattern.List(i) = "*** Left-Needle ***") Or (lstPattern.List(i) = "*** Right-Needle ***") Then
            Compare_String = lstPattern.List(i)
        End If
    Next i
End Function

Private Sub Expand_Click()
    Dim CountList, no, counter, currentIndex As Integer
    Dim Index As Integer
    Dim words() As String
    Dim Start, EndArray As Integer
        
        
    CountList = lstPattern.ListCount
    counter = 0
    EndArray = 0
    ExpandWithDrawSpeed = 0
        
    counter = SelectedItems(CountList, currentIndex)
    lstPattern.ListIndex = currentIndex
        
    If counter > 1 Then
        Call MsgBox("Only one node will be allowed to expand!", vbOKOnly, "Expanding")
        Exit Sub
    End If
        
    'Not allow the user to do array.
    If xRepeatNum.Text = 0 Or yRepeatNum.Text = 0 Then
        MsgBox ("Please check No of Rows or Columns")
        Exit Sub
    End If
        
    Index = lstPattern.ListIndex
            
    words() = Split(lstPattern.List(lstPattern.ListIndex), "(")
        
    If words(0) = "dot" Then
        ExpandWithDrawSpeed = WithDrawSpeed(lstPattern.ListIndex)
    End If
        
    'No (pot dot and pot line)
    'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5 _
    '    And xDev.Text <> 0 And yDev.Text <> 0 And xRepeatNum.Text > 0 And yRepeatNum.Text > 0) Then
    If (NodeType.ListIndex = 2 And xDev.Text <> 0 And yDev.Text <> 0 And xRepeatNum.Text > 0 And yRepeatNum.Text > 0) Then
         
        If StrComp(words(0), "dotArray", vbTextCompare) = 0 Then
            'Old
            'Call lstPattern.AddItem("StartDotArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & ExpandWithDrawSpeed & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", Index)
            Call lstPattern.AddItem("StartDotArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & ExpandWithDrawSpeed & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", Index)
            Call lstPattern.RemoveItem(Index + 1)
        'No (pot dot and pot line)
        'ElseIf StrComp(words(0), "dotPottingArray", vbTextCompare) = 0 Then
        '    'Old
        '    'Call lstPattern.AddItem("StartDotPottingArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & Expand + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", Index)
        '    Call lstPattern.AddItem("StartDotPottingArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & Expand + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", Index)
        '    Call lstPattern.RemoveItem(Index + 1)
        'ElseIf StrComp(words(0), "linePottingArray", vbTextCompare) = 0 Then
        '    'Old
        '    'Call lstPattern.AddItem("StartLinePottingArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", Index)
        '    Call lstPattern.AddItem("StartLinePottingArray(x=" & ExpandX + referenceX & ", y=" & ExpandY + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", Index)
        '    Call lstPattern.RemoveItem(Index + 1)
        End If
            
        ClickExpand = True
            
        WriteArrayTextLine (Index)
            
        ClickExpand = False
        EndArray = D_A_rows * D_A_columns
        Call lstPattern.AddItem("EndArray", Index + EndArray + 1)
    End If
        
    FirstLineSelect = False
    selectNodeIndex = lstPattern.ListIndex
    lstPattern.Selected(lstPattern.ListIndex) = True
    Expand.Enabled = False

    Click = 0
End Sub

Private Function WithDrawSpeed(ByVal Index As Integer) As Integer
    Dim withdrawalSpeed() As String   'To split the withdrawspeed
    Dim StringLenght As Integer
    Dim ReadStringValue As String
    Dim flag As Boolean
    
    flag = False
    ReadStringValue = ""
    
    withdrawalSpeed() = Split(lstPattern.List(Index), ";")
    
    StringLenght = Len(withdrawalSpeed(4))
    For Start = 1 To StringLenght
        If flag = True Then
            ReadStringValue = ReadStringValue & Mid(withdrawalSpeed(4), Start, 1)
        End If
        If Mid(withdrawalSpeed(4), Start, 1) = "=" Then
            flag = True
        End If
    Next Start
    
    WithDrawSpeed = Val(ReadStringValue)
    
End Function

Private Sub WriteArrayTextLine(ByVal Index As Integer)
    Dim r As Integer
    Dim C As Integer
            
    D_A_row_pitch = yDev.Text * 1000
    D_A_column_pitch = xDev.Text * 1000
    D_A_rows = yRepeatNum.Text
    D_A_columns = xRepeatNum.Text
    add_row_pitch = 0
    add_column_pitch = 0
    'yRepeatNum.Text = 1
    'xRepeatNum.Text = 1
            
    For r = 1 To D_A_rows
        For C = 1 To D_A_columns
            If ClickExpand = True Then
                Call lstPattern.AddItem(processAddNode, Index + 1)
            Else
                Call lstPattern.AddItem(ModifyArrayElements, Index + 1)
            End If
            If (r Mod 2 = 0) And (C <> D_A_columns) Then
                add_column_pitch = add_column_pitch - D_A_column_pitch
            ElseIf (C <> D_A_columns) Then
                add_column_pitch = add_column_pitch + D_A_column_pitch
            End If
            Index = Index + 1
        Next
        add_row_pitch = add_row_pitch + D_A_row_pitch
    Next
    
    add_row_pitch = 0
    add_column_pitch = 0
    'yRepeatNum.Text = 1
    'xRepeatNum.Text = 1
End Sub

Private Sub addPt_Click()           'Xu Long's array pattern
    addPtClicked
End Sub

Private Sub addPtClicked()
    Dim CountList, Count, currentIndex As Integer
    
    Count = 0
    CountList = lstPattern.ListCount
    
    Count = SelectedItems(CountList, currentIndex)
    lstPattern.ListIndex = currentIndex
    
    If (LeftNeedle.value = True) And (VisionTeach.value = 1) Then
        If (reference_ZHigh = False) Then
            MsgBox "Teach the board high first before adding the position"
            Exit Sub
        End If
    ElseIf (RightNeedle.value = True) And (VisionTeach.value = 1) Then
        If (reference_R_ZHigh = False) Then
            MsgBox "Teach the board high first before adding the position"
            Exit Sub
        End If
    End If
    
    If (Count > 1) Then
        Call MsgBox("Can't select multiple line!", vbOKOnly, "Notice")
        Exit Sub
    ElseIf (Count = 1) Then
        If (Different_Part = True) Then
            MsgBox ("Please choose the corrected neelde before adding the element!")
            Exit Sub
        Else
            If (Trim(NodeType.List(NodeType.ListIndex)) <> "Part Array") Then
                If ((lstPattern.SelCount = 1) And (lstPattern.ListIndex = 0)) Then
                    LeftNeedle_No = 0
                    RightNeedle_No = 0
                End If
            
                If (LeftNeedle.value = True) And (LeftNeedle_No = 0) Then
                    Call lstPattern.AddItem("*** Left-Needle ***", lstPattern.ListIndex)
                
                    lstPattern.ListIndex = lstPattern.ListIndex + 1
        
                    LeftNeedle_No = LeftNeedle_No + 1
                ElseIf (RightNeedle.value = True) And (RightNeedle_No = 0) Then
                    Call lstPattern.AddItem("*** Right-Needle ***", lstPattern.ListIndex)
                
                    lstPattern.ListIndex = lstPattern.ListIndex + 1
            
                    RightNeedle_No = RightNeedle_No + 1
                End If
            End If
        End If
    Else
        If (LeftNeedle.value = False) And (RightNeedle.value = False) Then
            'If we set the default as left valve, the following message will not be come out.
            MsgBox "Please choose the needle mode first before teachnig the position."
            Exit Sub
        ElseIf (LeftNeedle.value = True) And (LeftNeedle_No = 0) Then
            If (Trim(NodeType.List(NodeType.ListIndex)) <> "Part Array") Then
                'Add the text string for Left-Needle.
                If (lstPattern.SelCount = 0) Then
                    Call lstPattern.AddItem("*** Left-Needle ***", lstPattern.ListCount)
                End If
                lstPattern.ListIndex = lstPattern.ListIndex + 1
        
                LeftNeedle_No = LeftNeedle_No + 1
            End If
        ElseIf (RightNeedle.value = True) And (RightNeedle_No = 0) Then
            If (Trim(NodeType.List(NodeType.ListIndex)) <> "Part Array") Then
                'Add the text string for Right-Needle.
                If (lstPattern.SelCount = 0) Then
                    Call lstPattern.AddItem("*** Right-Needle ***", lstPattern.ListCount)
                End If
                lstPattern.ListIndex = lstPattern.ListIndex + 1
        
                RightNeedle_No = RightNeedle_No + 1
            End If
        End If
    End If
    
    'Refresh the listbox after clicking de-highlight
    If lstPattern.SelCount = 0 Then
        lstPattern.ListIndex = 0
        SingleClick = False
        FirstLineSelect = False
    End If
    
    If (lstPattern.ListIndex = 0) Then
        If FirstLineSelect = False Then
            lstPattern.AddItem (processAddNode)
        Else
            If NodeType.ListIndex = 20 Then
                Call lstPattern.AddItem(processAddNode)
            Else
                Call lstPattern.AddItem(processAddNode, lstPattern.ListIndex)
                If (processAddNode <> "") Then
                    lstPattern.ListIndex = lstPattern.ListIndex + 1
                End If
            End If
        End If
        
        If lstPattern.List(lstPattern.ListCount - 1) = "" Then
            lstPattern.RemoveItem (lstPattern.ListCount - 1)
        End If
        
        If lstPattern.ListCount > 2 And lstPattern.List(lstPattern.ListCount - 2) = "" Then
            lstPattern.RemoveItem (lstPattern.ListCount - 2)
        End If
        
        If lstPattern.List(lstPattern.ListIndex) = "" Then
            'Not to delete the second time
            If (lstPattern.ListCount <> 0) Then
                lstPattern.RemoveItem (lstPattern.ListIndex)
            End If
        End If
        
        If lstPattern.ListCount > 0 Then
            lstPattern.TopIndex = lstPattern.ListCount - 1
        Else
            lstPattern.TopIndex = 0
        End If
    Else
        clearLockedNode
        
        Call lstPattern.AddItem(processAddNode, lstPattern.ListIndex)
        lstPattern.ListIndex = lstPattern.ListIndex + 1
                
        If lstPattern.List(lstPattern.ListIndex - 1) = "" Then
            lstPattern.RemoveItem (lstPattern.ListIndex - 1)
        End If
        If lstPattern.ListIndex > 0 Then
            lstPattern.TopIndex = lstPattern.ListIndex - 1
        End If
    End If
    
    'Check for loading "old file"
    If (fileDirty = False) Then
        editorForm.Caption = "Profile Editor"
    End If
    
    fileDirty = True
    
    'Because the listBox doesn't know there is one line still heighlight
    'So, we should do the focus again       'XW
    If (lstPattern.SelCount = 1) Then
        If lstPattern.ListIndex <> 0 Then
            FirstLineSelect = False
        Else
            FirstLineSelect = True
        End If
        selectNodeIndex = lstPattern.ListIndex
    End If
    Click = 0
End Sub

Private Sub cmdGoffset_Click()
    Dim i As Integer
    
    If (GlobalOffsetX <> 0) Or (GlobalOffsetY <> 0) Or (GlobalOffsetZ <> 0) Then
        For i = 0 To (lstPattern.ListCount - 1)
            If (i = lstPattern.ListCount) Then
                Exit Sub
            End If
        
            Call GettingPosition(lstPattern.List(i))
            Call GlobalOffsetSpeed(i)
            
            If (AddingElement = True) Then
                lstPattern.RemoveItem (i + 1)
                AddingElement = False
            End If
        Next i
    End If
    FlagGlobalOffset = False
End Sub

Private Sub cmdGDsiTimeSpeed_Click()
    Dim i As Integer
    
    
    For i = 0 To (lstPattern.ListCount - 1)
        If (i = lstPattern.ListCount) Then
            Exit Sub
        End If
               
        Call GlobalOffsetSpeed(i)
        
        If (AddingElement = True) Then
            lstPattern.RemoveItem (i + 1)
            AddingElement = False
        End If
    Next i
    
    FlagGlobalTimeSpeed = False
End Sub

Private Sub GlobalOffsetSpeed(ByVal LineIndex As Integer)
    Dim words() As String, elements() As String, NodeType() As String
    
    elements() = Split(lstPattern.List(LineIndex), "=")
    words() = Split(lstPattern.List(LineIndex), ";")
    NodeType() = Split(lstPattern.List(LineIndex), "(")
    
    Select Case NodeType(0)
        Case "reference"
            If (FlagGlobalOffset = True) Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                AddingElement = True
            End If
        Case "       arcPoint"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ")", LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                AddingElement = True
            End If
        Case "dot", "   dot", "dotArray", "StartDotArray"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalDispenseTime <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & "; " & Format(GlobalDispenseTime, "####0.000") & ";" & words(6) & ";" & words(7), LineIndex)
                    Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & "; " & Format(GlobalDispenseTime, "####0.000") & ";" & words(6) & ";" & words(7) & ";" & words(8), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "dotPotting", "   dotPotting", "dotPottingArray", "StartDotPottingArray"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalDispenseTime <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & "; " & Format(GlobalDispenseTime, "####0.000") & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9), LineIndex)
                    Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & "; " & Format(GlobalDispenseTime, "####0.000") & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "linePotting", "   linePotting", "linePottingArray", "StartLinePottingArray"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11) & ";" & words(12), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalTravelSpeed <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11), LineIndex)
                    Call lstPattern.AddItem(words(0) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11) & ";" & words(12), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "lineStart"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2), LineIndex)
                AddingElement = True
            End If
        Case "arcStart"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2), LineIndex)
                AddingElement = True
            End If
        Case "lineEnd", "arcEnd"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalTravelSpeed <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6), LineIndex)
                    Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "   linksLinePoint", "   linksArcEnd"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalTravelSpeed <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2), LineIndex)
                    Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2) & ";" & words(3), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "   linksArcRestart", "   linksArcStart"
            If (FlagGlobalOffset = True) Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalTravelSpeed <> 0) Then
                    'Old (No angle)
                    'Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2), LineIndex)
                    Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2) & ";" & words(3), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "rectC1"
            If (FlagGlobalOffset = True) Then
                'Old
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                'Add no spray area
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4), LineIndex)
                AddingElement = True
            End If
        Case "   rectC2"
            If (FlagGlobalOffset = True) Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                AddingElement = True
            End If
        Case "rectC3"
            If (FlagGlobalOffset = True) Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                AddingElement = True
            ElseIf (FlagGlobalTimeSpeed = True) Then
                If (GlobalTravelSpeed <> 0) Then
                    Call lstPattern.AddItem(words(0) & "; " & "sp=" & GlobalTravelSpeed & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                    AddingElement = True
                Else
                    AddingElement = False
                End If
            End If
        Case "repeat"
            If (FlagGlobalOffset = True) Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + CLng(ModifyOffsetX) & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + CLng(ModifyOffsetY) & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + CLng(ModifyOffsetZ) & ";" & words(1), LineIndex)
                AddingElement = True
            End If
    End Select
End Sub

Private Sub cmdExecuteButton_Click()
    executionForm.Show (vbModal)
End Sub

Private Sub clearLockedNode()
    updateDispensePt = True
    decideDispensePt.Caption = "Teach On"
    
    updateXDevOnly = False
    decideXYDev.Caption = "Teach Off"
    
    updateYDevOnly = False
    
    updateMoveHeightOnly = False
    decideMoveHeight.Caption = "Teach Off"
    
    updateWithDrawalHeightOnly = False
    decideWithdrawalHeight.Caption = "Teach Off"
    
    
    decideWithdrawalHeight.Refresh
    decideMoveHeight.Refresh
    decideXYDev.Refresh
    decideDispensePt.Refresh
End Sub

Private Sub decideDispensePt_Click()
    If decideWithdrawalHeight.Caption = "Teach On" Or decideMoveHeight.Caption = "Teach On" Or decideXYDev.Caption = "Teach On" Then
        decideWithdrawalHeight.Caption = "Teach Off"
        decideMoveHeight.Caption = "Teach Off"
        decideXYDev.Caption = "Teach Off"
        decideDispensePt.Caption = "Teach On"
        
        decideWithdrawalHeight.Refresh
        decideMoveHeight.Refresh
        decideXYDev.Refresh
        decideDispensePt.Refresh

        updateWithDrawalHeightOnly = False
        updateMoveHeightOnly = False
        updateXDevOnly = False
        updateYDevOnly = False
        updateDispensePt = True
    Else
        If decideDispensePt.Caption = "Teach Off" Then
            updateDispensePt = True
            decideDispensePt.Caption = "Teach On"
        Else
            updateDispensePt = False
            decideDispensePt.Caption = "Teach Off"
        End If
    End If
End Sub

Private Sub decideMoveHeight_Click()
    
    If decideWithdrawalHeight.Caption = "Teach On" Or decideDispensePt.Caption = "Teach On" Or decideXYDev.Caption = "Teach On" Then
        decideWithdrawalHeight.Caption = "Teach Off"
        decideMoveHeight.Caption = "Teach On"
        decideXYDev.Caption = "Teach Off"
        decideDispensePt.Caption = "Teach Off"
        
        decideWithdrawalHeight.Refresh
        decideMoveHeight.Refresh
        decideXYDev.Refresh
        decideDispensePt.Refresh
        
        updateWithDrawalHeightOnly = False
        updateMoveHeightOnly = True
        updateXDevOnly = False
        updateYDevOnly = False
        updateDispensePt = False
    Else
        If decideMoveHeight.Caption = "Teach Off" Then
            updateMoveHeightOnly = True
            decideMoveHeight.Caption = "Teach On"
        Else
            updateMoveHeightOnly = False
            decideMoveHeight.Caption = "Teach Off"
            decideDispensePt_Click                      'XW
            decideDispensePt.Refresh                    'XW
        End If
    End If
End Sub

Private Sub decideWithdrawalHeight_Click()
    
    If decideDispensePt.Caption = "Teach On" Or decideMoveHeight.Caption = "Teach On" Or decideXYDev.Caption = "Teach On" Then
        decideWithdrawalHeight.Caption = "Teach On"
        decideMoveHeight.Caption = "Teach Off"
        decideXYDev.Caption = "Teach Off"
        decideDispensePt.Caption = "Teach Off"
        
        decideWithdrawalHeight.Refresh
        decideMoveHeight.Refresh
        decideXYDev.Refresh
        decideDispensePt.Refresh
        
        updateWithDrawalHeightOnly = True
        updateMoveHeightOnly = False
        updateXDevOnly = False
        updateYDevOnly = False
        updateDispensePt = False
    Else
        If decideWithdrawalHeight.Caption = "Teach Off" Then
            updateWithDrawalHeightOnly = True
            decideWithdrawalHeight.Caption = "Teach On"
        Else
            updateWithDrawalHeightOnly = False
            decideWithdrawalHeight.Caption = "Teach Off"
            decideDispensePt_Click                          'XW
            decideDispensePt.Refresh                        'XW
        End If
    End If
End Sub

Private Sub decideXYDev_Click()
    
    If decideWithdrawalHeight.Caption = "Teach On" Or decideMoveHeight.Caption = "Teach On" Or decideDispensePt.Caption = "Teach On" Then
        decideWithdrawalHeight.Caption = "Teach Off"
        decideMoveHeight.Caption = "Teach Off"
        decideXYDev.Caption = "Teach On"
        decideDispensePt.Caption = "Teach Off"
        
        decideWithdrawalHeight.Refresh
        decideMoveHeight.Refresh
        decideXYDev.Refresh
        decideDispensePt.Refresh
        
        updateWithDrawalHeightOnly = False
        updateMoveHeightOnly = False
        updateXDevOnly = True
        updateYDevOnly = True
        updateDispensePt = False
    Else
        If decideXYDev.Caption = "Teach Off" Then
            updateXDevOnly = True
            updateYDevOnly = True
            decideXYDev.Caption = "Teach On"
        Else
            updateXDevOnly = False
            updateYDevOnly = False
            decideXYDev.Caption = "Teach Off"
            decideDispensePt_Click                  'XW
            decideDispensePt.Refresh                'XW
        End If
    End If
End Sub

Private Sub deletePt_Click()
     SetFocusTimer.Enabled = False
    If (lstPattern.SelCount = 0) Then
        NodeError.Show (vbModal)
    Else
        Dim Count, no As Integer                        'Just for looping (main loop)
        Dim CheckExpandNumber As Integer                'Just for looping
        Dim DeleteNum, AddNum As Integer                'Count the deleted intems and remain items
        Dim SelectedNo, CountSelected As Integer        'For checking selected items
        Dim totalArray As Integer                       'Total sub array
        Dim DoArray() As String                         'Split the row and colum
        Dim i As Integer                                'For list index
        Dim StartPoint, currentIndex As Integer         'Start deleted point
        Dim DeleteArray()
        Dim Name As String
        Dim words() As String
        
        clearLockedNode
        
        Name = ""
        AddNum = 0
        ReDim DeleteArray(0)
        DeleteNum = 0
        CountSelected = 0
        i = lstPattern.ListIndex
        Count = lstPattern.ListCount
        
        For SelectedNo = 0 To Count - 1
            If lstPattern.Selected(SelectedNo) = True Then
                CountSelected = CountSelected + 1
                currentIndex = SelectedNo
            End If
        Next SelectedNo
        
        If CountSelected = 1 Then
            i = currentIndex
            words() = Split(lstPattern.List(i), "(")
            
            If (words(0) = "*** Left-Needle ***") Or (words(0) = "*** Right-Needle ***") Then
                If (lstPattern.List(i + 1) = "*** Left-Needle ***") Or (lstPattern.List(i + 1) = "*** Right-Needle ***") Or (lstPattern.List(i + 1) = "") Then
                    lstPattern.RemoveItem (i)
                'Check for the last line
                ElseIf ((i + 1) = lstPattern.ListCount) Then
                    lstPattern.RemoveItem (i)
                Else
                    MsgBox "Not allow to delete it!"
                End If
                Exit Sub
            End If
                            
            If StrComp(words(0), "reference", vbTextCompare) = 0 Then
                referenceX = 0
                referenceY = 0
                referenceZ = 0
                referenceSet = False
            End If
            
            'Reset the selection of the valve (XW) => Not allow to delete them for spray system
            'If (words(0) = "*** Left-Needle ***") Then
            '    LeftNeedle_No = 0
            'ElseIf (words(0) = "*** Right-Needle ***") Then
            '    RightNeedle_No = 0
            'End If
            
            If (words(0) = "StartDotArray") Or (words(0) = "StartDotPottingArray") Or (words(0) = "StartLinePottingArray") Then
                        
                DoArray() = Split(lstPattern.List(i), ";")
                Rows = 1
                Columns = 1
                totalArray = 1
                    
                If Right(DoArray(1), 1) <> "" Then
                    Columns = StringValue(DoArray(1))
                End If
                If Right(DoArray(2), 1) <> "" Then
                    Rows = StringValue(DoArray(2))
                End If
                
                totalArray = Columns * Rows
                        
                CheckExpandNumber = CheckSubArray(i)
                        
                If (CheckExpandNumber <> 0) Then
                    If CheckExpandNumber = totalArray Then
                        'Adding "+ 1" is to remove "EndArray"
                        For StartPoint = i To i + totalArray + 1
                            If lstPattern.List(i) <> "" Then
                                lstPattern.RemoveItem (i)
                            End If
                        Next StartPoint
                    Else
                        For StartPoint = i To i + CheckExpandNumber + 1
                            If lstPattern.List(i) <> "" Then
                                lstPattern.RemoveItem (i)
                            End If
                        Next StartPoint
                    End If
                Else
                    For StartPoint = i To i + 1
                        If lstPattern.List(i) <> "" Then
                            lstPattern.RemoveItem (i)
                        End If
                    Next StartPoint
                End If
            Else
                If lstPattern.List(i) <> "EndArray" Then
                    If i = 0 Then
                        lstPattern.RemoveItem (i)
                        Expand.Enabled = False
                    Else
                        words() = Split(lstPattern.List(i - 1), "(")
                        If ((words(0) = "StartDotArray") Or (words(0) = "StartDotPottingArray") Or (words(0) = "StartLinePottingArray")) And (lstPattern.List(i + 1) = "EndArray") Then
                            MsgBox ("This gets the last array's element. The whole procedure of array will be deleted automatically!")
                            'Deleting the whole procedure of array when it gets lasting array's element.
                            For StartPoint = (i - 1) To (i + 1)
                                If lstPattern.List(i - 1) <> "" Then
                                    lstPattern.RemoveItem (i - 1)
                                End If
                            Next StartPoint
                        Else
                            lstPattern.RemoveItem (i)
                        End If
                    End If
                End If
            End If
            
            Expand.Enabled = False
        Else
            For no = 0 To Count - 1
                words() = Split(lstPattern.List(no), "(")
                Name = words(0)
                
                'Reset the selection of the valve (XW)
                'If (words(0) = "*** Left-Needle ***") Then
                '    LeftNeedle_No = 0
                'ElseIf (words(0) = "*** Right-Needle ***") Then
                '    RightNeedle_No = 0
                'End If
                
                If Not ((words(0) = "*** Left-Needle ***") Or (words(0) = "*** Right-Needle ***")) Then
                    If lstPattern.Selected(no) = True Then
                        If StrComp(words(0), "reference", vbTextCompare) = 0 Then
                            referenceX = 0
                            referenceY = 0
                            referenceZ = 0
                            referenceSet = False
                        End If
                    
                        If (Name = "StartDotArray") Or (Name = "StartDotPottingArray") Or (Name = "StartLinePottingArray") Then
                        
                            DoArray() = Split(lstPattern.List(no), ";")
                            Rows = 1
                            Columns = 1
                            totalArray = 1
                    
                            If Right(DoArray(1), 1) <> "" Then
                                Columns = StringValue(DoArray(1))
                            End If
                            If Right(DoArray(2), 1) <> "" Then
                                Rows = StringValue(DoArray(2))
                            End If
                
                            totalArray = Columns * Rows
                        
                            CheckExpandNumber = CheckSubArray(no)
                        
                            If (CheckExpandNumber <> 0) Then
                                If CheckExpandNumber = totalArray Then
                                    For StartPoint = no To no + totalArray + 1
                                        If lstPattern.List(no) <> "" Then
                                            lstPattern.RemoveItem (no)
                                            DeleteNum = DeleteNum + 1
                                        End If
                                    Next StartPoint
                                    no = no - 1
                                    Count = Count - DeleteNum       'Resize the listcount
                                    DeleteNum = 0
                                Else
                                    For StartPoint = no To no + CheckExpandNumber + 1
                                        If lstPattern.List(no) <> "" Then
                                            lstPattern.RemoveItem (no)
                                            DeleteNum = DeleteNum + 1
                                        End If
                                    Next StartPoint
                                    no = no - 1
                                    Count = Count - DeleteNum       'Resize the listcount
                                    DeleteNum = 0
                                End If
                            Else
                                For StartPoint = no To no + 1
                                    lstPattern.RemoveItem (no)
                                Next StartPoint
                                '"- 2" is because system remove "StartArray" and "End Array"
                                no = no - 1
                                Count = Count - 2       'Resize the listcount
                            End If
                        Else
                            If lstPattern.List(no) <> "EndArray" Then
                                If no = 0 Then
                                    lstPattern.RemoveItem (no)
                                    Count = Count - 1      'Resize the listcount
                                Else
                                    words() = Split(lstPattern.List(no - 1), "(")
                                    If ((words(0) = "StartDotArray") Or (words(0) = "StartDotPottingArray") Or (words(0) = "StartLinePottingArray")) And (lstPattern.List(no + 1) = "EndArray") Then
                                        MsgBox ("This gets the last array's element. The whole procedure of array will be deleted automatically!")
                                        'Deleting the whole procedure of array when it gets lasting array's element.
                                        For StartPoint = (no - 1) To (no + 1)
                                            If lstPattern.List(no - 1) <> "" Then
                                                lstPattern.RemoveItem (no - 1)
                                                DeleteNum = DeleteNum + 1
                                            End If
                                        Next StartPoint
                                        Count = Count - DeleteNum
                                        AddNum = AddNum - 1
                                        DeleteNum = 0
                                    Else
                                        lstPattern.RemoveItem (no)
                                        Count = Count - 1      'Resize the listcount
                                    End If
                                End If
                                no = no - 1
                            Else
                                ReDim Preserve DeleteArray(AddNum)
                                DeleteArray(AddNum) = lstPattern.List(no)
                                AddNum = AddNum + 1
                            End If
                        End If
                    Else
                        ReDim Preserve DeleteArray(AddNum)
                        DeleteArray(AddNum) = lstPattern.List(no)
                        AddNum = AddNum + 1
                    End If
                Else
                    If lstPattern.Selected(no) = True Then
                        If (lstPattern.List(no + 1) = "*** Left-Needle ***") Or (lstPattern.List(no + 1) = "*** Right-Needle ***") Or (lstPattern.List(no + 1) = "") Then
                            lstPattern.RemoveItem (no)
                            Count = Count - 1
                        'Check for the last line
                        ElseIf ((no + 1) = lstPattern.ListCount) Then
                            lstPattern.RemoveItem (no)
                            Count = Count - 1
                        Else
                            MsgBox "Not allow to delete it!"
                            Exit For
                        End If
                    Else
                        ReDim Preserve DeleteArray(AddNum)
                        DeleteArray(AddNum) = lstPattern.List(no)
                        AddNum = AddNum + 1
                    End If
                End If
                
                If (DeleteNum + AddNum) = Count Then
                    Exit For
                End If
            Next no
            
            For no = 0 To AddNum - 1
                lstPattern.List(no) = DeleteArray(no)
            Next no
        End If
        
        FirstLineSelect = False
        selectNodeIndex = 0
        lstPattern.Refresh
        editorForm.Caption = "Profile Editor"         'XW
        FocusLstBox.Visible = False                   'XW
    End If
    fileDirty = True
    SetFocusTimer.Enabled = True
End Sub

Private Sub dispenseptx_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtX.Text, editorForm.XLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub dispensepty_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtY.Text, editorForm.YLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub dispenseptz_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtZ.Text, editorForm.dispensePtZLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub dispenseTime_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispenseTime.Text, editorForm.DispenseTimeLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(dispenseTime.Text) < 0) Then
            MsgBox "There is no negative value for dispense time."
            dispenseTime.Text = ""
            cancel = True
        End If
    End If
End Sub

Private Sub displayCoOrdsTimer_Timer()
    displayCoOrds
    
    'Check and show tower light
    Tower_Light
End Sub

Private Sub fileNew_Click()
    SetFocusTimer.Enabled = False
    proceedWithAction = True

    If fileDirty = True Then
        fileNotSavedForm.Show (vbModal)
    End If
    
    If proceedWithAction = True Then
        NodeType.ListIndex = 0              'XW
        FocusLstBox.Visible = False         'XW
        lstPattern.Clear
        selectNodeIndex = 0
        'Origin (NYP)
        'editorForm.Caption = "Epoxy Editor"
        editorForm.Caption = "Profile Editor"
        SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
        SystemMoveHeight = SystemMoveHeight * (-1)  'XW
        systemTrackMoveHeight = 0
        initializeInputParams
    
        'initializeNodeTypeItems
    
        disableAllInputParams
    
        enableInputParams
    
        clearLockedNode
    
        referenceX = 0
        referenceY = 0
        referenceZ = 0
        referenceSet = False

        referencePtDone = False
        readyStatus = True
        fileDirty = False
        doingModify = False
        Expand.Enabled = False       'Disable the button
        LeftNeedle.value = True      'Set as default (reinitialize)
        ForLeftNeedle
        LeftNeedle_No = 0
        RightNeedle_No = 0
        
        cmdReferenceHigh.Enabled = True     'Enable for teaching z_high again
        reference_ZHigh = False
        Z_High = 0
        reference_R_ZHigh = False
        R_Z_High = 0
    End If
    SetFocusTimer.Enabled = True
End Sub

Private Sub Form_Unload(cancel As Integer)

    If fileDirty = True Then
        fileSaveForm.Show (vbModal)
    End If

    Dim hWnd, hWnd2 As Long
    Dim ReadValue As Long, Tower_Light_Value As Long, DriverXYZ As Long

    hWnd = FindWindow(vbNullString, "Desktop Setup Panel")
    hWnd2 = FindWindow(vbNullString, "File Load")

    If App.PrevInstance <> True And hWnd = 0 And hWnd2 = 0 Then
        If loginsuccessful = False Or mCancel = True Then 'NNO
                Exit Sub
            Else
            'If Get_User_Name = "Techno Digm" Then
            
            PicImage.Enabled = False
            Close_AllTimer
            
            
            Call Turn_Off_LightIntensity            '@$K
            If mscomLighIntensity.PortOpen = True Then
                mscomLighIntensity.PortOpen = False
            End If
                
            '''''''''''''''''''''''''''''''''''''
            '   Alway put the valve in solvent  '
            '''''''''''''''''''''''''''''''''''''
            'Move z_axis to system move height
            Call setSpeed(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50"))
            returncode = P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, SystemMoveHeight, 0)
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
            Loop
            
'            'Right-needle will be gone up.
'            checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
'            ReadValue = ReadValue And &HFEFF
'            checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
    
            'Left-needle will be gone up.
            checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HF7FF
            checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
                
            Call Sleep(0.5)
                
            'Close "tilting valve"
            checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
            ReadValue = ReadValue And &HFEFF
            checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
            
            'Move to the original position (U_axis)
            Rotation_U (0)
            
            Leftslider_go_up
            
            Call setSpeed(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "100"))
            
            '@$K
            If GetStringSetting("EpoxyDispenser", "Setup", "EnableSolventPosition", "0") = "1" Then
                'Move to Solvent Position
                PTPToXYZ SolventPosX, SolventPosY, SolventPosZ
            Else
                'Move to System Home Position
                PTPToXYZ systemHomeX, systemHomeY, systemHomeZ
            End If
            
            'Left needles go down
            Leftslider_go_down
            
            'Servo OFF (XW)
            checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, DriverXYZ))
            DriverXYZ = (DriverXYZ And &HF8FF)
            checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, DriverXYZ))
                
            'Disable Red_Light,Yellow_light and Green_Light
            checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
            Tower_Light_Value = Tower_Light_Value And &HF1FF
            checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
            
            Call Sleep(0.2)
            
            WindowFree lstPattern.hWnd
           
            'returncode = P1240FreeContiBuf(0)
            
            VdeCameraLive False
            VdeReleaseVision
            
            Call Sleep(1)
            unInitializeBoard
            Close_PCI1750
        End If
    End If
End Sub

Private Sub Home_Click()

    Dim tempJogSpeed As Integer
    
    tempJogSpeed = jogSpeedSlider.value

    clearLockedNode
    
    fileNew.Enabled = False
    fileLoad.Enabled = False
    fileSave.Enabled = False
    Home.Enabled = False
    translateButton.Enabled = False
    deletePt.Enabled = False
    modifyPt.Enabled = False
    addPt.Enabled = False
    trackButton.Enabled = False
    yPlus.Enabled = False
    yMinus.Enabled = False
    xPlus.Enabled = False
    xMinus.Enabled = False
    zPlus.Enabled = False
    zMinus.Enabled = False
    loadPartArray.Enabled = False
    
    'For safety(XW)
    If (RightNeedle.value = True) And (RotationAngle.Text <> "None") Then
        Move_To_Zero
        RotationAngle.Text = "None"
        Tilt_Off
        Call Tilt_Rotate(0)
    End If
    
    'See Comments on 300505
    'doSoftHome
    If GetStringSetting("EpoxyDispenser", "Setup", "AlwaysRobotHome", "0") = "1" Then
        setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
        
        PTPToXYZ 1000, -1000, -1000
        
        'returncode = P1240MotHome(boardNum, X_axis Or Y_axis Or Z_axis)    'origin
        Call moveToHome         'XW
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
            DoEvents
            'Exit the task that hasn't finished
            If Emergency_Stop = True Then
                busyStatus = False
                Exit Sub
            End If
        Loop
                
        If (systemHomeX <> 0 Or systemHomeY <> 0 Or systemHomeZ <> 0) Then
        
            systemTrackMoveHeight = SystemMoveHeight

            If GetStringSetting("EpoxyDispenser", "Setup", "DirectSoftHome", "0") = "1" Then
    
                'returncode = P1240InitialContiBuf(0, 100)
        
                'Dim contiPathArray As ContiPathData
        
                'contiPathArray.PathType = IPO_L3
                'contiPathArray.EndPoint_1 = systemHomeX - convertToPulses(editorForm.xCoOrd.Text, X_axis) - needleOffsetX
                'contiPathArray.EndPoint_2 = systemHomeY - convertToPulses(editorForm.yCoOrd.Text, Y_axis) - needleOffsetY
                'contiPathArray.EndPoint_3 = systemHomeZ - convertToPulses(editorForm.zCoOrd.Text, Z_axis)
                'returncode = P1240SetContiData(0, contiPathArray, 1)
        
                'contiPathArray.EndPoint_1 = 0
                'contiPathArray.EndPoint_2 = 0
                'contiPathArray.EndPoint_3 = 0
        
                'returncode = P1240SetContiData(0, contiPathArray, 2)
            
                returncode = P1240MotLine(0, X_axis Or Y_axis Or Z_axis, 1, systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ, 0)
                returncode = P1240MotAxisParaSet(boardNum, 0, 0, StartVelocity, convertSpeed(CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")), X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate)
                
                'returncode = P1240StartContiDrive(boardNum, X_axis Or Y_axis Or Z_axis, 0)
    
                Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
                    DoEvents
                    'Exit the task that hasn't finished
                    If Emergency_Stop = True Then
                        busyStatus = False
                        Exit Sub
                    End If
                Loop
        
                'returncode = P1240FreeContiBuf(0)
            Else
                setSpeed (100)
                'Origin
                'PTPToXYZ systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, SystemMoveHeight
                systemTrackMoveHeight = SystemMoveHeight          'XW
                PTPToXYZ systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ
            End If
        End If
    Else
        'Origin
        'PTPToXYZ systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, SystemMoveHeight
        setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
        setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "zDefaultSpeed", "50")))
        
        systemTrackMoveHeight = SystemMoveHeight          'XW
        PTPToXYZ systemHomeX - needleOffsetX, systemHomeY - needleOffsetY, systemHomeZ
    End If

    fileNew.Enabled = True
    fileLoad.Enabled = True
    fileSave.Enabled = True
    Home.Enabled = True
    translateButton.Enabled = True
    deletePt.Enabled = True
    modifyPt.Enabled = True
    addPt.Enabled = True
    trackButton.Enabled = True
    yPlus.Enabled = True
    yMinus.Enabled = True
    xPlus.Enabled = True
    xMinus.Enabled = True
    zPlus.Enabled = True
    zMinus.Enabled = True
    loadPartArray.Enabled = True
       
    jogSpeedSlider.value = tempJogSpeed
    Call setSpeed(jogSpeedSlider.value)         'That is for the changing speed after doing the Home
                                                'XW
End Sub

'Private Sub jogSpeedSlider_Click()
    'setSpeed (jogSpeedSlider.Value - 1)
'End Sub
Private Sub loadPartArray_Click()
    SetFocusTimer.Enabled = False
    repeatPatternFileLoad.Show (vbModal)
    SetFocusTimer.Enabled = True
End Sub

Private Function StringValue(ByVal readstring As String) As Integer
    Dim no As Integer
    Dim StringLenght As Integer
    Dim ReadStringValue As String
    Dim flag As Boolean
    
    flag = False
    ReadStringValue = ""
    StringLenght = Len(readstring)

    For no = 1 To StringLenght
        If flag = True Then
            ReadStringValue = ReadStringValue & Mid(readstring, no, 1)
        End If
        If Mid(readstring, no, 1) = "," Then
            flag = True
        End If
    Next no
    
    StringValue = Val(ReadStringValue)
End Function

Private Sub ClickTimer_Timer()
    Dim Name As String
    Dim ListIndex As Integer
    Dim DoArray() As String
    
    ClickTimer.Enabled = False
    
    If (lstPattern.SelCount <> 0) Then
        ProgramStep.Visible = True
        ProgramStep.Caption = "Program step is " & lstPattern.ListIndex + 1
    Else
        ProgramStep.Visible = False
    End If
    
    'For Add
    If ((Click = 1) Or (Click = 2)) And (lstPattern.SelCount = 1) And (lstPattern.ListIndex = 0) Then
        If (lstPattern.List(lstPattern.ListIndex) = "*** Left-Needle ***") Or (lstPattern.List(lstPattern.ListIndex) = "*** Right-Needle ***") Then
            'If (LeftNeedle.value = True) Then
                LeftNeedle_No = 0
            'ElseIf (RightNeedle.value = True) Then
                RightNeedle_No = 0
            'End If
        End If
    End If
        
    If Click = 1 Then
        Name = lstPattern.List(lstPattern.ListIndex)
        ListIndex = lstPattern.ListIndex
        
        DoArray() = Split(lstPattern.List(lstPattern.ListIndex), "(")
        
        RemoveSingleClick = True
        trackButton_Click
        
        If lstPattern.ListIndex = 0 Then
            If lstPattern.SelCount = 1 Then
                If FirstLineSelect = False Then
                    FirstLineSelect = True
                    SingleClick = True
                    NodeTypeNoChange = True
                    ShowNodeType (Name)
                    FocusLstBox.Visible = True
                    
                    'Set the Focus again after click the listBox first time because
                    ' triggering the ActiveX control will lost the focus
                    lstPattern.SetFocus
                    
                    selectNodeIndex = lstPattern.ListIndex
                    
                    'No pot dot and pot line
                    'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5) _
                    '    And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
                    If (NodeType.ListIndex = 2) And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
              
                        'If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "StartDotPottingArray") And (DoArray(0) <> "StartLinePottingArray") And (DoArray(0) <> "EndArray") Then
                        If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "EndArray") Then
                            Expand.Enabled = True
                        Else
                            Expand.Enabled = False
                        End If
                    Else
                    'For "Modify" button, do check again
                        Expand.Enabled = False
                    End If
                Else
                    If selectNodeIndex = lstPattern.ListIndex Then
                        lstPattern.Selected(lstPattern.ListIndex) = False
                        FirstLineSelect = False
                        NodeTypeNoChange = False
                        FocusLstBox.Visible = False
                        Expand.Enabled = False
                        'Restore the old z_high value
                        'Z_High = Save_Old_zHigh
                    End If
                    selectNodeIndex = lstPattern.ListIndex
                End If
            Else
                FirstLineSelect = False
                NodeTypeNoChange = False
                'FocusLstBox.Visible = False
                selectNodeIndex = lstPattern.ListIndex
                ShowNodeType (Name)
            End If
        Else
            If lstPattern.SelCount = 1 Then
                If selectNodeIndex = lstPattern.ListIndex Then
                    lstPattern.Selected(lstPattern.ListIndex) = False
                    selectNodeIndex = 0
                    NodeTypeNoChange = False
                                                                                                                                FocusLstBox.Visible = False
                    Expand.Enabled = False
                    'Restore the old z_high value
                    'Z_High = Save_Old_zHigh
                Else
                    SingleClick = True
                    FirstLineSelect = False
                    selectNodeIndex = lstPattern.ListIndex
                    NodeTypeNoChange = True
                    ShowNodeType (Name)
                    FocusLstBox.Visible = True
                    
                    'No pot dot and pot line
                    'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5) _
                    '    And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
                    If (NodeType.ListIndex = 2) And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
              
                        'If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "StartDotPottingArray") And (DoArray(0) <> "StartLinePottingArray") And (DoArray(0) <> "EndArray") Then
                        If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "EndArray") Then
                            Expand.Enabled = True
                        Else
                            Expand.Enabled = False
                        End If
                    Else
                    'For "Modify" button, do check again
                        Expand.Enabled = False
                    End If
                End If
            Else
                SingleClick = False
                FirstLineSelect = False
                selectNodeIndex = 0
                NodeTypeNoChange = False
                ShowNodeType (Name)
            End If
        End If
        
        If lstPattern.SelCount = 0 Then
            EnableTextbox
            Click = 0
            FirstLineSelect = False
            SingleClick = False
            FocusLstBox.Visible = False
            ClickTimer.Enabled = True
            Exit Sub
        End If
        
    ElseIf Click = 2 Then
        RemoveSingleClick = False
        Name = lstPattern.List(lstPattern.ListIndex)
           
        DoArray() = Split(lstPattern.List(lstPattern.ListIndex), "(")
        
        If lstPattern.ListIndex = 0 Then
            If lstPattern.SelCount = 1 Then
                If SingleClick = True Then
                    If FirstLineSelect = True Then
                        FirstLineSelect = True
                        selectNodeIndex = lstPattern.ListIndex
                        NodeTypeNoChange = True
                        FocusLstBox.Visible = True
                    Else
                        FirstLineSelect = False
                        NodeTypeNoChange = False
                        FocusLstBox.Visible = False
                    End If
                    
                    trackButton_Click
                    
                    'No pot dot and pot line
                    'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5) _
                    '    And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
                    If (NodeType.ListIndex = 2) And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
              
                        'If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "StartDotPottingArray") And (DoArray(0) <> "StartLinePottingArray") And (DoArray(0) <> "EndArray") Then
                        If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "EndArray") Then
                            Expand.Enabled = True
                        Else
                            Expand.Enabled = False
                        End If
                    Else
                    'For "Modify" button, do check again
                        Expand.Enabled = False
                    End If
                Else
                    If FirstLineSelect = True Then
                        lstPattern.Selected(lstPattern.ListIndex) = False
                        FirstLineSelect = False
                        NodeTypeNoChange = False
                        FocusLstBox.Visible = False
                        Expand.Enabled = False
                        'Restore the old z_high value
                        'Z_High = Save_Old_zHigh
                    Else
                        FirstLineSelect = True
                        selectNodeIndex = lstPattern.ListIndex
                        NodeTypeNoChange = True
                        FocusLstBox.Visible = True
                        lstPattern.SetFocus
                        
                        trackButton_Click
                        
                        'No pot dot and pot line
                        'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5) _
                        '    And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
                        If (NodeType.ListIndex = 2) And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
              
                            'If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "StartDotPottingArray") And (DoArray(0) <> "StartLinePottingArray") And (DoArray(0) <> "EndArray") Then
                            If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "EndArray") Then
                                Expand.Enabled = True
                            Else
                                Expand.Enabled = False
                            End If
                        Else
                            'For "Modify" button, do check again
                            Expand.Enabled = False
                        End If
                    End If
                End If
                SingleClick = False
            Else
                SingleClick = False
                NodeTypeNoChange = False
                FocusLstBox.Visible = False
            End If
        Else
            If lstPattern.SelCount = 1 Then
                If (selectNodeIndex = lstPattern.ListIndex) And (SingleClick = False) Then
                    lstPattern.Selected(lstPattern.ListIndex) = False
                    selectNodeIndex = 0
                    NodeTypeNoChange = False
                    FocusLstBox.Visible = False
                    Expand.Enabled = False
                    'Restore the old z_high value
                    'Z_High = Save_Old_zHigh
                Else
                    SingleClick = False
                    selectNodeIndex = lstPattern.ListIndex
                    FocusLstBox.Visible = True
                    NodeTypeNoChange = True
                    
                    trackButton_Click
                    
                    'No pot dot and pot line
                    'If (NodeType.ListIndex = 2 Or NodeType.ListIndex = 3 Or NodeType.ListIndex = 5) _
                    '    And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
                    If (NodeType.ListIndex = 2) And xDev.Text <> 0 And yDev.Text <> 0 And (xRepeatNum.Text > 1 Or yRepeatNum.Text > 1) Then
              
                        'If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "StarDotPottingArray") And (DoArray(0) <> "StartLinePottingArray") And (DoArray(0) <> "EndArray") Then
                        If (DoArray(0) <> "StartDotArray") And (DoArray(0) <> "EndArray") Then
                            Expand.Enabled = True
                        Else
                            Expand.Enabled = False
                        End If
                    Else
                    'For "Modify" button, do check again
                        Expand.Enabled = False
                    End If
                End If
            Else
                SingleClick = False
                selectNodeIndex = 0
                NodeTypeNoChange = False
                FocusLstBox.Visible = False
            End If
        End If
        
        'To avoid mis-type match of Node Type and ListPattern when we click the 2nd times
        'XW (that is also for "Modify Button")
        If lstPattern.SelCount = 0 Then
            EnableTextbox
            Click = 0
            FirstLineSelect = False
            SingleClick = False
            FocusLstBox.Visible = False
            ClickTimer.Enabled = True
            Exit Sub
        End If
    
        If NodeTypeNoChange = True Then
            ShowNodeType (Name)
        End If
        
    End If
    
    Click = 0
    ClickTimer.Enabled = True
End Sub

Private Sub DoArrayLooping(ByVal Name As String, ByVal ListIndex As Integer)
    Dim counter As Integer          'Count the subarray
    Dim DoArray() As String         'Split the row and colum
    Dim totalArray As Integer       'Total sub array
    Dim ListCount As Integer
    Dim Start As Integer
    Dim ForLoop As Integer
    Dim flag As Boolean
    
    counter = 1
    Start = ListIndex
    ListCount = lstPattern.ListCount
    
    If ((Name = "StartDotArray") Or (Name = "StartDotPottingArray") Or (Name = "StartLinePottingArray")) Then
        
        flag = False
        Start = Start + 1
        ListIndex = ListIndex + 1
        
        DoArray() = Split(lstPattern.List(ListIndex), ";")
        Rows = 1
        Columns = 1
        totalArray = 1
                    
        If Right(DoArray(1), 1) <> "" Then
            Columns = StringValue(DoArray(1))
        End If
        If Right(DoArray(2), 1) <> "" Then
            Rows = StringValue(DoArray(2))
        End If
                
        totalArray = Columns * Rows
                
        Do While (Start <> ListCount)
            If lstPattern.List(Start) <> "EndArray" Then
                counter = counter + 1
            End If
            If lstPattern.List(Start) = "EndArray" Then
                flag = True
                Exit Do
            End If
            Start = Start + 1
            selectNodeIndex = 0
            FirstLineSelect = False
        Loop
        counter = counter - 1
        If counter = 1 Then
            MsgBox ("The system can't find 'EndArray'.")
        Else
            If (flag = True) Then
                If ((counter > totalArray) Or (counter < totalArray) Or (counter = totalArray)) And (counter < ListCount) Then
                    For ForLoop = ListIndex To ListIndex + counter
                        lstPattern.Selected(ForLoop) = True
                    Next ForLoop
                End If
            Else
                MsgBox ("The system can't find 'EndArray'.")
            End If
        End If
    End If
End Sub

Private Sub lstPattern_DblClick()
    If Modify = True Then
        Modify = False
        Exit Sub
    End If
    Click = 0
    Click = Click + 2
    
    '''''''''''''''''''''''''''''''''''''''''
    '   Move to "Modify" button process     '
    '                                       '
    '''''''''''''''''''''''''''''''''''''''''
    ''Save the reference z_high
    'Save_Old_zHigh = Z_High
End Sub

Private Sub lstPattern_Click()
    If Modify = True Then
        Modify = False
        Exit Sub
    End If
    
    Click = Click + 1
    
    '''''''''''''''''''''''''''''''''''''''''
    '   Move to "Modify" button process     '
    '                                       '
    '''''''''''''''''''''''''''''''''''''''''
    ''Save the reference z_high
    'Save_Old_zHigh = Z_High
End Sub

Private Function ShowNodeType(ByVal Name As String)
    'Show Type of Node
    
    Dim words() As String
        
    words() = Split(Name, "(")
    
    If words(0) = "reference" Then
        NodeType.Selected(0) = True
    ElseIf (words(0) = "dot") Or (words(0) = "   dot") Or (words(0) = "dotArray") Or (words(0) = "StartDotArray") Then
        NodeType.Selected(2) = True
    'No pot dot and pot line
    'ElseIf (words(0) = "dotPotting") Or (words(0) = "   dotPotting") Or (words(0) = "dotPottingArray") Or (words(0) = "StartDotPottingArray") Then
    '    NodeType.Selected(3) = True
    'ElseIf (words(0) = "linePotting") Or (words(0) = "   linePotting") Or (words(0) = "linePottingArray") Or (words(0) = "StartLinePottingArray") Then
    '    NodeType.Selected(5) = True
    ElseIf words(0) = "lineStart" Then
        NodeType.Selected(4) = True
    ElseIf words(0) = "lineEnd" Then
        NodeType.Selected(5) = True
    ElseIf words(0) = "arcStart" Then
        NodeType.Selected(7) = True
    ElseIf words(0) = "       arcPoint" Then
        NodeType.Selected(8) = True
    ElseIf words(0) = "arcEnd" Then
        NodeType.Selected(9) = True
    ElseIf words(0) = "   linksLinePoint" Then
        NodeType.Selected(11) = True
    ElseIf words(0) = "   linksArcRestart" Then
        NodeType.Selected(12) = True
    ElseIf words(0) = "   linksArcStart" Then
        NodeType.Selected(13) = True
    ElseIf words(0) = "   linksArcEnd" Then
        NodeType.Selected(14) = True
    ElseIf words(0) = "rectC1" Then
        NodeType.Selected(16) = True
    ElseIf words(0) = "   rectC2" Then
        NodeType.Selected(17) = True
    ElseIf words(0) = "rectC3" Then
        NodeType.Selected(18) = True
    ElseIf words(0) = "repeat" Then
        NodeType.Selected(20) = True
    End If
    
    If (words(0) = "   dot") Or (words(0) = "   dotPotting") Or (words(0) = "   linePotting") Then
        xRepeatNum.Enabled = False
        yRepeatNum.Enabled = False
        xDev.Enabled = False
        yDev.Enabled = False
        SubArray = True
    ElseIf (words(0) = "dot") Or (words(0) = "dotArray") Or (words(0) = "StartDotArray") _
        Or (words(0) = "dotPotting") Or (words(0) = "dotPottingArray") Or (words(0) = "StartDotPottingArray") _
        Or (words(0) = "linePotting") Or (words(0) = "linePottingArray") Or (words(0) = "StartLinePottingArray") Or (words(0) = "repeat") Then
        xRepeatNum.Enabled = True
        yRepeatNum.Enabled = True
        xDev.Enabled = True
        yDev.Enabled = True
        SubArray = False
    End If
    
End Function

Private Sub NextFudStep_Click()
    Dim x1 As Double, y1 As Double
    Dim x2 As Double, y2 As Double
    Dim tempStr As String

    FudicialTimer.Enabled = False
    step = VdeTeachRefPtDlg(VisionDlgNext, s)
    If (step = 1) Then
        CancelFud_Click
        Sleep (0.3)
        TeachFudicialPt_Click
        Exit Sub
    ElseIf (step = 3) Then
        jogSpeedSlider.value = 28       'XW
        prevStep = step
        FudMsgText.Caption = "Search Area for 2nd Fiducial Point, click Next->"
        FudicialTimer.Enabled = True
        Exit Sub
    End If
    
    step = VdeTeachRefPtDlg(VisionDlgNext, s)
    
    If (step = VisionDlgToFinish) Then
        step = prevStep + 1
    ElseIf (step = VisionDlgFinish) Then
        VdeGetRefPtPos x1, y1, x2, y2
        'To get the actual direction
        y1 = y1 * (-1)
        y2 = y2 * (-1)
      
        tempStr = "fudicial(x=" & convertToPulses(x1, X_axis) & ", y=" & convertToPulses(y1, Y_axis) & "; x=" & convertToPulses(x2, X_axis) & ", y=" & convertToPulses(y2, Y_axis) & "; " & Chr(34) & editorForm.Caption & "pat" & Chr(34) & "; " & LightingIntensity.Text & ")"
        
        Call editorForm.lstPattern.AddItem(tempStr, 0)

        CancelFud_Click
        Exit Sub
    End If
    
    jogSpeedSlider.value = 28       'XW
    
    prevStep = step
    'Origin (NYP)
    'FudMsgText.Caption = "Adjust ROI and Search Area for 2nd Fiducial Point, click Next->"
    FudMsgText.Caption = "Search Area for 2nd Fiducial Point, click Next->"
    
    FudicialTimer.Enabled = True
End Sub

'Check the driver's error and do the reset
Private Sub resetTimer_Timer()
    Dim resetX, resetY, resetZ As Long
    resetTimer.Enabled = False
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, resetX, resetY, resetZ, 0))
    resetX = resetX And &H10
    resetY = resetY And &H10
    resetZ = resetZ And &H10
    If (resetX = &H10) Or (resetY = &H10) Or (resetZ = &H10) Then
        'frmReset.Show (vbModal)
        
        MsgBox "Driver Error. Please check the Driver!"
        
        Close_AllTimer
                
        Servo_Off
                    
        Call Sleep(0.05)
                
        Close_TowerLight
                
        PicImage.Enabled = False
                
        WindowFree lstPattern.hWnd
        Call Sleep(0.9)
                
        VdeCameraLive False
        VdeReleaseVision
        Call Sleep(0.6)
        
        ResetDriver
        
        unInitializeBoard
        Close_PCI1750
        End
        
        Exit Sub
    End If
    resetTimer.Enabled = True
End Sub

Private Sub RotationAngle_Click()
    If (RotationAngle.Text <> "None") Then
        'No camera
        If (VisionTeach.value = 1) Then
            MsgBox "'Camera Teach' is not allowed for tilting and rotation."
            RotationAngle.Text = "None"
            Save_Angle = "None"
            Tilt_Off
            Call Tilt_Rotate(0)
            addPt.SetFocus
            Exit Sub
        End If
        
        Tilt_ON
        Call Tilt_Rotate(RotationAngle.Text)
        Save_Angle = ""
        Save_Index = RotationAngle.ListIndex
        Save_Angle = RotationAngle.List(RotationAngle.ListIndex)
    Else
        Move_To_Zero
        Tilt_Off
        'Should not to move the origin position
        Call Tilt_Rotate(0)
        Save_Angle = "None"
    End If
    
    addPt.SetFocus
End Sub

Private Function Tilt_Rotate(ByVal Rot As String)
    If (Rot = "0") Or (Rot = "-360") Then
        Rotation_U (0)
    ElseIf (Rot = "-90") Then
        'angle step is 0.072 (gear ration=10)
        'Rotation_U (1250)
        'angle step is 0.036 (gear ration=20)
        'Rotation_U (2500)
        'angle step is 0.1 (gear ration=18)
        Rotation_U (900)
    ElseIf (Rot = "-180") Then
        'Rotation_U (2500)
        'Rotation_U (5000)
        Rotation_U (1800)
    ElseIf (Rot = "-270") Then
        'Rotation_U (3750)
        'Rotation_U (7500)
        Rotation_U (2700)
    'End If
    ElseIf (Rot = "-45") Then
        'Rotation_U (3750)
        'Rotation_U (7500)
        Rotation_U (450)
    End If
End Function

Private Sub RotationAngle_LostFocus()
    'RotationAngle.List(Save_Index) = Save_Angle
    If (Save_Angle <> "") Then
        RotationAngle.Text = Save_Angle
    End If
End Sub

Private Sub TeachFudicialPt_Click()
    'Origin (NYP)
    'If editorForm.Caption = "Epoxy Editor" Then
    If editorForm.Caption = "Profile Editor" Then
        
        Call fileSave_Click
        
    Else
        lockNonMoveControls
        patFile = editorForm.Caption & "pat"
        NextFudStep.Visible = True
        CancelFud.Visible = True
        TeachFudicialPt.Visible = False
        'FindNeedleOffset.Visible = False
        
        VdeSetRefPtFilename patFile
        step = VdeTeachRefPtDlg(VisionDlgInit, s)
        'Origin (NYP)
        'FudMsgText.Caption = "Adjust ROI and Search Area for 1st Fiducial Point, click Next->"
        FudMsgText.Caption = "Search Area for 1st Fiducial Point, click Next->"
        s = ""
        prevStep = 1
        FudicialTimer.Enabled = True
    End If
End Sub

Private Sub trackButton_Click()
    If (lstPattern.SelCount = 0) Then
        'Commented to give constant tracking
        'NodeError.Show (vbModal)
    Else
        If (lstPattern.List(lstPattern.ListIndex) <> "") Then
            clearLockedNode
            
            Dim words() As String, words2() As String, angle() As String
            Dim patternList As String, Rotation_Angle As String, old_angle As String
            Dim StringLenght As Integer
            
            'Do initialization      'XW
            ExpandX = 0
            ExpandY = 0
            ExpandZ = 0
            
            patternList = lstPattern.List(lstPattern.ListIndex)
            words() = Split(patternList, "(")
            If (patternList <> "EndArray") And (patternList <> "*** Left-Needle ***") And (patternList <> "*** Right-Needle ***") Then
                words2() = Split(words(1), ";")
                angle() = Split(patternList, "=")
            End If
            
            'No angle for "reference", "rectC1", "rectC2" and "rectC3"
            'If (words(0) = "reference") Or (words(0) = "rectC1") Or (words(0) = "   rectC2") Or (words(0) = "rectC3") Then
            '    Rotation_Angle = "None"
            'End If
            
            If (words(0) = "reference") Then
                patternList = "reference(" & words2(0) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
            ElseIf (words(0) = "dot") Or (words(0) = "   dot") Or (words(0) = "dotArray") Or (words(0) = "StartDotArray") Then
                patternList = "dot(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ";" & words2(7) & ")"
                Rotation_Angle = Left(angle(7), Len(angle(7)) - 1)
                
            'No pot dot and pot line
            'ElseIf (words(0) = "dotPotting") Or (words(0) = "   dotPotting") Or (words(0) = "dotPottingArray") Or (words(0) = "StartDotPottingArray") Then
            '    patternList = "potType1(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ";" & words2(7) & ";" & words2(8) & ";" & words2(9) & ")"
            '    Rotation_Angle = Left(Angle(9), Len(Angle(9)) - 1)
                
            'ElseIf (words(0) = "linePotting") Or (words(0) = "   linePotting") Or (words(0) = "linePottingArray") Or (words(0) = "StartLinePottingArray") Then
            '    patternList = "potType2(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ";" & words2(7) & ";" & words2(8) & ";" & words2(9) & ";" & words2(10) & ";" & words2(11) & ")"
            '    Rotation_Angle = Left(Angle(11), Len(Angle(11)) - 1)
                
            ElseIf words(0) = "lineStart" Then
                patternList = "start(" & words2(0) & ";" & words2(1) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                
            ElseIf words(0) = "lineEnd" Then
                'No angle
                'patternList = "end3D(" & words(1)
                patternList = "end3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Rotation_Angle = Left(angle(8), Len(angle(8)) - 1)
            
            ElseIf words(0) = "arcStart" Then
                patternList = "arcStart(" & words2(0) & ";" & words2(1) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                
            ElseIf words(0) = "       arcPoint" Then
                'No angle
                'StringLenght = Len(words(1))
                'patternList = "arcStart(" & Left(words(1), StringLenght - 1) & "; 1.000)"
                patternList = "arcStart(" & words2(0) & "; 1.000)"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
            
            ElseIf words(0) = "arcEnd" Then
                patternList = "arcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Rotation_Angle = Left(angle(8), Len(angle(8)) - 1)
            
            ElseIf words(0) = "   linksLinePoint" Then
                patternList = "line3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
            'No angle
            'ElseIf (words(0) = "   linksArcRestart") Then
            '    patternList = "linksArcStart(" & words(1)
            ElseIf (words(0) = "   linksArcRestart") Or (words(0) = "   linksArcStart") Then
                patternList = "linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
            
            ElseIf (words(0) = "   linksArcEnd") Then
                patternList = "linksArcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
            ElseIf (words(0) = "rectC1") Then
                patternList = "start(" & words2(0) & ";" & words2(1) & ")"
                txtPitch.Text = CStr(Left(angle(4), Len(angle(4)) - 9) / 1000)
               
                If (words2(3) = "1") Then
                    NoSprayArea.value = 1
                Else
                    NoSprayArea.value = 0
                End If
                
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
            ElseIf (words(0) = "   rectC2") Then
                StringLenght = Len(words(1))
                'patternList = "Start(" & Left(words(1), StringLenght - 1) & "; 1.000)"
                patternList = "Start(" & words2(0) & "; 1.000)"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
            ElseIf (words(0) = "rectC3") Then
                'patternList = "end3D(" & words(1)
                patternList = "end3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Rotation_Angle = Left(angle(8), Len(angle(8)) - 1)
                
            ElseIf (words(0) = "repeat") Then
                Rotation_Angle = "None"
                
            ElseIf (patternList = "EndArray") Or (patternList = "*** Left-Needle ***") Or (patternList = "*** Right-Needle ***") Then
                   Exit Sub
            End If
            
            'doTrack (lstPattern.List(lstPattern.ListIndex) & vbNewLine)
            doTrack (patternList & vbNewLine)
            'jogSpeedSlider.value = 50
            'setSpeed (jogSpeedSlider.value - 1)
            If RemoveSingleClick = True Then
                RemoveSingleClick = False
                Exit Sub
            End If

            fileNew.Enabled = False
            fileLoad.Enabled = False
            fileSave.Enabled = False
            Home.Enabled = False
            translateButton.Enabled = False
            deletePt.Enabled = False
            modifyPt.Enabled = False
            addPt.Enabled = False
            trackButton.Enabled = False
            yPlus.Enabled = False
            yMinus.Enabled = False
            xPlus.Enabled = False
            xMinus.Enabled = False
            zPlus.Enabled = False
            zMinus.Enabled = False
         
            systemTrackMoveHeight = SystemMoveHeight
            setSpeed (CLng(GetStringSetting("EpoxyDispenser", "Setup", "xyDefaultSpeed", "50")))
            
            'Move_To_Zero
            
            'Second, do change the slider and rotate U_axis
            If (Different_Part = True) Then
                If (LeftNeedle.value = True) Then
                    RightNeedle_Click
                Else
                    LeftNeedle_Click
                End If
            Else
                Move_To_Zero
            End If
            
            RotationAngle.Text = Rotation_Angle
            If (Rotation_Angle <> "None") Then
                If (RightNeedle.value = True) Then               '@$K1
                    If (words(0) = "lineStart") Or (words(0) = "lineEnd") Or (words(0) = "   linksLinePoint") Or (words(0) = "   linksArcEnd") Or (words(0) = "   linksArcRestart") Or (words(0) = "   linksArcStart") Or (words(0) = "arcStart") Or (words(0) = "arcEnd") Then
                        Tilt_ON
                    Else
                        Tilt_Off
                    End If
                End If
                
                Call Tilt_Rotate(Rotation_Angle)
                'Save_Angle = ""
                'Save_Index = RotationAngle.ListIndex
                'Save_Angle = RotationAngle.List(RotationAngle.ListIndex)
            Else
                Tilt_Off
                'Should not to move the origin position
                Call Tilt_Rotate(0)
                'Save_Angle = "None"
            End If
            
           'Third, move to destination point.
            'If (words(0) = "fudicial") Or (VisionTeach.value = 1) Then
            'Without camera
            If (words(0) = "fudicial") Then
                Call New_PTPToXYZ(pointX, pointY, 0)
            Else
                If (RightNeedle.value = True) Then
                    If (VisionTeach.value = 0) Then
                        Call New_PTPToXYZ(pointX + Offset_DistanceX_Camera_R_Needle, pointY - Offset_DistanceY_Camera_R_Needle, pointZ + needleOffsetZ_R)
                    Else
                        Call New_PTPToXYZ(pointX, pointY, 0)
                    End If
                Else
                    'If (Rotation_Angle = "None") Then
                    '    Call New_PTPToXYZ(pointX + Offset_DistanceX_Camera_L_Needle, pointY - Offset_DistanceY_Camera_L_Needle, pointZ)
                    'ElseIf (Rotation_Angle = "") Then
                    '    'If (old_angle = "") And (old_angle = "No") Then
                    '    If (old_angle = "") Or (old_angle = "No") Or (old_angle = "None") Then
                    '        'There is angle in previous note.
                    '        Call New_PTPToXYZ(pointX + Offset_DistanceX_Camera_L_Needle, pointY - Offset_DistanceY_Camera_L_Needle, pointZ)
                    '    Else
                    '        Call New_PTPToXYZ(pointX, pointY, pointZ)
                    '    End If
                    'Else
                    '    Call New_PTPToXYZ(pointX, pointY, pointZ)
                    'End If
                    
                    If (VisionTeach.value = 0) Then
                        Call New_PTPToXYZ(pointX + Offset_DistanceX_Camera_L_Needle, pointY - Offset_DistanceY_Camera_L_Needle, pointZ + needleOffsetZ_L)
                    Else
                        Call New_PTPToXYZ(pointX, pointY, 0)
                    End If
                End If
            End If
            
            setSpeed (jogSpeedSlider.value - 1)
    
            fileNew.Enabled = True
            fileLoad.Enabled = True
            fileSave.Enabled = True
            Home.Enabled = True
            translateButton.Enabled = True
            deletePt.Enabled = True
            'If (LeftNeedle.value = True) And (VisionTeach.value = 0) Then
            '    modifyPt.Enabled = False
            'Else
                modifyPt.Enabled = True
            'End If
            addPt.Enabled = True
            trackButton.Enabled = True
            yPlus.Enabled = True
            yMinus.Enabled = True
            xPlus.Enabled = True
            xMinus.Enabled = True
            zPlus.Enabled = True
            zMinus.Enabled = True
        End If
    End If
End Sub

Private Sub New_PTPToXYZ(ByVal X As Long, ByVal y As Long, ByVal Z As Long)
    readyStatus = False
    busyStatus = True
    
    'To get the actual position
    y = y * (-1)
    Z = Z * (-1)
    
    'Exit the task that hasn't finished
    If Emergency_Stop = True Then
        readyStatus = True
        busyStatus = False
        'Emergency_Stop = False
        Exit Sub
    Else
        checkSuccess (P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, X, y, 0, 0))
    End If
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success) 'Loop while XY motor is still spinning
        DoEvents
    Loop
    
    'Exit the task that hasn't finished
    If Emergency_Stop = True Then
        readyStatus = True
        busyStatus = False
        'Emergency_Stop = False
        Exit Sub
    Else
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, Z, 0))
    End If
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
        DoEvents
    Loop
    
    Emergency_Stop = False
    readyStatus = True
    busyStatus = False
End Sub

Private Sub modifyPt_Click()
    SetFocusTimer.Enabled = False
    If (lstPattern.SelCount = 0) Then
        NodeError.Show (vbModal)
    Else
        clearLockedNode
        doingModify = True
        ErrorPartArray = False
        
        Dim i, currentIndex As Integer      'For list index
        Dim StartCount As Integer           'Define start of the main looping
        Dim ToStoreSelItems()               'To store the index of the selected items
        Dim StoreNum As Integer             'Create the size of index array
        Dim Count As Integer                'Number of sub array
        Dim ListCount As Integer            'Total list count
        Dim CountList As Integer            'Count the total number of selected items
        Dim no As Integer                   'Just for highlight looping
        Dim Name As String                  'List of List2
        Dim MessageForRegeneration As String 'Message String
        Dim Respond As Integer              'Return value of vbOK and vbCancel
        Dim words() As String               'To split the string
        Dim DoArray() As String             'Split the row and colum
        Dim totalArray As Integer           'Total sub array
        Dim NewprocessAddNode As String     'String text line
        
        i = 0
        no = 0
        StartCount = 0
        CountList = 0
        ListCount = 0
        ExpandWithDrawSpeed = 0
        NewprocessAddNode = ""
        ReDim ToStoreSelItems(0)
        StoreNum = 0
        
        'Do check the expanded array's header and sub array are selected instantenously
        If CheckingArrayIndividual = True Then
            Exit Sub
        End If
        
        i = lstPattern.ListIndex
        ListCount = lstPattern.ListCount
        CountList = SelectedItems(lstPattern.ListCount, currentIndex)
        
        If CountList = 1 Then
            i = currentIndex
            
            'un-modify the following string (XW)
            If (lstPattern.List(i) = "EndArray") Or (lstPattern.List(i) = "*** Left-Needle ***") _
                Or (lstPattern.List(i) = "*** Right-Needle ***") Then
                selectNodeIndex = lstPattern.ListIndex
                Exit Sub
            Else
                'Extract "Z value"
                SaveX = 0
                SaveY = 0
                SaveZ = 0
                
                Call TakeXYZValue(i, SaveX, SaveY, SaveZ)
                
                If (LeftNeedle.value = True) Then
                    Save_Old_zHigh = Z_High
                    Z_High = SaveZ
                Else
                    Save_Old_zHigh = R_Z_High
                    R_Z_High = SaveZ
                End If
            End If
            
            totalArray = 1
            Rows = CInt(yRepeatNum.Text)
            Columns = CInt(xRepeatNum.Text)
                
            totalArray = Columns * Rows
            
            words() = Split(lstPattern.List(i), "(")
            If (words(0) = "StartDotArray") Or (words(0) = "StartDotPottingArray") Or (words(0) = "StartLinePottingArray") Then
                
                SingleLineSelected = True
                Name = words(0)
            
                Count = CheckSubArray(i)
                        
                If (Count <> 0) Then
                    MessageForRegeneration = "Press 'OK' to regenerate Array or Press 'Cancel' not to do Modify!"
                    Respond = MsgBox(MessageForRegeneration, vbOKCancel, "Doing Modification")
                                
                    If Respond = vbOK Then
                        If StrComp(words(0), "StartDotArray", vbTextCompare) = 0 Then
                            'Old
                            'Call lstPattern.AddItem("StartDotArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", i)
                            'Call lstPattern.AddItem("StartDotArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                            
                            If (VisionTeach.value = 1) Then
                                If (LeftNeedle.value = True) Then
                                    Call lstPattern.AddItem("StartDotArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                                Else
                                    Call lstPattern.AddItem("StartDotArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & R_Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                                End If
                            Else
                                Call lstPattern.AddItem("StartDotArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                            End If
                            Call lstPattern.RemoveItem(i + 1)
                            
                        'No (pot dot and pot line)
                        'ElseIf StrComp(words(0), "StartDotPottingArray", vbTextCompare) = 0 Then
                        '    'Old
                        '    'Call lstPattern.AddItem("StartDotPottingArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", i)
                        '    Call lstPattern.AddItem("StartDotPottingArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                        '    Call lstPattern.RemoveItem(i + 1)
                        'ElseIf StrComp(words(0), "StartLinePottingArray", vbTextCompare) = 0 Then
                        '    'Old
                        '    'Call lstPattern.AddItem("StartLinePottingArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", i)
                        '    Call lstPattern.AddItem("StartLinePottingArray(x=" & convertToPulses(dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", i)
                        '    Call lstPattern.RemoveItem(i + 1)
                        End If
                                
                        For no = i To (i + Count - 1)
                            If lstPattern.List(i + 1) <> "EndArray" Then
                                Call lstPattern.RemoveItem(i + 1)
                            End If
                        Next no
                        WriteArrayTextLine (i)
                    Else
                        editorForm.Caption = "Profile Editor"
                        fileDirty = True
                        Click = 0
                        Modify = True
                        SingleLineSelected = False
                        lstPattern.Selected(i) = True
                        
                        If (LeftNeedle.value = True) Then
                            Z_High = Save_Old_zHigh
                        Else
                            R_Z_High = Save_Old_zHigh
                        End If
                        
                        Exit Sub
                    End If
                Else
                    MsgBox ("There is no array to Modify!")
                End If
                
                SingleLineSelected = False
            Else
                If editorForm.NodeType.ListIndex > 1 Then
                    'Modify "reference" by other elements
                    'or Modify other elements by other elements
                    Call lstPattern.AddItem(processAddNode, i)
                    If StrComp(words(0), "reference", vbTextCompare) = 0 Then
                        referenceX = 0
                        referenceY = 0
                        referenceZ = 0
                        referenceSet = False
                    End If
                ElseIf editorForm.NodeType.ListIndex = 0 And StrComp(words(0), "reference", vbTextCompare) = 0 Then
                    'Modify "reference" by itself, "reference".
                    Call lstPattern.AddItem(processAddNode, i)
                    referenceX = referenceX + convertToPulses(editorForm.dispensePtX.Text, X_axis)
                    referenceY = referenceY + convertToPulses(editorForm.dispensePtY.Text, Y_axis)
                    referenceZ = referenceZ + convertToPulses(editorForm.dispensePtZ.Text, Z_axis)
                    
                    lstPattern.RemoveItem (i + 1)
                    OffsetModify i + 1
                    
                    Modify = True
                    'To show the selected item after doing modify
                    lstPattern.Selected(i) = True
                    
                    If (LeftNeedle.value = True) Then
                        Z_High = Save_Old_zHigh
                    Else
                        R_Z_High = Save_Old_zHigh
                    End If
                    
                    Exit Sub
                Else
                    'Modify other elements by "reference".
                    Call lstPattern.AddItem(processAddNode, i)
                End If
        
                'Added to prevent empty lines on modify partarray
                If lstPattern.List(i) = "" Then
                    lstPattern.RemoveItem (i)
                End If
                
                If (ErrorPartArray = False) Then
                    lstPattern.RemoveItem (i + 1)
                Else
                    ErrorPartArray = False
                End If
        
            End If
            
            If totalArray > 1 Then
                words() = Split(lstPattern.List(i), "(")
                If (words(0) = "dotArray") Or (words(0) = "dotPottingArray") Or (words(0) = "linePottingArray") Then
                    Expand.Enabled = True
                End If
            Else
                Expand.Enabled = False
            End If
            
            Modify = True
            'To show the selected item after doing modify
            lstPattern.Selected(i) = True
            
            If (LeftNeedle.value = True) Then
                Z_High = Save_Old_zHigh
            Else
                R_Z_High = Save_Old_zHigh
            End If
            
        Else
            Do While (StartCount <> ListCount)
                If NodeType.ListIndex = 0 Then
                    MsgBox ("Reference Node Type only will be allowed to modify one line!")
                    Exit Do
                End If
                
                ''un-modify the following string (XW)
                'If (lstPattern.List(StartCount) = "EndArray") Or (lstPattern.List(StartCount) = "*** Left-Needle ***") _
                '    Or (lstPattern.List(StartCount) = "*** Right-Needle ***") Then
                '    selectNodeIndex = lstPattern.ListIndex
                '    Exit Sub
                'End If
            
                SaveX = 0
                SaveY = 0
                SaveZ = 0
                
                words() = Split(lstPattern.List(StartCount), "(")
                Name = words(0)
                    
                If lstPattern.Selected(StartCount) = True Then
                    If (lstPattern.List(i) = "EndArray") Or (lstPattern.List(i) = "*** Left-Needle ***") _
                        Or (lstPattern.List(i) = "*** Right-Needle ***") Then
                    Else
                        ReDim Preserve ToStoreSelItems(StoreNum)
                        ToStoreSelItems(StoreNum) = StartCount
                        StoreNum = StoreNum + 1
                
                    If editorForm.NodeType.ListIndex > 1 Then
                        'Modify "reference" by other elements.
                        'or Modify other elements by other elements
                    
                        If (Name = "StartDotArray") Or (Name = "StartDotPottingArray") Or (Name = "StartLinePottingArray") Then
                            
                            Call TakeXYZValue(StartCount, SaveX, SaveY, SaveZ)
                    
                            If StrComp(words(0), "StartDotArray", vbTextCompare) = 0 Then
                                'Old
                                'Call lstPattern.AddItem("StartDotArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", StartCount)
                                Call lstPattern.AddItem("StartDotArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", StartCount)
                                Call lstPattern.RemoveItem(StartCount + 1)
                            'No (pot dot and pot line)
                            'ElseIf StrComp(words(0), "StartPotPottingArray", vbTextCompare) = 0 Then
                            '    'Old
                            '    'Call lstPattern.AddItem("StartDotPottingArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", StartCount)
                            '    Call lstPattern.AddItem("StartDotPottingArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", StartCount)
                            '    Call lstPattern.RemoveItem(StartCount + 1)
                            'ElseIf StrComp(words(0), "StartLinePottingArray", vbTextCompare) = 0 Then
                            '    'Old
                            '    'Call lstPattern.AddItem("StartLinePottingArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")", StartCount)
                            '    Call lstPattern.AddItem("StartLinePottingArray(x=" & SaveX & ", y=" & SaveY & ", z=" & SaveZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")", StartCount)
                            '    Call lstPattern.RemoveItem(StartCount + 1)
                            End If
                            
                            totalArray = 1
                            Rows = CInt(yRepeatNum.Text)
                            Columns = CInt(xRepeatNum.Text)
                           
                            totalArray = Columns * Rows
                            
                            Count = CheckSubArray(StartCount)
                        
                            If (Count <> 0) Then
                                If (Count <> totalArray) Then
                                    If (Count < totalArray) Then
                                        MessageForRegeneration = "Press 'OK' to regenerate Array or Press 'Cancel' not to do Modify!"
                                        Respond = MsgBox(MessageForRegeneration, vbOKCancel, "Doing Modification")
                                
                                        If Respond = vbOK Then
                                            For no = StartCount To (StartCount + Count - 1)
                                                If lstPattern.List(StartCount + 1) <> "EndArray" Then
                                                    Call lstPattern.RemoveItem(StartCount + 1)
                                                End If
                                            Next no
                                            WriteArrayTextLine (StartCount)
                                            'Adding "+ 1" is not to modify "EndArray"
                                            StartCount = StartCount + totalArray + 1
                                            ListCount = ListCount + (totalArray - Count)
                                        Else
                                            Exit Do
                                        End If
                                    ElseIf (Count > totalArray) Then
                                        If NoEndArray = True Then
                                            
                                        Else
                                            MessageForRegeneration = "Press 'OK' to regenerate Array or Press 'Cancel' not to do Modify!"
                                            Respond = MsgBox(MessageForRegeneration, vbOKCancel, "Doing Modification")
                                            If Respond = vbOK Then
                                                For no = StartCount To (StartCount + Count - 1)
                                                    If lstPattern.List(StartCount + 1) <> "EndArray" Then
                                                        Call lstPattern.RemoveItem(StartCount + 1)
                                                    End If
                                                Next no
                                                WriteArrayTextLine (StartCount)
                                                StartCount = StartCount + totalArray + 1
                                                ListCount = ListCount + (totalArray - Count)
                                            Else
                                                Exit Do
                                            End If
                                        End If
                                    End If
                                Else
                                    For no = StartCount To (StartCount + Count - 1)
                                        If lstPattern.List(StartCount + 1) <> "EndArray" Then
                                            Call lstPattern.RemoveItem(StartCount + 1)
                                        End If
                                    Next no
                                    WriteArrayTextLine (StartCount)
                                    StartCount = StartCount + totalArray + 1
                                End If
                            Else
                                MsgBox ("There is no array to Modify!")
                            End If
                        Else
                            NewprocessAddNode = ModifyElements(StartCount)
                            Call lstPattern.AddItem(NewprocessAddNode, StartCount)
                            
                            lstPattern.RemoveItem (StartCount + 1)
                                        
                            'Added to prevent empty lines on modify partarray
                            If lstPattern.List(StartCount) = "" Then
                                lstPattern.RemoveItem (StartCount)
                            End If
                        
                        End If
                    End If
                    End If
                End If
                StartCount = StartCount + 1
            Loop
        
            For no = 0 To StoreNum - 1
                lstPattern.Selected(ToStoreSelItems(no)) = True
                Click = 0
            Next no
            
        End If
        
        editorForm.Caption = "Profile Editor"     'XW
    End If

    fileDirty = True
    'To prevent not to de-highlight the selected line when we click the "Modify Button"
    Click = 0
    SetFocusTimer.Enabled = True
End Sub

Private Sub cmdOffset_Click()
    Dim reference() As String
    Dim TotalList As Integer
    Dim ToStoreSelItems()               'To store the index of the selected items
    Dim StoreNum As Integer             'Create the size of index array
    
    clearLockedNode
    doingModify = True
        
    StoreNum = 0
    ReDim ToStoreSelItems(0)
        
    Do While (TotalList < (lstPattern.ListCount))
        If (lstPattern.Selected(TotalList) = True) Then
                    
            reference() = Split(lstPattern.List(TotalList), "(")
                    
            ReDim Preserve ToStoreSelItems(StoreNum)
            ToStoreSelItems(StoreNum) = TotalList
            StoreNum = StoreNum + 1
                    
            If (reference(0) <> "EndArray") Then
                If (reference(0) = "reference") Then
                    referenceX = referenceX + convertToPulses(GlobalOffsetX, X_axis)
                    referenceY = referenceY + convertToPulses(GlobalOffsetY, Y_axis)
                    referenceZ = referenceZ + convertToPulses(GlobalOffsetZ, Z_axis)
                    OffsetModify TotalList
                    Exit Do
                End If
            
                Call GettingPosition(lstPattern.List(TotalList))
                    
                If (reference(0) = "StartDotArray") Or (reference(0) = "StartDotPottingArray") Or (reference(0) = "StartLinePottingArray") Then
                    Call OffsetString(ModifyOffsetX, ModifyOffsetY, ModifyOffsetZ, TotalList)
                    lstPattern.RemoveItem (TotalList + 1)
                    TotalList = TotalList + 1
                            
                    Do While (lstPattern.List(TotalList) <> "EndArray")
                        Call GettingPosition(lstPattern.List(TotalList))
                        Call OffsetString(ModifyOffsetX, ModifyOffsetY, ModifyOffsetZ, TotalList)
                        lstPattern.RemoveItem (TotalList + 1)
                        TotalList = TotalList + 1
                    Loop
                    
                    TotalList = TotalList - 1
                Else
                    Call OffsetString(ModifyOffsetX, ModifyOffsetY, ModifyOffsetZ, TotalList)
                    If (reference(0) = "fudicial") Then
                        lstPattern.RemoveItem (TotalList)
                    Else
                        lstPattern.RemoveItem (TotalList + 1)
                    End If
                End If
            End If
        End If
            
        TotalList = TotalList + 1
    Loop
            
    For TotalList = 0 To StoreNum - 1
        lstPattern.Selected(ToStoreSelItems(TotalList)) = True
    Next TotalList
                
    'Reinitialize the offset value.
    If (GroupOffset = True) Then
        GroupOffset = False
    End If
    
    FlagGlobalOffset = False
    fileDirty = True
    'To prevent not to de-highlight the selected line when we click the "Offset Button"
    Click = 0
    
End Sub

Private Sub OffsetModify(ByVal Start As Integer)
    Dim i As Integer
    Dim words() As String
    
    For i = Start To (lstPattern.ListCount - 1)
        If (i = lstPattern.ListCount) Then
            Exit Sub
        End If
    
        words() = Split(lstPattern.List(i), "=")
                    
        Call GettingPosition(lstPattern.List(i))
        Call OffsetString(ModifyOffsetX, ModifyOffsetY, ModifyOffsetZ, i)
        lstPattern.RemoveItem (i + 1)
    Next i
    
End Sub

Private Function TakeXYZValue(ByVal Index As Integer, SaveX As Variant, SaveY As Variant, SaveZ As Variant)
    Dim StringLine As String
    Dim WriteString As String
    Dim OneChar As String
    Dim Start As Integer
    Dim CountEqualSign As Integer
    Dim flag As Boolean
    
    Start = 1
    CountEqualSign = 0
    TotalNumberCount = 0
    OneChar = ""
    WriteString = ""
    flag = False
    StringLine = lstPattern.List(Index)
    
    Do
        OneChar = Mid(StringLine, Start, 1)
        If (flag = True) And (OneChar <> ",") And (OneChar <> ";") Then
            WriteString = WriteString & OneChar
        End If
        If OneChar = "=" Then
            flag = True
            CountEqualSign = CountEqualSign + 1
        ElseIf OneChar = "," Then
            If CountEqualSign = 1 Then
                SaveX = Val(WriteString)
            ElseIf CountEqualSign = 2 Then
                SaveY = Val(WriteString)
            End If
            flag = False
            WriteString = ""
            TotalNumberCount = 0
        ElseIf OneChar = ";" And CountEqualSign = 3 Then
            SaveZ = Val(WriteString)
        End If
        Start = Start + 1
    Loop While (OneChar <> ";")
    
End Function

Private Function ModifyArrayElements() As String
    If SingleLineSelected = True Then
        Select Case NodeType.ListIndex
            Case 2
                'Old
                'ModifyArrayElements = "   dot(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                'ModifyArrayElements = "   dot(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                
                If (VisionTeach.value = 1) Then
                    If (LeftNeedle.value = True) Then
                        ModifyArrayElements = "   dot(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                    Else
                        ModifyArrayElements = "   dot(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & R_Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                    End If
                 Else
                    ModifyArrayElements = "   dot(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                End If
            Case 3
                'Old
                'ModifyArrayElements = "   dotPotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                'ModifyArrayElements = "   dotPotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                
                If (VisionTeach.value = 1) Then
                    If (LeftNeedle.value = True) Then
                        ModifyArrayElements = "   dotPotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                    Else
                        ModifyArrayElements = "   dotPotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & R_Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                    End If
                Else
                    ModifyArrayElements = "   dotPotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                End If
            Case 5
                'Old
                'ModifyArrayElements = "   linePotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                'ModifyArrayElements = "   linePotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
                
                If (VisionTeach.value = 1) Then
                    If (LeftNeedle.value = True) Then
                        ModifyArrayElements = "   linePotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                    Else
                        ModifyArrayElements = "   linePotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & R_Z_High + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                    End If
                Else
                    ModifyArrayElements = "   linePotting(x=" & convertToPulses(dispensePtX.Text, X_axis) + add_column_pitch + referenceX & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + add_row_pitch + referenceY & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                End If
        End Select
    Else
        Select Case NodeType.ListIndex
            Case 2
                'Old
                'ModifyArrayElements = "   dot(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                ModifyArrayElements = "   dot(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
            Case 3
                'Old
                'ModifyArrayElements = "   dotPotting(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                ModifyArrayElements = "   dotPotting(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
            Case 5
                'Old
                'ModifyArrayElements = "   linePotting(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
                ModifyArrayElements = "   linePotting(x=" & SaveX + add_column_pitch + referenceX & ", y=" & SaveY + add_row_pitch + referenceY & ", z=" & SaveZ + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        End Select
    End If
End Function

''''''''''''''''''
'Modiy + Offset  '
''''''''''''''''''
'Private Function ModifyElements(ByVal StartCount As Integer) As String
'    Dim words() As String
'    Dim SplitName() As String
'    Dim filePart() As String
'    Dim WriteString As String
'    Dim Name As String
    
'    WriteString = ""
    
'    SplitName() = Split(lstPattern.List(StartCount), "(")
'    words() = Split(lstPattern.List(StartCount), "=")
'    filePart() = Split(lstPattern.List(StartCount), ";")
    
'    Name = SplitName(0)
    
'    Call GettingPosition(lstPattern.List(StartCount))
    
'    Select Case Name
'        Case "dot"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "   dot"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "dotPotting"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "   dotPotting"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "linePotting"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "   linePotting"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "lineStart"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & Format(delay.Text, "####0.000") & ")"
'        Case "lineEnd"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "arcStart"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & Format(delay.Text, "####0.000") & ")"
'        Case "       arcPoint"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & ")"
'        Case "arcEnd"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "   linksLinePoint"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
'        Case "   linksArcStart"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
'        Case "   linksArcRestart"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
'        Case "   linksArcEnd"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
'        Case "rectC1"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; " & Format(delay.Text, "####0.000") & ")"
'        Case "   rectC2"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & ")"
'        Case "rectC3"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
'        Case "repeat"
'            WriteString = words(0) & "=" & ModifyOffsetX + convertToPulses(txtOffsetX.Text, X_axis) + referenceX & ", y=" & ModifyOffsetY + convertToPulses(txtOffsetY.Text, Y_axis) + referenceY & ", z=" & ModifyOffsetZ + convertToPulses(txtOffsetZ.Text, Z_axis) + referenceZ & ";" & filePart(1)
'    End Select
    
'    ModifyElements = WriteString
    
'    GroupOffset = True
'End Function
Private Function ModifyElements(ByVal StartCount As Integer) As String
    Dim words() As String
    Dim SplitName() As String
    Dim WriteString As String
    Dim Name As String
    
    WriteString = ""
    
    SplitName() = Split(lstPattern.List(StartCount), "(")
    words() = Split(lstPattern.List(StartCount), ";")
    
    Name = SplitName(0)
    
    Select Case Name
        Case "dot"
            'Old
            'We don't need to chang "withDrawalSpeed" to "ExpandWithDrawSpeed" because this parameter is created by the user.
            'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        Case "   dot"
            'Old
            'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        Case "dotArray"
            'Old
            'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(CLng(WithDrawalZ.Text), Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        'No pot dot and pot line
        'Case "dotPotting"
        '    'Old
        '    'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
        '    WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        'Case "   dotPotting"
        '    'Old
        '    'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
        '    WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        'Case "linePotting"
        '    'Old
        '    'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
        '    WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        'Case "   linePotting"
        '    'Old
        '    'WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
        '    WriteString = words(0) & "; " & convertToPulses(xDev.Text, X_axis) & ", " & xRepeatNum.Text & "; " & convertToPulses(yDev.Text, Y_axis) & ", " & yRepeatNum.Text & "; z=" & convertToPulses(potDepth.Text, Z_axis) & "; sp=" & depthSpeed.Text & "; " & Format(delay.Text, "####0.000") & "; z=" & convertToPulses(endDispenseHeight.Text, Z_axis) & "; sp=" & DispenseSpeed.Text & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        Case "lineStart"
            'Old
            'WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & ")"
            WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & "; Angle=" & RotationAngle.Text & ")"
        Case "lineEnd"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        Case "arcStart"
            'Old
            'WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & ")"
            WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & "; Angle=" & RotationAngle.Text & ")"
        Case "       arcPoint"
            'Old
            'WriteString = words(0)
            WriteString = words(0) & "; Angle=" & RotationAngle.Text & ")"
        Case "arcEnd"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
        Case "   linksLinePoint"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; Angle=" & RotationAngle.Text & ")"
        Case "   linksArcStart"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; Angle=" & RotationAngle.Text & ")"
        Case "   linksArcRestart"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; Angle=" & RotationAngle.Text & ")"
        Case "   linksArcEnd"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.Value & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; Angle=" & RotationAngle.Text & ")"
        Case "rectC1"
            'Old
            'WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & ")"
            'WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & "; pitch=" & convertToPulses(yDev.Text, Y_axis) & ")"
            WriteString = words(0) & "; " & Format(delay.Text, "####0.000") & "; pitch=" & convertToPulses(txtPitch.Text, Y_axis) & ";" & NoSprayArea.value & "; Angle=" & RotationAngle.Text & ")"
        Case "   rectC2"
            'Old
            'WriteString = words(0)
            WriteString = words(0) & "; Angle=" & RotationAngle.Text & ")"
        Case "rectC3"
            'Old
            'WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & ")"
            WriteString = words(0) & "; sp=" & DispenseSpeed.Text & "; " & dispenseOnOff.value & "; " & Format(retractDelay.Text, "####0.000") & "; z=" & convertToPulses(WithDrawalZ.Text, Z_axis) & "; sp=" & withdrawalSpeed.Text & "; z=" & convertToPulses(moveHeight.Text, Z_axis) & "; Angle=" & RotationAngle.Text & ")"
            
            'repeat (no group modify for "part array")
    End Select
    
    ModifyElements = WriteString
End Function

Private Sub OffsetString(ByVal xValue As Long, ByVal yValue As Long, ByVal zValue As Long, ByVal LineIndex As Integer)
    
    Dim words() As String, elements() As String, NodeType() As String
    
    elements() = Split(lstPattern.List(LineIndex), "=")
    words() = Split(lstPattern.List(LineIndex), ";")
    NodeType() = Split(lstPattern.List(LineIndex), "(")
    
    Select Case NodeType(0)
        Case "reference"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
            Else
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1), LineIndex)
            End If
        Case "       arcPoint"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ")", LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & "; " & words(1), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ")", LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & "; " & words(1), LineIndex)
            End If
        Case "dot", "   dot", "dotArray"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8), LineIndex)
            End If
        'No pot dot and pot line
        'Case "dotPotting", "   dotPotting", "dotPottingArray"
        '    If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
        '        'Old (No angle)
        '        'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9), LineIndex)
        '        Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10), LineIndex)
        '    Else
        '        'Old (No angle)
        '        'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9), LineIndex)
        '        Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10), LineIndex)
        '    End If
        'Case "linePotting", "   linePotting", "linePottingArray"
        '    If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
        '        'Old (No angle)
        '        'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11), LineIndex)
        '        Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11) & ";" & words(12), LineIndex)
        '    Else
        '        'Old (No angle)
        '        'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11), LineIndex)
        '        Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7) & ";" & words(8) & ";" & words(9) & ";" & words(10) & ";" & words(11) & ";" & words(12), LineIndex)
        '    End If
        Case "lineStart"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
            End If
        Case "arcStart"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
            End If
        Case "lineEnd", "arcEnd"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
            End If
        Case "   linksLinePoint", "   linksArcEnd"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
            End If
        Case "   linksArcRestart", "   linksArcStart"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
            Else
                'Old (No angle)
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3), LineIndex)
            End If
        Case "rectC1"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4), LineIndex)
            Else
                'Old
                'Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(txtOffsetX.Text, X_axis) + xValue & ", y=" & convertToPulses(txtOffsetY.Text, Y_axis) + yValue & ", z=" & convertToPulses(txtOffsetZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4), LineIndex)
            End If
        Case "   rectC2"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
            Else
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1), LineIndex)
            End If
        Case "rectC3"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
            Else
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1) & ";" & words(2) & ";" & words(3) & ";" & words(4) & ";" & words(5) & ";" & words(6) & ";" & words(7), LineIndex)
            End If
        Case "repeat"
            'If (txtOffsetX.Text = "0") And (txtOffsetY.Text = "0") And (txtOffsetZ.Text = "0") Then
            If (GlobalOffsetX = "0") And (GlobalOffsetY = "0") And (GlobalOffsetZ = "0") Then
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(dispensePtX.Text, X_axis) + xValue & ", y=" & convertToPulses(dispensePtY.Text, Y_axis) + yValue & ", z=" & convertToPulses(dispensePtZ.Text, Z_axis) + zValue & ";" & words(1), LineIndex)
            Else
                Call lstPattern.AddItem(elements(0) & "=" & convertToPulses(GlobalOffsetX, X_axis) + xValue & ", y=" & convertToPulses(GlobalOffsetY, Y_axis) + yValue & ", z=" & convertToPulses(GlobalOffsetZ, Z_axis) + zValue & ";" & words(1), LineIndex)
            End If
    End Select
    
    GroupOffset = True
End Sub

Private Function CheckingArrayIndividual() As Boolean
    Dim no As Integer                   'For looping
    Dim Name As String                  'For the List
    Dim CompareString As String         'To compare the two string
    Dim NotCheckFirstTime As Boolean    'Will not check in the first time
    Dim words() As String
    Dim Program_Part As String          'Save the part to compare the left/right option
    
    Name = ""
    CompareString = ""
    Program_Part = ""
    no = 0
    NotCheckFirstTime = False
    
    'This looping is checking the whole list of the list box whether the user select
    'the expanded array and sub array instantenously before doing the modification
    For no = 0 To lstPattern.ListCount - 1
        If (lstPattern.Selected(no) = True) Then
            'Set the high priority for checking the needle
            If (Program_Part = "*** Left-Needle ***") Then
                If (RightNeedle.value = True) Then
                    MsgBox ("Please choose the corrected neelde before doing the modification!")
                    CheckingArrayIndividual = True
                    Exit For
                End If
            ElseIf (Program_Part = "*** Right-Needle ***") Then
                If (LeftNeedle.value = True) Then
                    MsgBox ("Please choose the corrected neelde before doing the modification!")
                    CheckingArrayIndividual = True
                    Exit For
                End If
            End If
            
            Name = lstPattern.List(no)
            words() = Split(Name, "(")
            
            If (words(0) = "dotArray") Or (words(0) = "   dot") Or (words(0) = "StartDotArray") Then
                words(0) = "dot"
            'No pot dot and pot line
            'ElseIf (words(0) = "dotPottingArray") Or (words(0) = "   dotPotting") Or (words(0) = "StartDotPottingArray") Then
            '    words(0) = "dotPotting"
            'ElseIf (words(0) = "linePottingArray") Or (words(0) = "   linePotting") Or (words(0) = "StartLinePottingArray") Then
            '    words(0) = "linePotting"
            End If
            
            If (NotCheckFirstTime = True) Then
                'If we don't allow the user to do group modify for part array, we can use the following code.
                If (CompareString = words(0)) And (words(0) = "repeat") Then
                    MsgBox ("Part Array could not be allowed to do group modify!")
                    CheckingArrayIndividual = True
                    Exit For
                End If
                       
                'Not allowed to modify the different Names
                If (CompareString <> words(0)) Then
                    MsgBox ("Different elements will not be allowed to do modify!")
                    CheckingArrayIndividual = True
                    Exit For
                End If
            End If
            
            CompareString = words(0)
            NotCheckFirstTime = True
        End If
        
        'Save Part
        If (lstPattern.List(no) = "*** Left-Needle ***") Or (lstPattern.List(no) = "*** Right-Needle ***") Then
            Program_Part = lstPattern.List(no)
        End If
    Next no
End Function

Private Function CheckSubArray(ByVal StartCount As Integer) As Integer
    Dim Count As Integer
    Dim words() As String
    
    'This one is set "1" because we do start the looping from the next index
    Count = 1
    
    words() = Split(lstPattern.List(StartCount + Count), "(")
                        
    Do While (words(0) <> "EndArray")
        If (lstPattern.ListCount = StartCount + Count) Then
            CheckSubArray = 0
            NoEndArray = True
            Exit Function
        End If
        Count = Count + 1
        words() = Split(lstPattern.List(StartCount + Count), "(")
    Loop
    
    Count = Count - 1
    
    CheckSubArray = Count
End Function

Private Sub potDepth_Validate(cancel As Boolean)
    Call validateNumber(editorForm.potDepth.Text, editorForm.PotDepthLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub depthspeed_Validate(cancel As Boolean)
    Call validateNumber(editorForm.depthSpeed.Text, editorForm.depthSpeedLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(depthSpeed.Text) < 0) Then
            MsgBox "There is no negative value in depthSpeed."
            depthSpeed.Text = ""
            cancel = True
        Else
            depthSpeed.Text = CStr(CLng(depthSpeed.Text))
        End If
    End If
End Sub

Private Sub enddispenseheight_Validate(cancel As Boolean)
    Call validateNumber(editorForm.endDispenseHeight.Text, editorForm.endDispenseHeightLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub delay_Validate(cancel As Boolean)
    Call validateNumber(editorForm.delay.Text, editorForm.delayLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
    
    If (CDbl(delay.Text) < 0) Then
        MsgBox "There is no negative value in delay time."
        delay.Text = ""
        cancel = True
    End If
End Sub

Private Sub dispensespeed_Validate(cancel As Boolean)
    Call validateNumber(editorForm.DispenseSpeed.Text, editorForm.dispenseSpeedLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If CLng(editorForm.DispenseSpeed.Text) > 500 Then
            editorForm.DispenseSpeed.Text = "500"
        End If
        
        If (CDbl(DispenseSpeed.Text) < 0) Then
            MsgBox "There is no negative value in DispenseSpeed."
            DispenseSpeed.Text = ""
            cancel = True
        Else
            DispenseSpeed.Text = CStr(CLng(DispenseSpeed.Text))
        End If
    End If
End Sub

Private Sub retractdelay_Validate(cancel As Boolean)
    Call validateNumber(editorForm.retractDelay.Text, editorForm.retractDelayLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(retractDelay.Text) < 0) Then
            MsgBox "There is no negative value for time."
            retractDelay.Text = ""
            cancel = True
        End If
    End If
End Sub

Private Sub translateButton_Click()
    
    Dim xlimit, ylimit, zlimit, ulimit As Long
    Dim ReadValue As Long
    
    jogSpeedSlider.value = 28
    Call setSpeed(jogSpeedSlider.value)
    
    'Move Z_axis to '0' before doing Tilt Off    '@SK
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
    Loop
    
    '@SK
    Tilt_Off
    RotationAngle.Text = "None"
    Call Tilt_Rotate(0)
    
    leftside = False
    rightside = False
    
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, xlimit, ylimit, zlimit, ulimit))
    
    xlimit = xlimit And &HC
    ylimit = ylimit And &HC
    zlimit = zlimit And &HC
   
    'To avoid the limit sensor
    If ((xlimit <> 0) Or (ylimit <> 0) Or (zlimit <> 0)) And home_limit_flag = False Then    'xu long
        Call moveToHome
        homeRobotForm.Show (vbModal)
    Else
    
        clearLockedNode
        If editorForm.Caption = "Profile Editor" Then
            fileNotLoadedError.Show (vbModal)
        Else
            If fileDirty = True Then
                fileNotSavedForm.Show (vbModal)
                If proceedWithAction = True Then
                    Do_TranslateForm
                End If
            Else
                'check whether no program or not 'NNO
                File1size = FileLen(editorForm.Caption)
                If File1size <> 0 Then
                    Do_TranslateForm
                Else
                    MsgBox "There is no program to run the machine!"
                End If
            End If
            'Reinitialize because an error occur when translation is fail.
            SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
            SystemMoveHeight = SystemMoveHeight * (-1)      'XW
        End If
    End If
    jogSpeedSlider.value = 28
    Call setSpeed(jogSpeedSlider.value)
End Sub

Private Sub Do_TranslateForm()
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
        
    'Wait for a few (m sec)
    Sleep (0.3)
        
'    'Right-needle will be gone up.
'    checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
'    ReadValue = ReadValue And &HFEFF
'    checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
'
'    'Wait for a few (m sec)
'    Sleep (0.3)
        
    'Tilting OFF
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HFEFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Sleep (0.3)
                
    lockSaveMoveControls
    translateButton.Enabled = False
    SetFocusTimer.Enabled = False
    'These two parameters will not be changed after doing the translation
    'XW
    systemHomeY = systemHomeY * (-1)
    systemHomeZ = systemHomeZ * (-1)
    TotalLine = lstPattern.ListCount    'Total Line in listbox (XW)
    translateForm.startTranslate
    systemHomeY = systemHomeY * (-1)
    systemHomeZ = systemHomeZ * (-1)
                
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
                
    unLockSaveMoveControls
    NodeType_Click
    translateButton.Enabled = True
    SetFocusTimer.Enabled = True
End Sub

Private Sub UpDownDepthSpeed_upClick()
        depthSpeed.Text = CDbl(depthSpeed.Text) + 1
End Sub

Private Sub UpDownDepthSpeed_DownClick()
    If depthSpeed.Text <> 0 Then
        depthSpeed.Text = CDbl(depthSpeed.Text) - 1
    End If
End Sub

Private Sub UpDownEndDispenseHeight_upClick()
        endDispenseHeight.Text = CDbl(endDispenseHeight.Text) + 0.1
End Sub

Private Sub UpDownEndDispenseHeight_DownClick()
        endDispenseHeight.Text = CDbl(endDispenseHeight.Text) - 0.1
End Sub

Private Sub UpDownPotDepth_DownClick()
        potDepth.Text = CDbl(potDepth.Text) - 0.1
End Sub

Private Sub UpDownPotDepth_upClick()
        potDepth.Text = CDbl(potDepth.Text) + 0.1
End Sub

Private Sub withdrawalspeed_Validate(cancel As Boolean)
    Call validateNumber(editorForm.withdrawalSpeed.Text, editorForm.withDrawalSpeedLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(withdrawalSpeed.Text) < 0) Then
            MsgBox "There is no negative value in withdrawalSpeed."
            withdrawalSpeed.Text = ""
            cancel = True
        Else
            withdrawalSpeed.Text = CStr(CLng(withdrawalSpeed.Text))
        End If
    End If
End Sub

Private Sub withdrawalz_Validate(cancel As Boolean)
    Call validateNumber(editorForm.WithDrawalZ.Text, editorForm.withDrawalZLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub moveheight_Validate(cancel As Boolean)
    Call validateNumber(editorForm.moveHeight.Text, editorForm.moveHeightLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub xCoOrd_GotFocus()
    'XW
    editorForm.KeyPreview = False
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempX = xCoOrd.Text
        xCoOrd.Locked = False
        lockMoveControls
    End If
End Sub

Private Sub yCoOrd_GotFocus()
    'XW
    editorForm.KeyPreview = False
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempY = yCoOrd.Text
        yCoOrd.Locked = False
        lockMoveControls
    End If
End Sub

Private Sub zCoOrd_GotFocus()
    'XW
    editorForm.KeyPreview = False
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempZ = zCoOrd.Text
        zCoOrd.Locked = False
        lockMoveControls
    End If
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
    Call validateNumber(editorForm.StepDistance.Text, editorForm.LabelDistance.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If CDbl(editorForm.StepDistance.Text) <= 0 Then
            editorForm.StepDistance.Text = "0.001"
        ElseIf CDbl(editorForm.StepDistance.Text) >= 10 Then
            editorForm.StepDistance.Text = "10.000"
        Else
            editorForm.StepDistance.Text = Format(editorForm.StepDistance.Text, "#0.000")
        End If
    End If
End Sub

Private Sub UpDownStep_DownClick()
    If CDbl(StepDistance.Text <> 0) Then
        StepDistance.Text = CDbl(StepDistance.Text) - 0.001
    End If
End Sub

Private Sub UpDownStep_UpClick()
    StepDistance.Text = CDbl(StepDistance.Text) + 0.001
End Sub

Private Sub xMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    'Show the indication    'XW
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, X_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) + CDbl(StepDistance.Text), X_axis), 0, 0, 0))
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Private Sub xPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, X_axis, 1))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, X_axis, X_axis, convertToPulses(CDbl(xCoOrd.Text) - CDbl(StepDistance.Text), X_axis), 0, 0, 0))
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Private Sub xMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
End Sub

Private Sub xPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
End Sub

Private Sub yMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 2))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) + CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0))
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success)  'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Private Sub yPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Y_axis, Y_axis, 0, (convertToPulses(CDbl(yCoOrd.Text) - CDbl(StepDistance.Text), Y_axis)) * (-1), 0, 0))
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success) 'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Private Sub yMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
End Sub

Private Sub yPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
End Sub

Private Sub zMinus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 4))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) + CDbl(StepDistance.Text), Z_axis)) * (-1), 0))
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Private Sub zPlus_mouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    readyStatus = False
    busyStatus = True
    If Jogging.value = True Then
        setSpeed (jogSpeedSlider.value - 1)
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 0))
    ElseIf JoggingStep.value = True Then
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, (convertToPulses(CDbl(zCoOrd.Text) - CDbl(StepDistance.Text), Z_axis)) * (-1), 0))
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Private Sub zMinus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
End Sub

Private Sub zPlus_mouseup(Button As Integer, Shift As Integer, X As Single, y As Single)
    busyStatus = False
    readyStatus = True
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
End Sub

Private Sub xrepeatnum_Validate(cancel As Boolean)
    Call validateNumber(editorForm.xRepeatNum.Text, editorForm.xRepeatNumLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(xRepeatNum.Text) < 0) Then
            MsgBox "There is no negative value in xRepeatNum."
            xRepeatNum.Text = ""
            cancel = True
        Else
            xRepeatNum.Text = CStr(CLng(xRepeatNum.Text))
        End If
    End If
End Sub

Private Sub yrepeatnum_Validate(cancel As Boolean)
    Call validateNumber(editorForm.yRepeatNum.Text, editorForm.yRepeatNumLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    Else
        If (CDbl(yRepeatNum.Text) < 0) Then
            MsgBox "There is no negative value in yRepeatNum."
            yRepeatNum.Text = ""
            cancel = True
        Else
            yRepeatNum.Text = CStr(CLng(yRepeatNum.Text))
        End If
    End If
End Sub

Private Sub xdev_Validate(cancel As Boolean)
    Call validateNumber(editorForm.xDev.Text, editorForm.xDevLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub ydev_Validate(cancel As Boolean)
    Call validateNumber(editorForm.yDev.Text, editorForm.yDevLabel.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub fileLoad_Click()
    SetFocusTimer.Enabled = False
    proceedWithAction = True
    
    If fileDirty = True Then
        fileNotSavedForm.Show (vbModal)
    End If
    
    If proceedWithAction = True Then

        clearLockedNode
        initializeInputParams
        fileLoadForm.Show (vbModal)
        
        'Check whether the user loads a new file or not (XW)
        If (File_Load_Cancel = True) Then
            File_Load_Cancel = False
        Else
            fileDirty = False
            doingModify = False
        
            'XW
            If (Right_Needle_ON = True) Then
                RightNeedle_Click
                'RightNeedle_No = 1     'removed by NNO (after reinitialize, should not be 1)
                Right_Needle_ON = False
            Else
                LeftNeedle_Click        'Set the default (reinitialize)
                'LeftNeedle_No = 1      'removed by NNO
            End If
        End If
    End If
    SetFocusTimer.Enabled = True
End Sub

Private Sub fileSave_Click()
    SetFocusTimer.Enabled = False
    'If (lstPattern.ListCount = 0 And firstteaching = False) Then
    '    MsgBox "There is no program to save!"
    '    SetFocusTimer.Enabled = True
    '    Exit Sub
    'End If
    clearLockedNode
    fileSaveForm.Show (vbModal)
    SetFocusTimer.Enabled = True
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
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
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
            Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
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
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success) 'Loop while Y motor is still spinning
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
            Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success)  'Loop while Y motor is still spinning
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
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
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
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
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
    
    If (KeyCode = 17) Or (KeyCode = 97) Then
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
    ElseIf Shift = vbShiftMask + vbCtrlMask Then
        MsgBox ("Please don't press the two keys at the same time!")
        Exit Sub
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
                Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
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
    lngScreenY = Point.y
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
    Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)
    Loop
End Sub

Private Sub YaxisStop()
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success)
    Loop
End Sub

Private Sub XYaxisStop()
    checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis, 1 Or 2))
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
    Loop
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
            Exit Sub
        End If
    End If
    
    If (KeyCode = 97) Or (KeyCode = 103) Then
        'If we don't do this checking, the processing will be fast.
        'If (OneSevenKey > 1) Then
            addPt.SetFocus
        'End If
        Disable
        OneSevenKey = 0
    End If
    
    If (Indicator = True) Then
        If (KeyCode = vbKeyRight) Or (KeyCode = 102) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = 100) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
            Indicator = True
        ElseIf (KeyCode = vbKeyUp) Or (KeyCode = 104) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
            Indicator = True
        ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
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
        'addPt.SetFocu
    ElseIf (Reflector = True) Then
        If (KeyCode = vbKeyUp) Or (KeyCode = 104) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf (KeyCode = vbKeyDown) Or (KeyCode = 98) Then
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = True
        ElseIf (KeyCode = vbKeyLeft) Or (KeyCode = vbKeyRight) Or (KeyCode = 100) Or (KeyCode = 102) Then
            Indicator = False
            Reflector = True
        Else
            busyStatus = False
            readyStatus = True
            checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
            Reflector = False
        End If
        'addPt.SetFocus
    ElseIf (MovingMouse = True) Then
        busyStatus = False
        readyStatus = True
        checkSuccess (P1240MotStop(boardNum, X_axis Or Y_axis, 1 Or 2))
        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
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

    hWnd = FindWindow(vbNullString, "Desktop Setup Panel")
    hWnd2 = FindWindow(vbNullString, "File Load")

    SetFocusTimer.Enabled = False
    resetTimer.Enabled = False
    ClickTimer.Enabled = False
    CheckEmergencyStop.Enabled = False
    MouseMovement.Enabled = False

    If App.PrevInstance Or hWnd <> 0 Or hWnd2 <> 0 Then
    
        MsgBox ("Another conflicting process has been detected! This process will abort")
        Unload Me
        
    Else
        'Skin1.LoadSkin (".\skin\epoxySkin.skn")
        Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
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
        
            WindowHook Me.lstPattern.hWnd
    
            readyStatus = False
            busyStatus = False
            errorStatus = False
            ClickTimer.Enabled = False      'XW
            
            'Tower Light
            Red_Light = False
            Yellow_Light = False
            Green_Light = True
            
            referenceX = 0
            referenceY = 0
            referenceZ = 0
    
            fileDirty = False
    
            readRegistryOptions
    
            determineProfile
 
            SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
            SystemMoveHeight = SystemMoveHeight * (-1)  'XW
            systemTrackMoveHeight = 0
            ActualSystemMoveHeight = SystemMoveHeight   'XW
        
            systemHomeX = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "xSystemHome", "0")), X_axis)
            systemHomeY = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "ySystemHome", "0")), Y_axis)
            systemHomeZ = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "zSystemHome", "0")), Z_axis)
            
            '@$K
            SolventPosX = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "xSolventPos", "0")), X_axis)
            SolventPosY = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "ySolventPos", "0")), Y_axis)
            SolventPosZ = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "zSolventPos", "0")), Z_axis)
            
            'The offset value between upper camera and left_needle
            Offset_DistanceX_Camera_L_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_L_Needle", "0")
            Offset_DistanceY_Camera_L_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_L_Needle", "0")
            
            'The offset value between upper camera and right_needle
            Offset_DistanceX_Camera_R_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_R_Needle", "0")
            Offset_DistanceY_Camera_R_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_R_Needle", "0")
            
            'Z offset valve for both left and right needle
            needleOffsetZ_L = GetStringSetting("EpoxyDispenser", "NeedleOffset", "needleOffsetZ_L", "0")
            needleOffsetZ_R = GetStringSetting("EpoxyDispenser", "NeedleOffset", "needleOffsetZ_R", "0")

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
            
            initializeInputParams
    
            initializeNodeTypeItems
    
            disableAllInputParams
    
            enableInputParams
    
            clearLockedNode
    
            displayCoOrdsTimer.Enabled = True
    
            readyStatus = True
            
            referencePtDone = False
            
            'With vision
            PicImage.Width = 461
            PicImage.height = 346
            PicImage.height = 328       'XW
   
            'VdeInitializeVision PicImage.hWnd, 461, 346
            VdeInitializeVision PicImage.hWnd, 461, 346, 1 'NNO
    
            VdeSelectCamera 2
            VdeCameraLive 1
            'LightingIntensity.Text = VdeGetLightIntensity
            
            Initialize_LightIntensity_Com        'for lightIntensity   '@$K
            Call Turn_On_LightIntensity

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
            
            selectNodeIndex = 0             'XW
            LeftNeedleValve                 'Set Default as Let_Cylinder
            LeftValve = True                'Just a default flag
            RightValve = False
            ErrorKeyIn = False
            resetTimer.Enabled = True
            ClickTimer.Enabled = True
            SetFocusTimer.Enabled = True
            Click = 0
            CheckEmergencyStop.Enabled = True
        
            setSpeed (28)
            Call SetLightIntensity(20)
        End If
    End If
    
End Sub

Private Sub NodeType_Click()
    'Origin (NYP)
    'If NodeType.Text = "----------------------" Then
    If (NodeType.Text = "----------Dot-----------") Or (NodeType.Text = "----------Line----------") Or (NodeType.Text = "----------Arc-----------") _
        Or (NodeType.Text = "---------Linker---------") Or (NodeType.Text = "--------Rectangle-------") Or (NodeType.Text = "----------Array---------") Then
        NodeType.Selected(NodeType.ListIndex - 1) = True
    End If
    
    LeftNeedle.Enabled = True
    RightNeedle.Enabled = True
    
    'Check_Node  '@$K
    
    disableAllInputParams

    If Click = 2 Then
        enableInputParams2
    Else
        enableInputParams
    End If
    
    'RotationAngle_Click
   NodeTypeNoChange = False
End Sub

Private Sub Check_Node()

    If (NodeType.Text = "Dot") Or (NodeType.Text = "Arc Start") Or (NodeType.Text = "Arc Point") _
        Or (NodeType.Text = "Arc End") Or (NodeType.Text = "Links Arc Start") Or (NodeType.Text = "Links Arc End") Or (NodeType.Text = "Links Arc Restart") Or (NodeType.Text = "RectC1") Then
        If (Click = 0) And (RightNeedle.value = True) Then
            If (NodeType.Text = "Links Arc Restart") Then
                MsgBox "System will not allow to do the program for spray valve!"
                NodeType.Selected(NodeType.ListIndex - 1) = True
            ElseIf (NodeType.Text = "Links Arc Start") Then
                MsgBox "System will not allow to do the program for spray valve!"
                NodeType.Selected(NodeType.ListIndex - 2) = True
            ElseIf (NodeType.Text = "Links Arc End") Then
                MsgBox "System will not allow to do the program for spray valve!"
                NodeType.Selected(NodeType.ListIndex - 3) = True
            Else
                Move_To_Zero
                Tilt_Off
                Call Tilt_Rotate(0)
            End If
        End If
    End If
    
End Sub

Private Sub quit_Click()
    Unload Me
End Sub

Private Sub TimerDrawStatus_Timer()
    drawStatus
End Sub

Private Sub UpDowndelay_DownClick()
    If delay.Text <> 0 Then
        delay.Text = CDbl(delay.Text) - 0.1
    End If
End Sub

Private Sub UpDowndelay_UpClick()
    delay.Text = CDbl(delay.Text) + 0.1
End Sub

Private Sub UpDownretractdelay_DownClick()
    If retractDelay.Text <> 0 Then
        retractDelay.Text = CDbl(retractDelay.Text) - 0.1
    End If
End Sub

Private Sub UpDownretractdelay_UpClick()
    retractDelay.Text = CDbl(retractDelay.Text) + 0.1
End Sub

Private Sub UpDownDispenseTime_DownClick()
    If dispenseTime.Text <> 0 Then
        dispenseTime.Text = CDbl(dispenseTime.Text) - 0.1
    End If
End Sub

Private Sub UpDowndispensetime_UpClick()
    dispenseTime.Text = CDbl(dispenseTime.Text) + 0.1
End Sub

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(y)

    VdeOnLButtonDown Button, ix, iy
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(y)

    VdeOnMouseMove Button, ix, iy
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(X)
    iy = CInt(y)

    VdeOnLButtonUp Button, CInt(X), CInt(y)
End Sub

Private Sub CheckEmergencyStop_Timer()
    Dim CheckValueX, CheckValueY, CheckValueZ As Long, DriverXYZ As Long
    
    Emergency_Stop = False
    CheckEmergencyStop.Enabled = False
    checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, CheckValueX, CheckValueY, CheckValueZ, 0))
    
    CheckValueX = (CheckValueX And &H20)
    CheckValueY = (CheckValueY And &H20)
    CheckValueZ = (CheckValueZ And &H20)
    
    If ((CheckValueX <> 0) Or (CheckValueY <> 0) Or (CheckValueZ <> 0)) Then
        Dim A As Long
        
        Servo_Off
         
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
        If fileDirty = True Then
        '    fileNotSavedForm.Show (vbModal)
            fileSave_Click
        Else
            moveToHome
            If JoggingStep.value = True Then
                setSpeed (40)
            Else
                setSpeed (jogSpeedSlider - 1)
            End If
            Ext = True
            
            'lstPattern.ListIndex = -1
            lstPattern.ListIndex = 0

            CheckEmergencyStop.Enabled = True
            Exit Sub
        End If
    
        moveToHome
        If JoggingStep.value = True Then
            setSpeed (40)
        Else
            setSpeed (jogSpeedSlider - 1)
        End If
        Ext = True
    End If

    CheckEmergencyStop.Enabled = True
End Sub

Private Sub EnableTextbox()
    'No pot dot and pot line
    'If (NodeType.ListIndex = 2) Or (NodeType.ListIndex = 3) Or (NodeType.ListIndex = 5) Then
    If (NodeType.ListIndex = 2) Then
        xRepeatNum.Enabled = True
        yRepeatNum.Enabled = True
        xDev.Enabled = True
        yDev.Enabled = True
        SubArray = False
    End If
End Sub

Private Sub dispensePtX_GotFocus()
    TabNumber = 16
    OneSevenKey = 0
End Sub

Private Sub dispensePtX_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 16) Then
        dispensePtX.Text = TextString(dispensePtX.Text, OneSevenKey)
    End If
End Sub

Private Sub dispensePtY_GotFocus()
    TabNumber = 18
    OneSevenKey = 0
End Sub

Private Sub dispensePtY_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 18) Then
        dispensePtY.Text = TextString(dispensePtY.Text, OneSevenKey)
    End If
End Sub

Private Sub dispensePtZ_GotFocus()
    TabNumber = 17
    OneSevenKey = 0
End Sub

Private Sub dispensePtZ_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 17) Then
        dispensePtZ.Text = TextString(dispensePtZ.Text, OneSevenKey)
    End If
End Sub

Private Sub dispenseTime_GotFocus()
    TabNumber = 54
    OneSevenKey = 0
End Sub

Private Sub dispenseTime_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 54) Then
        If (CDbl(dispenseTime.Text) <> 0) Then
            dispenseTime.Text = TextString(dispenseTime.Text, OneSevenKey)
        End If
    End If
End Sub

Private Sub potDepth_GotFocus()
    TabNumber = 56
    OneSevenKey = 0
End Sub

Private Sub potDepth_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 56) Then
        potDepth.Text = TextString(potDepth.Text, OneSevenKey)
    End If
    
    If (CDbl(potDepth.Text) < 0) Then
        potDepth.Text = CDbl(potDepth.Text) * (-1)
    End If
End Sub

Private Sub depthSpeed_GotFocus()
    TabNumber = 57
    OneSevenKey = 0
End Sub

Private Sub depthSpeed_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 57) Then
        depthSpeed.Text = TextString(depthSpeed.Text, OneSevenKey)
    End If
End Sub

Private Sub endDispenseHeight_GotFocus()
    TabNumber = 58
    OneSevenKey = 0
End Sub

Private Sub endDispenseHeight_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 58) Then
        endDispenseHeight.Text = TextString(endDispenseHeight.Text, OneSevenKey)
    End If
    
    'If (CDbl(endDispenseHeight.Text) < 0) Then
    '    endDispenseHeight.Text = CDbl(endDispenseHeight.Text) * (-1)
    'End If
End Sub

Private Sub delay_GotFocus()
    TabNumber = 27
    OneSevenKey = 0
End Sub

Private Sub delay_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 27) Then
        If (CDbl(delay.Text) <> 0) Then
            delay.Text = TextString(delay.Text, OneSevenKey)
        End If
    End If
End Sub

Private Sub DispenseSpeed_GotFocus()
    TabNumber = 29
    OneSevenKey = 0
End Sub

Private Sub DispenseSpeed_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 29) Then
        DispenseSpeed.Text = TextString(DispenseSpeed.Text, OneSevenKey)
    End If
End Sub

Private Sub retractDelay_GotFocus()
    TabNumber = 28
    OneSevenKey = 0
End Sub

Private Sub retractDelay_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 28) Then
        If (CDbl(retractDelay.Text) <> 0) Then
            retractDelay.Text = TextString(retractDelay.Text, OneSevenKey)
        End If
    End If
End Sub

Private Sub withdrawalSpeed_GotFocus()
    TabNumber = 26
    OneSevenKey = 0
End Sub

Private Sub withdrawalSpeed_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 26) Then
        withdrawalSpeed.Text = TextString(withdrawalSpeed.Text, OneSevenKey)
    End If
End Sub

Private Sub WithDrawalZ_GotFocus()
    If (decideWithdrawalHeight.Caption = "Teach On") Then
        decideWithdrawalHeight_Click
    ElseIf (decideMoveHeight.Caption = "Teach On") Then
        decideMoveHeight_Click
    End If
    
    TabNumber = 25
    OneSevenKey = 0
End Sub

Private Sub WithDrawalZ_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 25) Then
        WithDrawalZ.Text = TextString(WithDrawalZ.Text, OneSevenKey)
    End If
    
    'If (CDbl(WithDrawalZ.Text) < 0) Then
    '    WithDrawalZ.Text = CDbl(WithDrawalZ.Text) * (-1)
    'End If
End Sub

Private Sub moveHeight_GotFocus()
    If (decideWithdrawalHeight.Caption = "Teach On") Then
        decideWithdrawalHeight_Click
    ElseIf (decideMoveHeight.Caption = "Teach On") Then
        decideMoveHeight_Click
    End If
    
    TabNumber = 64
    OneSevenKey = 0
End Sub

Private Sub moveHeight_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 64) Then
        moveHeight.Text = TextString(moveHeight.Text, OneSevenKey)
    End If
    
    'If (CDbl(moveHeight.Text) < 0) Then
    '    moveHeight.Text = CDbl(moveHeight.Text) * (-1)
    'End If
End Sub

Private Sub xRepeatNum_GotFocus()
    TabNumber = 13
    OneSevenKey = 0
End Sub

Private Sub xRepeatNum_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 13) Then
        xRepeatNum.Text = TextString(xRepeatNum.Text, OneSevenKey)
    End If
End Sub

Private Sub yRepeatNum_GotFocus()
    TabNumber = 12
    OneSevenKey = 0
End Sub

Private Sub yRepeatNum_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 12) Then
        yRepeatNum.Text = TextString(yRepeatNum.Text, OneSevenKey)
    End If
End Sub

Private Sub xDev_GotFocus()
    TabNumber = 15
    OneSevenKey = 0
End Sub

Private Sub xDev_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 15) Then
        xDev.Text = TextString(xDev.Text, OneSevenKey)
    End If
End Sub

Private Sub yDev_GotFocus()
    TabNumber = 14
    OneSevenKey = 0
End Sub

Private Sub yDev_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 14) Then
        yDev.Text = TextString(yDev.Text, OneSevenKey)
    End If
End Sub

Private Sub txtPitch_GotFocus()
    TabNumber = 300
    OneSevenKey = 0
End Sub

Private Sub txtPitch_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 300) Then
        txtPitch.Text = TextString(txtPitch.Text, OneSevenKey)
    End If
End Sub

Private Sub txtPitch_Validate(cancel As Boolean)
        Call validateNumber(editorForm.txtPitch.Text, editorForm.lblPitch.Caption)

    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub StepDistance_GotFocus()
    TabNumber = 108
    OneSevenKey = 0
End Sub

Private Sub StepDistance_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 108) Then
        StepDistance.Text = TextString(StepDistance.Text, OneSevenKey)
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

Private Sub txtOffsetX_Validate(cancel As Boolean)
    Call validateNumber(editorForm.txtOffsetX.Text, editorForm.OffSetX.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub txtOffsetY_Validate(cancel As Boolean)
    Call validateNumber(editorForm.txtOffsetY.Text, editorForm.OffSetY.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub
Private Sub txtOffsetZ_Validate(cancel As Boolean)
    Call validateNumber(editorForm.txtOffsetZ.Text, editorForm.OffSetZ.Caption)
    'XW
    If (ErrorKeyIn = True) Then
        ErrorKeyIn = False
        cancel = True
    End If
End Sub

Private Sub txtOffsetX_GotFocus()
    TabNumber = 105
    OneSevenKey = 0
End Sub

Private Sub txtOffsetX_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 105) Then
        txtOffsetX.Text = TextString(txtOffsetX.Text, OneSevenKey)
    End If
End Sub

Private Sub txtOffsetY_GotFocus()
    TabNumber = 103
    OneSevenKey = 0
End Sub

Private Sub txtOffsetY_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 103) Then
        txtOffsetY.Text = TextString(txtOffsetY.Text, OneSevenKey)
    End If
End Sub

Private Sub txtOffsetZ_GotFocus()
    TabNumber = 104
    OneSevenKey = 0
End Sub

Private Sub txtOffsetZ_LostFocus()
    If (OneSevenKey > 1) And (TabNumber = 104) Then
        txtOffsetZ.Text = TextString(txtOffsetZ.Text, OneSevenKey)
    End If
End Sub

'This procedure will take the robort's position, (X, Y and Z) from the whold string (XW)
Private Sub GettingPosition(ByVal txtString As String)
    Dim n, m, i As Integer
    Dim step, number, Count As Long
    Dim char, Check1, check2 As String
            
    char = ""
    Check1 = ""
    check2 = ""
    n = 0
    m = 0
    ModifyOffsetX = 0
    ModifyOffsetY = 0
    ModifyOffsetZ = 0
    One = False
    step = 1
    number = Len(txtString)
            
     For Count = 1 To number
        char = char & Mid(txtString, step, 1)
        Check1 = Right(char, 2)
        check2 = Right(char, 1)
        If (Check1 = "x=") Or (Check1 = "y=") Or (Check1 = "z=") Then
            n = step
            step = step + 1
        ElseIf check2 = "," Or check2 = ")" Or (check2 = ";") Then
            m = step
            step = step + 1
            
            If (Mid(char, n - 1, 2) = "x=") Then
                ModifyOffsetX = Val(Right(char, m - n))
            ElseIf (Mid(char, n - 1, 2) = "y=") Then
                ModifyOffsetY = Val(Right(char, m - n))
            ElseIf (Mid(char, n - 1, 2) = "z=") Then
                ModifyOffsetZ = Val(Right(char, m - n))
                i = step
            End If
            
            If (check2 = ";") Or (check2 = ")") Then
                Exit Sub
            End If
        Else
            step = step + 1
        End If
    Next Count
End Sub

'The following two functions will not use now. (XW)
Private Sub DisableOffsetXYZ()
    OffSetX.Visible = False
    OffSetY.Visible = False
    OffSetZ.Visible = False
    txtOffsetX.Enabled = False
    txtOffsetY.Enabled = False
    txtOffsetZ.Enabled = False
End Sub

Private Sub EnableOffsetXYZ()
    OffSetX.Visible = True
    OffSetY.Visible = True
    OffSetZ.Visible = True
    txtOffsetX.Enabled = True
    txtOffsetY.Enabled = True
    txtOffsetZ.Enabled = True
End Sub

Private Sub Close_AllTimer()
    ClickTimer.Enabled = False
    SetFocusTimer.Enabled = False
    MouseMovement.Enabled = False
    resetTimer.Enabled = False
    displayCoOrdsTimer.Enabled = False
    CheckEmergencyStop.Enabled = False
    FudicialTimer.Enabled = False
    TimerDrawStatus.Enabled = False
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''
'                                               '
'   This is only for "Vision and Needle teach"  '
'        (Need to change the option type)       '
'                                               '
'''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub VisionTeach_Click()
    '''''''''''''''''''''
    '   For SCS4000M    '
    '''''''''''''''''''''
    'If (VisionTeach.value = 1) Then
    '    modifyPt.Enabled = True
    '    If (RightNeedle.value = True) Then
    '        If (RotationAngle.Text <> "None") Then
    '            MsgBox "'Camera Teach'is not allowed for tilting and rotation."
    '            VisionTeach.value = 0
    '        Else
    '            checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    '            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
    '                DoEvents
    '            Loop
    '            RotationAngle.Text = "None"
    '            Tilt_Off
    '            Call Tilt_Rotate(0)
    '        End If
    '    End If
    'Else
    '    If (LeftNeedle.value = True) Then
    '        modifyPt.Enabled = False
    '    End If
    'End If
    
    '''''''''''''''''''''
    '   For SCS4000N    '
    '''''''''''''''''''''
    If (VisionTeach.value = 1) Then
        SetLightIntensity (Val(LightingIntensity.Text))
        Move_To_Zero
    Else
        SetLightIntensity (0)
    End If
    
End Sub

Private Function Previous_Angle(ByVal previous_program As String) As String
    Dim angle() As String, words() As String
    
    angle() = Split(previous_program, "=")
    words() = Split(previous_program, "(")
    
    If (words(0) = "lineStart") Then
        Previous_Angle = Left(angle(4), Len(angle(4)) - 1)
    ElseIf (words(0) = "   linksLinePoint") Then
        Previous_Angle = Left(angle(5), Len(angle(5)) - 1)
    ElseIf (words(0) = "   linksArcEnd") Then
        Previous_Angle = Left(angle(5), Len(angle(5)) - 1)
    Else
        Previous_Angle = "No"
    End If
End Function


