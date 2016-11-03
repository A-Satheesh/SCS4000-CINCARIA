VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form editorForm 
   BackColor       =   &H8000000B&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Epoxy Editor"
   ClientHeight    =   10530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15240
   FillColor       =   &H00FFFFFF&
   FillStyle       =   0  'Solid
   Icon            =   "editorForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   702
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1016
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton AbortNeedleOffset 
      Caption         =   "Abort"
      Height          =   495
      Left            =   10680
      TabIndex        =   97
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel NeedleOffsetErrMsg 
      Height          =   615
      Left            =   7920
      OleObjectBlob   =   "editorForm.frx":08CA
      TabIndex        =   98
      Top             =   9840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin ComCtl2.UpDown UpDownLighting 
      Height          =   255
      Left            =   7560
      TabIndex        =   94
      Top             =   9840
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   327681
      BuddyControl    =   "LightingIntensity"
      BuddyDispid     =   196610
      OrigLeft        =   504
      OrigTop         =   656
      OrigRight       =   520
      OrigBottom      =   673
      Max             =   100
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox LightingIntensity 
      Height          =   285
      Left            =   7080
      TabIndex        =   93
      Text            =   "50"
      Top             =   9840
      Width           =   495
   End
   Begin ACTIVESKINLibCtl.SkinLabel FudMsgText 
      Height          =   735
      Left            =   10680
      OleObjectBlob   =   "editorForm.frx":09A2
      TabIndex        =   92
      Top             =   9840
      Width           =   2175
   End
   Begin VB.Timer FudicialTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   720
      Top             =   1680
   End
   Begin VB.CommandButton CancelFud 
      Caption         =   "Cancel Fudicial"
      Height          =   495
      Left            =   8880
      TabIndex        =   90
      Top             =   9840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox PicImage 
      Height          =   4410
      Left            =   7320
      ScaleHeight     =   290
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   372
      TabIndex        =   87
      Top             =   4680
      Width           =   5640
   End
   Begin VB.CommandButton decideDispensePt 
      Caption         =   "Teach On"
      Height          =   375
      Left            =   3600
      TabIndex        =   86
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton translateButton 
      Caption         =   "Translate"
      Height          =   495
      Left            =   12000
      TabIndex        =   85
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Timer displayCoOrdsTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1680
   End
   Begin VB.ListBox NodeType 
      Height          =   1035
      Left            =   1920
      TabIndex        =   84
      Top             =   120
      Width           =   2655
   End
   Begin VB.Timer TimerDrawStatus 
      Interval        =   1000
      Left            =   -3720
      Top             =   4920
   End
   Begin VB.CheckBox dispenseOnOff 
      Alignment       =   1  'Right Justify
      Caption         =   "Dispense On"
      Height          =   255
      Left            =   480
      TabIndex        =   83
      Top             =   4560
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel27 
      Height          =   255
      Left            =   5760
      OleObjectBlob   =   "editorForm.frx":0A00
      TabIndex        =   82
      Top             =   240
      Width           =   1215
   End
   Begin VB.PictureBox PictureError 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6600
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   81
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox PictureBusy 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   6000
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   80
      Top             =   480
      Width           =   615
   End
   Begin VB.PictureBox PictureReady 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   5400
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   79
      Top             =   480
      Width           =   615
   End
   Begin MSComCtl2.UpDown UpDownYRepeatNum 
      Height          =   285
      Left            =   3240
      TabIndex        =   78
      Top             =   6480
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2
      BuddyControl    =   "yRepeatNum"
      BuddyDispid     =   196651
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
      Left            =   3240
      TabIndex        =   77
      Top             =   6720
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2
      BuddyControl    =   "xRepeatNum"
      BuddyDispid     =   196650
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
      Left            =   3240
      TabIndex        =   76
      Top             =   5520
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "withdrawalSpeed"
      BuddyDispid     =   196637
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
      Left            =   3240
      TabIndex        =   75
      Top             =   5280
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDispenseSpeed 
      Height          =   285
      Left            =   3240
      TabIndex        =   74
      Top             =   3960
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   10
      BuddyControl    =   "DispenseSpeed"
      BuddyDispid     =   196634
      OrigLeft        =   3240
      OrigTop         =   5336
      OrigRight       =   3480
      OrigBottom      =   5741
      Max             =   70
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDelay 
      Height          =   285
      Left            =   3240
      TabIndex        =   73
      Top             =   3600
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "delay"
      BuddyDispid     =   196636
      OrigLeft        =   3240
      OrigTop         =   4765
      OrigRight       =   3480
      OrigBottom      =   5170
      Increment       =   0
      Max             =   100
      Wrap            =   -1  'True
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownEndDispenseHeight 
      Height          =   255
      Left            =   3240
      TabIndex        =   72
      Top             =   3240
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDepthSpeed 
      Height          =   255
      Left            =   3240
      TabIndex        =   71
      Top             =   3000
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownPotDepth 
      Height          =   255
      Left            =   3240
      TabIndex        =   70
      Top             =   2760
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UpDownDispenseTime 
      Height          =   255
      Left            =   3240
      TabIndex        =   69
      Top             =   2280
      Width           =   240
      _ExtentX        =   450
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton loadPartArray 
      Caption         =   "Load"
      Height          =   255
      Left            =   5760
      TabIndex        =   68
      Top             =   7680
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel pathFileNameLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0A78
      TabIndex        =   67
      Top             =   7680
      Width           =   1215
   End
   Begin VB.TextBox PathFileName 
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   7680
      Width           =   3735
   End
   Begin VB.CommandButton decideMoveHeight 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3480
      TabIndex        =   65
      Top             =   6000
      Width           =   1095
   End
   Begin VB.TextBox moveHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   64
      Top             =   6000
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel moveHeightLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0AF0
      TabIndex        =   63
      Top             =   6000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   240
      TabIndex        =   62
      Top             =   4440
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel delayLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0B64
      TabIndex        =   61
      Top             =   3600
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel endDispenseHeightLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0BDA
      TabIndex        =   60
      Top             =   3240
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel depthSpeedLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0C56
      TabIndex        =   59
      Top             =   3000
      Width           =   1215
   End
   Begin VB.TextBox endDispenseHeight 
      Height          =   285
      Left            =   1920
      TabIndex        =   58
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox depthSpeed 
      Height          =   285
      Left            =   1920
      TabIndex        =   57
      Top             =   3000
      Width           =   1335
   End
   Begin VB.TextBox potDepth 
      Height          =   285
      Left            =   1920
      TabIndex        =   56
      Top             =   2760
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel PotDepthLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0CCA
      TabIndex        =   55
      Top             =   2760
      Width           =   1215
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
      Left            =   1920
      TabIndex        =   54
      Text            =   "1.0"
      Top             =   2280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel DispenseTimeLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0D42
      TabIndex        =   53
      Top             =   2280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel20 
      Height          =   255
      Left            =   6120
      OleObjectBlob   =   "editorForm.frx":0DBA
      TabIndex        =   52
      Top             =   9600
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "editorForm.frx":0E1E
      TabIndex        =   51
      Top             =   9600
      Width           =   375
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel18 
      Height          =   495
      Left            =   360
      OleObjectBlob   =   "editorForm.frx":0E82
      TabIndex        =   50
      Top             =   9960
      Width           =   735
   End
   Begin ComctlLib.Slider jogSpeedSlider 
      Height          =   375
      Left            =   960
      TabIndex        =   49
      Top             =   9960
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
   Begin VB.CommandButton trackButton 
      Caption         =   "Track"
      Height          =   495
      Left            =   5280
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel NodeLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":0EF2
      TabIndex        =   46
      Top             =   1320
      Width           =   1095
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   255
      Left            =   12000
      OleObjectBlob   =   "editorForm.frx":0F5A
      TabIndex        =   45
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
      Height          =   255
      Left            =   10320
      OleObjectBlob   =   "editorForm.frx":0FBA
      TabIndex        =   44
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
      Height          =   255
      Left            =   8640
      OleObjectBlob   =   "editorForm.frx":101A
      TabIndex        =   43
      Top             =   480
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel yRepeatNumLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":107A
      TabIndex        =   42
      Top             =   6720
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel yDevLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":10F2
      TabIndex        =   41
      Top             =   7200
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.SkinLabel xRepeatNumLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":1166
      TabIndex        =   40
      Top             =   6480
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel xDevLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":11D8
      TabIndex        =   39
      Top             =   6960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel retractDelayLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":124C
      TabIndex        =   38
      Top             =   5280
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel withDrawalSpeedLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":12C4
      TabIndex        =   37
      Top             =   5520
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel withDrawalZLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":133C
      TabIndex        =   36
      Top             =   5760
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel dispenseSpeedLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":13B6
      TabIndex        =   35
      Top             =   3960
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.SkinLabel dispensePtZLabel 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "editorForm.frx":142C
      TabIndex        =   34
      Top             =   1800
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel YLabel 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "editorForm.frx":148C
      TabIndex        =   33
      Top             =   1560
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel XLabel 
      Height          =   255
      Left            =   1680
      OleObjectBlob   =   "editorForm.frx":14EC
      TabIndex        =   32
      Top             =   1320
      Width           =   135
   End
   Begin ACTIVESKINLibCtl.SkinLabel NodeTypeLabel 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":154C
      TabIndex        =   31
      Top             =   120
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "editorForm.frx":15C4
      Top             =   480
   End
   Begin VB.CommandButton Home 
      Caption         =   "Home"
      Height          =   495
      Left            =   10320
      TabIndex        =   30
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox DispenseSpeed 
      Height          =   285
      Left            =   1920
      TabIndex        =   29
      Text            =   "10"
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox retractDelay 
      Height          =   285
      Left            =   1920
      TabIndex        =   28
      Text            =   "1.0"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.TextBox delay 
      Height          =   285
      Left            =   1920
      TabIndex        =   27
      Text            =   "1.0"
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox withdrawalSpeed 
      Height          =   285
      Left            =   1920
      TabIndex        =   26
      Text            =   "10"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.TextBox WithDrawalZ 
      Height          =   285
      Left            =   1920
      TabIndex        =   25
      Top             =   5760
      Width           =   1335
   End
   Begin VB.TextBox xCoOrd 
      Height          =   375
      Left            =   8760
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox yCoOrd 
      Height          =   375
      Left            =   10440
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox zCoOrd 
      Height          =   375
      Left            =   12120
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton deletePt 
      Caption         =   "Delete Node"
      Height          =   495
      Left            =   11160
      TabIndex        =   21
      Top             =   3840
      Width           =   1455
   End
   Begin VB.CommandButton addPt 
      Caption         =   "Add Node"
      Height          =   495
      Left            =   7200
      TabIndex        =   20
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton decideWithdrawalHeight 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3480
      TabIndex        =   19
      Top             =   5760
      Width           =   1095
   End
   Begin VB.TextBox dispensePtY 
      Height          =   285
      Left            =   1920
      TabIndex        =   18
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox dispensePtZ 
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Top             =   1800
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
      Left            =   1920
      TabIndex        =   16
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox xDev 
      Height          =   285
      Left            =   1920
      TabIndex        =   15
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox yDev 
      Height          =   285
      Left            =   1920
      TabIndex        =   14
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox xRepeatNum 
      Height          =   285
      Left            =   1920
      TabIndex        =   13
      Text            =   "1"
      Top             =   6720
      Width           =   1335
   End
   Begin VB.TextBox yRepeatNum 
      Height          =   285
      Left            =   1920
      TabIndex        =   12
      Text            =   "1"
      Top             =   6480
      Width           =   1335
   End
   Begin VB.CommandButton decideXYDev 
      Caption         =   "Teach Off"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton modifyPt 
      Caption         =   "Modify Node"
      Height          =   495
      Left            =   9240
      TabIndex        =   10
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton fileNew 
      Caption         =   "New"
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton fileLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton fileSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   8640
      TabIndex        =   7
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton zPlus 
      Caption         =   "Z+"
      Height          =   615
      Left            =   4440
      TabIndex        =   6
      Top             =   9120
      Width           =   855
   End
   Begin VB.CommandButton zMinus 
      Caption         =   "Z-"
      Height          =   615
      Left            =   4440
      TabIndex        =   5
      Top             =   8280
      Width           =   855
   End
   Begin VB.CommandButton xPlus 
      Caption         =   "X -"
      Height          =   615
      Left            =   720
      TabIndex        =   4
      Top             =   8640
      Width           =   735
   End
   Begin VB.CommandButton yMinus 
      Caption         =   "Y -"
      Height          =   615
      Left            =   1680
      TabIndex        =   3
      Top             =   8880
      Width           =   735
   End
   Begin VB.CommandButton yPlus 
      Caption         =   "Y +"
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   8160
      Width           =   735
   End
   Begin VB.CommandButton xMinus 
      Caption         =   "X +"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   8640
      Width           =   735
   End
   Begin VB.ListBox lstPattern 
      Height          =   1425
      ItemData        =   "editorForm.frx":17F8
      Left            =   5280
      List            =   "editorForm.frx":17FA
      TabIndex        =   0
      Top             =   2160
      Width           =   7575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Robot Position"
      Height          =   735
      Left            =   8520
      TabIndex        =   47
      Top             =   240
      Width           =   5175
   End
   Begin VB.CommandButton TeachFudicialPt 
      Caption         =   "Teach Fudicial Points"
      Height          =   495
      Left            =   10680
      TabIndex        =   88
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton Calibrate1 
      Caption         =   "Calibrate"
      Height          =   495
      Left            =   12960
      TabIndex        =   96
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Calibrate 
      Caption         =   "Calibrate"
      Height          =   495
      Left            =   12960
      TabIndex        =   95
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton FindNeedleOffset 
      Caption         =   "Needle Calibration"
      Height          =   495
      Left            =   12960
      TabIndex        =   91
      Top             =   9840
      Width           =   1935
   End
   Begin VB.CommandButton NextFudStep 
      Caption         =   "Next"
      Height          =   495
      Left            =   12960
      TabIndex        =   89
      Top             =   9840
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSCommLib.MSComm mscomLighIntensity 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
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

Private Sub AbortNeedleOffset_Click()
        AbortNeedleOffset.Visible = False
        NeedleOffsetErrMsg.Visible = False
        Calibrate1.Visible = False
        FindNeedleOffset.Visible = True
        TeachFudicialPt.Visible = True
        'VdeSelectCamera 2
    
End Sub

Private Sub Calibrate_Click()
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    
    PTPToXYZ xDatum, yDatum, systemMoveHeight
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), systemMoveHeight
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationZ", "0")
    Calibrate.Visible = False
    FindNeedleOffset.Visible = False
    Calibrate1.Visible = True
    'VdeSelectCamera 1
    'VdeReadSettings ("VisionSetup.txt")

End Sub

Private Sub Calibrate1_Click()
    Dim OffSetX, OffSetY As Double
    
    'returncode = VdeFindNeedleOffset(offsetX, offsetY)
    
    If returncode = 1 Then
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "XOff", OffSetX
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "YOff", OffSetY
        AbortNeedleOffset_Click
    Else
        AbortNeedleOffset.Visible = True
        NeedleOffsetErrMsg.Visible = True
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
    'step = VdeTeachRefPtDlg(VisionDlgCancel, s)
    FudMsgText.Caption = ""
    CancelFud.Visible = False
    NextFudStep.Visible = False
    TeachFudicialPt.Visible = True
    FindNeedleOffset.Visible = True

    'VdeCameraLive 1
End Sub

Private Sub FindNeedleOffset_Click()
    TeachFudicialPt.Visible = False
    
    setSpeed (60)
    
    Dim xDatum, yDatum, zDatum As Long
    
    xDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "xDatum", "0"), X_axis)
    yDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "yDatum", "0"), Y_axis)
    zDatum = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "zDatum", "0"), Z_axis)

    PTPToXYZ xDatum, yDatum, systemMoveHeight
    PTPToXYZ xDatum, yDatum, zDatum
    FindNeedleOffset.Visible = False
    Calibrate.Visible = True
    FudMsgText.Caption = "Adjust Needle to Datum Height and Click>"

End Sub

Private Sub FudicialTimer_Timer()
    FudicialTimer.Enabled = False
    'step = VdeTeachRefPtDlg(VisionDlgOnTimer, s)
    FudicialTimer.Enabled = True
End Sub

Private Sub jogSpeedSlider_Change()
    setSpeed (jogSpeedSlider.Value - 1)

End Sub

Private Sub LightingIntensity_Change()
    'VdeSetLightIntensity LightingIntensity.Text
End Sub

Private Sub mnucut_click()
    If (editorForm.lstPattern.ListIndex <> -1) Then
        tempCutPasteString = editorForm.lstPattern.List(editorForm.lstPattern.ListIndex)
        editorForm.lstPattern.RemoveItem (editorForm.lstPattern.ListIndex)
    End If
End Sub

Private Sub mnucopy_click()
    If (editorForm.lstPattern.ListIndex <> -1) Then
        tempCutPasteString = editorForm.lstPattern.List(editorForm.lstPattern.ListIndex)
    End If
End Sub

Private Sub mnupaste_click()
    If (editorForm.lstPattern.ListIndex <> -1) Then
        Call lstPattern.AddItem(tempCutPasteString, editorForm.lstPattern.ListIndex)
    End If

End Sub

Private Sub addPt_Click()
    If (lstPattern.ListIndex = -1) Then
        lstPattern.AddItem (processAddNode)
    Else
        clearLockedNode
        Call lstPattern.AddItem(processAddNode, lstPattern.ListIndex)
    End If
    
    fileDirty = True

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
        End If
    End If
    
End Sub

Private Sub deletePt_Click()
    If (lstPattern.ListIndex = -1) Then
        NodeError.Show (vbModal)
    Else
        clearLockedNode
        lstPattern.RemoveItem (lstPattern.ListIndex)
        selectNodeIndex = -1
    End If
    
    fileDirty = True
    
End Sub

Private Sub dispenseptx_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtX.Text, editorForm.XLabel.Caption)
End Sub
Private Sub dispensepty_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtY.Text, editorForm.YLabel.Caption)
End Sub
Private Sub dispenseptz_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispensePtZ.Text, editorForm.dispensePtZLabel.Caption)
End Sub
Private Sub dispenseTime_Validate(cancel As Boolean)
    Call validateNumber(editorForm.dispenseTime.Text, editorForm.DispenseTimeLabel.Caption)
End Sub

Private Sub displayCoOrdsTimer_Timer()
    'XW
    'It should be colsed while change to Execution form
    If (CloseBoard = False) Then
        displayCoOrds
    End If
End Sub

Private Sub fileNew_Click()

    proceedWithAction = True

    If fileDirty = True Then
        fileNotSavedForm.Show (vbModal)
    End If
    
    If proceedWithAction = True Then
        lstPattern.Clear
        selectNodeIndex = -1
        'editorForm.Caption = "Epoxy Editor"
        editorForm.Caption = "Profile Editor"
        systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
        systemMoveHeight = systemMoveHeight * (-1)
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
        readyStatus = True
    End If
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    Dim DriverXYZ As Long
    'proceedWithAction = True
    
    'If fileDirty = True Then
        'fileNotSavedForm.Show (vbModal)
    'End If
    
    'If proceedWithAction = True Then
        WindowFree lstPattern.hWnd
        
        returncode = P1240FreeContiBuf(0)  'XW
        
        'Both needles go down
        Leftslider_go_down
        
        Call Sleep(0.3)
        
        Servo_Off
        
        unInitializeBoard
    'end if
    
End Sub

Private Sub Home_Click()

    Dim tempJogSpeed As Integer
    
    tempJogSpeed = jogSpeedSlider.Value

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
    
    'moveToHome
    
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, -1000, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
    Loop
    
    checkSuccess (P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, 1000, -1000, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> Success)
    Loop

    'returncode = P1240MotHome(boardNum, X_axis Or Y_axis Or Z_axis)
    Call moveToHome
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis Or Z_axis) <> Success)
    Loop
    
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
    
    jogSpeedSlider.Value = tempJogSpeed
    
End Sub

'Private Sub jogSpeedSlider_Click()
    'setSpeed (jogSpeedSlider.Value - 1)
'End Sub
Private Sub loadPartArray_Click()
    repeatPatternFileLoad.Show (vbModal)
End Sub

Private Sub lstPattern_Click()
    If lstPattern.ListIndex <> -1 Then
        If selectNodeIndex = lstPattern.ListIndex Then
            lstPattern.ListIndex = -1
            selectNodeIndex = -1
        End If
    End If
    selectNodeIndex = lstPattern.ListIndex
    trackButton_Click
End Sub

Private Sub NextFudStep_Click()
Dim x1 As Double
Dim y1 As Double
Dim x2 As Double
Dim y2 As Double
Dim tempStr As String


    FudicialTimer.Enabled = False
    'step = VdeTeachRefPtDlg(VisionDlgNext, s)
    'step = VdeTeachRefPtDlg(VisionDlgNext, s)
    
    If (step = VisionDlgToFinish) Then
        step = prevStep + 1
    ElseIf (step = VisionDlgFinish) Then
        'VdeGetRefPtPos x1, y1, x2, y2
          
        tempStr = "fudicial(x=" & convertToPulses(x1, X_axis) & ", y=" & convertToPulses(y1, Y_axis) & "; x=" & convertToPulses(x2, X_axis) & ", y=" & convertToPulses(y2, Y_axis) & "; " & Chr(34) & editorForm.Caption & "pat" & Chr(34) & "; " & LightingIntensity.Text & ")"
        
        Call editorForm.lstPattern.AddItem(tempStr, 0)

        CancelFud_Click
        Exit Sub
    End If

    prevStep = step
    FudMsgText.Caption = "Adjust ROI and Search Area for 2nd Fudicial Point and Click>"
    
    FudicialTimer.Enabled = True
End Sub

Private Sub TeachFudicialPt_Click()
    
    If editorForm.Caption = "Epoxy Editor" Then
        Call fileSave_Click
    Else
        patFile = editorForm.Caption & "pat"
        NextFudStep.Visible = True
        'CancelFud.Visible = True
        TeachFudicialPt.Visible = False
        FindNeedleOffset.Visible = False
        
        'VdeSetRefPtFilename patFile
        'step = VdeTeachRefPtDlg(VisionDlgInit, s)
        FudMsgText.Caption = "Adjust ROI and Search Area for 1st Fudicial Point and Click>"
        s = ""
        prevStep = 1
        FudicialTimer.Enabled = True
    End If
End Sub

Private Sub trackButton_Click()
    If (lstPattern.ListIndex = -1) Then
        'Commented to give constant tracking
        'NodeError.Show (vbModal)
    Else
        clearLockedNode
        doTrack (lstPattern.List(lstPattern.ListIndex) & vbNewLine)
        jogSpeedSlider.Value = 50
        setSpeed (jogSpeedSlider.Value - 1)

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

        Call PTPToXYZ(CLng(dispensePtX.Text) - needleOffsetX, CLng(dispensePtY.Text) - needleOffsetY, CLng(dispensePtZ.Text))
    
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
    End If
    
End Sub

Private Sub modifyPt_Click()
    If (lstPattern.ListIndex = -1) Then
        NodeError.Show (vbModal)
    Else
        clearLockedNode
        Dim i As Long
        i = lstPattern.ListIndex
        lstPattern.RemoveItem (lstPattern.ListIndex)
        Call lstPattern.AddItem(processAddNode, i)
    End If
    
    fileDirty = True
End Sub

Private Sub potDepth_Validate(cancel As Boolean)
    Call validateNumber(editorForm.potDepth.Text, editorForm.PotDepthLabel.Caption)
End Sub
Private Sub depthspeed_Validate(cancel As Boolean)
    Call validateNumber(editorForm.depthSpeed.Text, editorForm.depthSpeedLabel.Caption)
End Sub
Private Sub enddispenseheight_Validate(cancel As Boolean)
    Call validateNumber(editorForm.endDispenseHeight.Text, editorForm.endDispenseHeightLabel.Caption)
End Sub
Private Sub delay_Validate(cancel As Boolean)
    Call validateNumber(editorForm.delay.Text, editorForm.delayLabel.Caption)
End Sub
Private Sub dispensespeed_Validate(cancel As Boolean)
    Call validateNumber(editorForm.DispenseSpeed.Text, editorForm.dispenseSpeedLabel.Caption)
    If CLng(editorForm.DispenseSpeed.Text) > 500 Then
        editorForm.DispenseSpeed.Text = "500"
    End If
End Sub

Private Sub retractdelay_Validate(cancel As Boolean)
    Call validateNumber(editorForm.retractDelay.Text, editorForm.retractDelayLabel.Caption)
End Sub

Private Sub translateButton_Click()
    clearLockedNode
    If editorForm.Caption = "Epoxy Editor" Then
        fileNotLoadedError.Show (vbModal)
    Else
        If fileDirty = True Then
            fileNotSavedForm.Show (vbModal)
            If proceedWithAction = True Then
                translateForm.startTranslate
            End If
        Else
            translateForm.startTranslate
        End If
    
    End If
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
End Sub
Private Sub withdrawalz_Validate(cancel As Boolean)
    Call validateNumber(editorForm.WithDrawalZ.Text, editorForm.withDrawalZLabel.Caption)
End Sub
Private Sub moveheight_Validate(cancel As Boolean)
    Call validateNumber(editorForm.moveHeight.Text, editorForm.moveHeightLabel.Caption)
End Sub

Private Sub xCoOrd_GotFocus()
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempX = xCoOrd.Text
        xCoOrd.Locked = False
    End If
End Sub

Private Sub xCoOrd_LostFocus()
    If busyStatus = False Then
        If xCoOrd.Text <> tempX Then
            Call PTPToXYZ(convertToPulses(xCoOrd.Text, X_axis), convertToPulses(yCoOrd.Text, Y_axis), convertToPulses(zCoOrd.Text, Z_axis))
        End If
        displayCoOrdsTimer.Enabled = True
        xCoOrd.Locked = True
    End If
End Sub

Private Sub yCoOrd_GotFocus()
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempY = yCoOrd.Text
        yCoOrd.Locked = False
    End If
End Sub

Private Sub yCoOrd_LostFocus()
    If busyStatus = False Then
        If yCoOrd.Text <> tempY Then
            Call PTPToXYZ(convertToPulses(xCoOrd.Text, X_axis), convertToPulses(yCoOrd.Text, Y_axis), convertToPulses(zCoOrd.Text, Z_axis))
        End If
        displayCoOrdsTimer.Enabled = True
        yCoOrd.Locked = True
    End If
End Sub
Private Sub zCoOrd_GotFocus()
    If busyStatus = False Then
        displayCoOrdsTimer.Enabled = False
        tempZ = zCoOrd.Text
        zCoOrd.Locked = False
    End If
End Sub

Private Sub zCoOrd_LostFocus()
    If busyStatus = False Then
        If zCoOrd.Text <> tempZ Then
            Call PTPToXYZ(convertToPulses(xCoOrd.Text, X_axis), convertToPulses(yCoOrd.Text, Y_axis), convertToPulses(zCoOrd.Text, Z_axis))
        End If
        displayCoOrdsTimer.Enabled = True
        zCoOrd.Locked = True
    End If
End Sub

Private Sub xMinus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, X_axis, 0))
End Sub

Private Sub xPlus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, X_axis, 1))
End Sub
Private Sub xMinus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
End Sub

Private Sub xPlus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
End Sub
Private Sub yMinus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, Y_axis, 2))
End Sub

Private Sub yPlus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, Y_axis, 0))
End Sub
Private Sub yMinus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
End Sub

Private Sub yPlus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
End Sub
Private Sub zMinus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, Z_axis, 4))
End Sub

Private Sub zPlus_mouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotCmove(boardNum, Z_axis, 0))
End Sub
Private Sub zMinus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
End Sub

Private Sub zPlus_mouseup(Button As Integer, Shift As Integer, x As Single, y As Single)
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
End Sub
Private Sub xrepeatnum_Validate(cancel As Boolean)
    Call validateNumber(editorForm.xRepeatNum.Text, editorForm.xRepeatNumLabel.Caption)
End Sub
Private Sub yrepeatnum_Validate(cancel As Boolean)
    Call validateNumber(editorForm.yRepeatNum.Text, editorForm.yRepeatNumLabel.Caption)
End Sub
Private Sub xdev_Validate(cancel As Boolean)
    Call validateNumber(editorForm.xDev.Text, editorForm.xDevLabel.Caption)
End Sub
Private Sub ydev_Validate(cancel As Boolean)
    Call validateNumber(editorForm.yDev.Text, editorForm.yDevLabel.Caption)
End Sub

Private Sub fileLoad_Click()

    proceedWithAction = True
    
    If fileDirty = True Then
        fileNotSavedForm.Show (vbModal)
    End If
    
    If proceedWithAction = True Then

        clearLockedNode
        initializeInputParams
        fileLoadForm.Show (vbModal)
    End If
    
End Sub

Private Sub fileSave_Click()
    clearLockedNode
    fileSaveForm.Show (vbModal)
End Sub

Private Sub Form_Load()

        'Skin1.LoadSkin (".\skin\epoxySkin.skn")
        Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
        Skin1.ApplySkin Me.hWnd
    
        WindowHook Me.lstPattern.hWnd
        
        'Not to show(Two e-stop form)       'XW
        TimerDrawStatus.Enabled = False
        readyStatus = False
        busyStatus = False
        errorStatus = False

        referenceX = 0
        referenceY = 0
        referenceZ = 0
    
        fileDirty = False

        selectNodeIndex = -1
    
        readRegistryOptions
    
        determineProfile
 
        systemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
        systemMoveHeight = systemMoveHeight * (-1)
        systemTrackMoveHeight = 0
        
        systemHomeX = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "xSystemHome", "0")), X_axis)
        systemHomeY = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "ySystemHome", "0")), Y_axis)
        systemHomeZ = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "zSystemHome", "0")), Z_axis)
 
        initializeInputParams
    
        initializeNodeTypeItems
    
        disableAllInputParams
    
        enableInputParams
    
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
    
        clearLockedNode
    
        displayCoOrdsTimer.Enabled = True
    
        readyStatus = True
    
        'PicImage.Width = 461
        'PicImage.Height = 346
        'VdeInitializeVision PicImage.hWnd, 461, 346

        'VdeSelectCamera 2
        'VdeCameraLive 1
       
        'moveToHome
     
        'If ballScrew = 1 Then
            'jogSpeedSlider.Max = 85
        'End If
 
End Sub

Private Sub NodeType_Click()
    If NodeType.Text = "----------------------" Then
        NodeType.Selected(NodeType.ListIndex - 1) = True
    End If
    disableAllInputParams
    'dispensePtX.Text = xCoOrd.Text
    'dispensePtY.Text = yCoOrd.Text
    'dispensePtZ.Text = zCoOrd.Text
    enableInputParams
End Sub

Private Sub quit_Click()
    Unload Me
End Sub

Private Sub TimerDrawStatus_Timer()
    'XW
    If (CloseBoard = False) Then
        drawStatus
    End If
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

Private Sub picImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(x)
    iy = CInt(y)

    'VdeOnLButtonDown Button, ix, iy
End Sub

Private Sub picImage_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(x)
    iy = CInt(y)

    'VdeOnMouseMove Button, ix, iy
End Sub

Private Sub picImage_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim ix As Integer
Dim iy As Integer
    ix = CInt(x)
    iy = CInt(y)

    'VdeOnLButtonUp Button, CInt(x), CInt(y)
End Sub

