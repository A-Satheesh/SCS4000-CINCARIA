VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form RetryFudicial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retry Fudicial ?"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4920
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "RetryFudicial.frx":0000
      Top             =   360
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   1320
      OleObjectBlob   =   "RetryFudicial.frx":0234
      TabIndex        =   2
      Top             =   720
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Skip"
      Height          =   495
      Left            =   2760
      TabIndex        =   1
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Countine"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1455
   End
End
Attribute VB_Name = "RetryFudicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    visionRetry = False
    doneFudicial = False
    Unload Me
End Sub

Private Sub Command2_Click()
    abortValue = True
    visionRetry = False
    doneFudicial = True
    executionForm.cmdStartButton.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

