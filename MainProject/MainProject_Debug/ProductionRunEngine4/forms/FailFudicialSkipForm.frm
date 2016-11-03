VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FailFudicialSkipForm 
   Caption         =   "Form1"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4800
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Click to Abort"
      Height          =   495
      Left            =   1560
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   1440
      OleObjectBlob   =   "FailFudicialSkipForm.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FailFudicialSkipForm.frx":0092
      Top             =   1320
   End
End
Attribute VB_Name = "FailFudicialSkipForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    visionRetry = False
    doneFudicial = False
    Unload Me
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.ApplySkin Me.hWnd
End Sub
