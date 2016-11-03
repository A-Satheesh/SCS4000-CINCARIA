VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form NodeError 
   Caption         =   "Node Error"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "NodeError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3720
      OleObjectBlob   =   "NodeError.frx":030A
      Top             =   1920
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   372
      Left            =   600
      OleObjectBlob   =   "NodeError.frx":053E
      TabIndex        =   1
      Top             =   840
      Width           =   3492
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "NodeError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")

    Skin1.ApplySkin Me.hWnd

End Sub

Private Sub ok_Click()
    Unload Me
End Sub
