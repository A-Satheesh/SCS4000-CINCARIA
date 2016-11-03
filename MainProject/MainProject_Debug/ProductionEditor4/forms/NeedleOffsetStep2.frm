VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form NeedleOffsetStep2 
   Caption         =   "Determine Needle Offset Step 2 of 4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   615
      Left            =   360
      OleObjectBlob   =   "NeedleOffsetStep2.frx":0000
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Proceed"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   2160
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "NeedleOffsetStep2.frx":0090
      Top             =   0
   End
End
Attribute VB_Name = "NeedleOffsetStep2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    setSpeed (60)
    PTPToXYZ 4466, -50000, 0
    PTPToXYZ 4466, -50000, 13326
    Unload Me
    NeedleOffsetStep3.Show

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")

    Skin1.ApplySkin Me.hWnd

End Sub
