VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form NeedleOffsetStep3 
   Caption         =   "Determine Needle Offset Step 3 of 4"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   975
      Left            =   480
      OleObjectBlob   =   "NeedleOffsetStep3.frx":0000
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   2040
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "NeedleOffsetStep3.frx":00B0
      Top             =   0
   End
End
Attribute VB_Name = "NeedleOffsetStep3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    PTPToXYZ 4466, -50000, 0
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), 0
    PTPToXYZ GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationX", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationY", "0"), GetStringSetting("EpoxyDispenser", "NeedleOffset", "calibrationZ", "0")
    Unload Me
    NeedleOffsetStep4.Show
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")

    Skin1.ApplySkin Me.hWnd

End Sub
