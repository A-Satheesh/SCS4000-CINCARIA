VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form FileSaveOKForm 
   Caption         =   "File Save OK"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "FileSaveOKForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   Begin ACTIVESKINLibCtl.SkinLabel filePathLabel 
      Height          =   372
      Left            =   600
      OleObjectBlob   =   "FileSaveOKForm.frx":030A
      TabIndex        =   2
      Top             =   1320
      Width           =   4812
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "FileSaveOKForm.frx":0368
      Top             =   2280
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   492
      Left            =   1680
      OleObjectBlob   =   "FileSaveOKForm.frx":059C
      TabIndex        =   1
      Top             =   600
      Width           =   2412
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2160
      Width           =   1455
   End
End
Attribute VB_Name = "FileSaveOKForm"
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

