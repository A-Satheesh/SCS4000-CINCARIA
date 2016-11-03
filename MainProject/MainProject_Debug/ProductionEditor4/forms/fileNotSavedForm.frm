VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form fileNotSavedForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File is not saved!"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "fileNotSavedForm.frx":0000
      Top             =   120
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   975
      Left            =   840
      OleObjectBlob   =   "fileNotSavedForm.frx":0234
      TabIndex        =   2
      Top             =   600
      Width           =   4095
   End
   Begin VB.CommandButton OKCommand 
      Caption         =   "Yes"
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton CancelCommand 
      Caption         =   "No"
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "fileNotSavedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelCommand_Click()
    proceedWithAction = False
    Unload Me
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd

End Sub

Private Sub OKCommand_Click()
    proceedWithAction = True
    Unload Me
End Sub
