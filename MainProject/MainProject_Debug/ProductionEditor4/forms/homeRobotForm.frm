VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form homeRobotForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Please home robot!"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "homeRobotForm.frx":0000
      Top             =   240
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   1095
      Left            =   1200
      OleObjectBlob   =   "homeRobotForm.frx":0234
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "homeRobotForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub
