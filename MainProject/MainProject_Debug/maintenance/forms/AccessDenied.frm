VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form AccessDenied 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Denied"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   5415
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   480
      OleObjectBlob   =   "AccessDenied.frx":0000
      TabIndex        =   1
      Top             =   840
      Width           =   4575
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "AccessDenied.frx":00BE
      Top             =   240
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "AccessDenied"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub
