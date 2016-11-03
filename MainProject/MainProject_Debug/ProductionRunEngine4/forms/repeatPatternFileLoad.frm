VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form repeatPatternFileLoad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Load Pattern to repeat"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   9390
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   3375
   End
   Begin VB.FileListBox File1 
      Height          =   3990
      Left            =   4320
      TabIndex        =   2
      Top             =   600
      Width           =   4455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   5400
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   5400
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   720
      OleObjectBlob   =   "repeatPatternFileLoad.frx":0000
      Top             =   4320
   End
End
Attribute VB_Name = "repeatPatternFileLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub ok_Click()

    editorForm.PathFileName.Text = File1.Path & "\" & File1.FileName


    Unload Me
    
End Sub
Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub
