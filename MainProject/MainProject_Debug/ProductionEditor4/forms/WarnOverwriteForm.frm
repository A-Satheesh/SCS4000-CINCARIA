VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form WarnOverwriteForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Warning!"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   600
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   855
      Left            =   600
      OleObjectBlob   =   "WarnOverwriteForm.frx":0000
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "WarnOverwriteForm.frx":00C8
      Top             =   840
   End
End
Attribute VB_Name = "WarnOverwriteForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim line As Integer
    
    
    Set A = fs.createtextfile(fileSaveForm.fileSavePath.Text, True)

    For line = 0 To editorForm.lstPattern.ListCount
        If editorForm.lstPattern.List(line) <> "" Then
            A.writeline (editorForm.lstPattern.List(line))
        End If
    Next line
    
    A.Close
        
    editorForm.Caption = fileSaveForm.fileSavePath.Text
    fileDirty = False
    Unload Me
    
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd

End Sub
