VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form fileSaveForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Save As"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9345
   Icon            =   "fileSaveForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9345
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   600
      OleObjectBlob   =   "fileSaveForm.frx":08CA
      TabIndex        =   6
      Top             =   4560
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "fileSaveForm.frx":0940
      Top             =   5280
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.TextBox fileSavePath 
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Top             =   4560
      Width           =   6855
   End
   Begin VB.FileListBox File1 
      Height          =   3405
      Left            =   4560
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   3375
   End
End
Attribute VB_Name = "fileSaveForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    
    If Len(File1.Path) <> 3 Then
        fileSavePath.Text = File1.Path & "\"
    Else
        fileSavePath.Text = File1.Path
    End If
End Sub
Private Sub Drive1_Change()

    Dim tempstr1

    tempstr1 = Split(Drive1.Drive)

    If fs.folderexists(tempstr1(0)) Then
        Dir1.Path = Drive1.Drive
    End If
End Sub

Private Sub File1_Click()
    fileSavePath.Text = File1.Path & "\" & File1.FileName
End Sub

Private Sub Form_Load()
    fileSavePath.Text = File1.Path & "\"
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)

    Skin1.ApplySkin Me.hWnd

    Dim tempStr As String
    
    tempStr = GetStringSetting("EpoxyDispenser", "Setup", "defaultPatternDir")
    
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If fs.folderexists(tempStr) Then
        Dir1.Path = tempStr
    End If


End Sub

Private Sub ok_Click()

    If Right(fileSavePath.Text, 1) <> "\" Then

        Dim line As Integer

        Set fs = CreateObject("Scripting.FileSystemObject")
        
        If fs.FileExists(fileSavePath.Text) Then
            WarnOverwriteForm.Show (vbModal)
        Else
            Set A = fs.createtextfile(fileSavePath.Text, True)

            For line = 0 To editorForm.lstPattern.ListCount
                If editorForm.lstPattern.List(line) <> "" Then
                    A.writeline (editorForm.lstPattern.List(line))
                End If
            Next line
    
            A.Close
        
            editorForm.Caption = fileSavePath.Text
            fileDirty = False
        End If
    
    Else
        MsgBox ("Invalid file path!")
    End If
    
    Unload Me
End Sub

