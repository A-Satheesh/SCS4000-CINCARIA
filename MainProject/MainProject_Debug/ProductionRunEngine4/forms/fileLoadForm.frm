VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form fileLoadForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Load"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "fileLoadForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8520
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "fileLoadForm.frx":08CA
      Top             =   4800
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5640
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Height          =   4185
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      Height          =   3690
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   3375
   End
End
Attribute VB_Name = "fileLoadForm"
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
    Dim tempstr1

    tempstr1 = Split(Drive1.Drive)

    If fs.folderexists(tempstr1(0)) Then
        Dir1.Path = Drive1.Drive
    End If
End Sub

Private Sub Form_Load()

    Dim hWnd, hWnd2 As Long

    hWnd = FindWindow(vbNullString, "Desktop Setup Panel")
    hWnd2 = FindWindow(vbNullString, "Profile Editor")

    If App.PrevInstance Or hWnd <> 0 Or hWnd2 <> 0 Then
        MsgBox ("Another conflicting process has been detected! This process will abort")
        Unload Me
    Else
        'Skin1.LoadSkin (".\skin\epoxySkin.skn")
        Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
        Skin1.ApplySkin Me.hWnd
        
        Dim tempStr As String
    
        tempStr = GetStringSetting("EpoxyDispenser", "Setup", "defaultPatternDir")
        
        'The offset value between upper camera and left_needle
        Offset_DistanceX_Camera_L_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_L_Needle", "0")
        Offset_DistanceY_Camera_L_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_L_Needle", "0")
            
        'The offset value between upper camera and right_needle
        Offset_DistanceX_Camera_R_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistX_Camera_R_Needle", "0")
        Offset_DistanceY_Camera_R_Needle = GetStringSetting("EpoxyDispenser", "NeedleOffset", "Off_DistY_Camera_R_Needle", "0")
            
        'Z offset valve for both left and right needle
        needleOffsetZ_L = GetStringSetting("EpoxyDispenser", "NeedleOffset", "needleOffsetZ_L", "0")
        needleOffsetZ_R = GetStringSetting("EpoxyDispenser", "NeedleOffset", "needleOffsetZ_R", "0")
    
        '@$K
        solventposX = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "xSolventPos", "0")), Z_axis)
        solventposY = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "ySolventPos", "0")), Z_axis)
        solventposZ = convertToPulses(CDbl(GetStringSetting("EpoxyDispenser", "Setup", "zSolventPos", "0")), Z_axis)
        
        Set fs = CreateObject("Scripting.FileSystemObject")
        
        If fs.folderexists(tempStr) Then
            Dir1.Path = tempStr
        End If
        
        SetWindowOnTop Me, True    '@$K
    End If
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    SetWindowOnTop Me, False    '@$K
End Sub

  Private Sub ok_Click()
    ok.Enabled = False
    cancel.Enabled = False
    CloseBoard = False
    CaptionRunEngine = ""
    
    If File1.FileName <> "" Then
        File1size = FileLen(File1.Path & "\" & File1.FileName)
        
        If File1size <> 0 Then
    
            editorForm.lstPattern.Clear
        
            'Origin
            'executionForm.Caption = "Production " & File1.Path & "\" & File1.FileName
        
            'XW
            CaptionRunEngine = "Production " & File1.Path & "\" & File1.FileName
        
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set A = fs.OpenTextFile(File1.Path & "\" & File1.FileName, 1, False)

            Do While A.AtEndOfStream <> True
                editorForm.lstPattern.AddItem (A.ReadLine)
                'Total Line in listbox (XW)
                TotalLine = TotalLine + 1
            Loop
            A.Close

            editorForm.Caption = File1.Path & "\" & File1.FileName

            'Run Engine
            'doTrack (editorForm.lstPattern.List(0) & vbNewLine)
            'selectNodeIndex = -1
    
            'editorForm.lstPattern.Selected(0) = True
            'selectNodeIndex = 0

            'Removed based on 280904 evaluation
            'If referenceSet = False Then
                'referenceWarning.Show (vbModal)
            'End If

            'XW
            leftside = False
            rightside = False
            editorForm.displayCoOrdsTimer.Enabled = False
            'Run Engine Additions
            translateForm.startTranslate
        
            editorForm.displayCoOrdsTimer.Enabled = True
        
        Else
            MsgBox ("There is no program to run the machine!")
        End If
    End If
    'Run Engine Additions
    Unload executionForm
    Unload editorForm
    ok.Enabled = True
    cancel.Enabled = True
End Sub

