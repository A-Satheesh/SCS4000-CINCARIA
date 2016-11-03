VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form fileLoadForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Load"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8505
   Icon            =   "fileLoadForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1320
      OleObjectBlob   =   "fileLoadForm.frx":08CA
      Top             =   4680
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5520
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
Dim Not_selected As Boolean

Private Sub cancel_Click()
    File_Load_Cancel = True
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

Private Sub File1_Click()
    Not_selected = True
End Sub

Private Sub File1_DblClick()
    File_Load
End Sub

Private Sub Form_Load()
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

    'Dim words() As String
    'Dim line As String

    'If File1.FileName <> "" Then

    '    editorForm.lstPattern.Clear

    '    Set fs = CreateObject("Scripting.FileSystemObject")
    '    Set A = fs.OpenTextFile(File1.Path & "\" & File1.FileName, 1, False)

    '    Do While A.AtEndOfStream <> True
    '        line = A.ReadLine
    '        If (line <> "") Then
    '            words() = Split(line, "(")
    '
    '            If (line = "*** Left-Needle ***") Then
    '                Right_Needle_ON = False
    '            ElseIf (line = "*** Right-Needle ***") Then
    '                Right_Needle_ON = True
    '            End If
            
    '            If StrComp(words(0), "reference", vbTextCompare) = 0 Then
    '                'Read X,Y and Z reference
    '                Call GettingPosition(line)
    '                referenceX = ModifyOffsetX
    '                referenceY = ModifyOffsetY
    '                referenceZ = ModifyOffsetZ
    '                referenceSet = True
    '            End If
                
    '            editorForm.lstPattern.AddItem (line)
    '        End If
    '    Loop
    '    A.Close

    '    editorForm.Caption = File1.Path & "\" & File1.FileName

    '    'That is for not to do the first element whenever we load a new program     'XW
    '    'doTrack (editorForm.lstPattern.List(0) & vbNewLine)    'Origin
    '    'selectNodeIndex = -1                                   'Origin
    '    'editorForm.lstPattern.Selected(0) = True               'Origin
    '    'selectNodeIndex = 0                                    'Origin

    '    'Removed based on 280904 evaluation
    '    'If referenceSet = False Then
    '        'referenceWarning.Show (vbModal)
    '    'End If

    'End If
    
    'Unload Me
    If (Not_selected = True) Then
        File_Load
    Else
        MsgBox "Please select one program before pressing 'OK'."
    End If
End Sub

Private Sub File_Load()
    Dim words() As String
    Dim words2() As String, Dir_path As String  'To take the directory for part array
    Dim line As String
    Dim First_Line As Boolean               'Flage the 1st line


    If File1.FileName <> "" Then

        editorForm.lstPattern.Clear

        Set fs = CreateObject("Scripting.FileSystemObject")
        Set A = fs.OpenTextFile(File1.Path & "\" & File1.FileName, 1, False)

        Do While A.AtEndOfStream <> True
            line = A.ReadLine
            If (line <> "") Then
                words() = Split(line, "(")
            
                If (line = "*** Left-Needle ***") Then
                    Right_Needle_ON = False
                    First_Line = False
                ElseIf (line = "*** Right-Needle ***") Then
                    Right_Needle_ON = True
                    First_Line = False
                ElseIf (First_Line = False) Then
                    If StrComp(words(0), "fudicial", vbTextCompare) <> 0 Then
                        Call GettingPosition(line)
            
                        If StrComp(words(0), "reference", vbTextCompare) = 0 Then
                            'Read X,Y and Z reference
                            Call GettingPosition(line)
                            referenceX = ModifyOffsetX
                            referenceY = ModifyOffsetY
                            referenceZ = ModifyOffsetZ
                            referenceSet = True
                    
                        ElseIf StrComp(words(0), "repeat", vbTextCompare) = 0 Then
                            words2() = Split(line, ";")
                            Dir_path = Trim(words2(1))
                            Dir_path = Left(Dir_path, Len(Dir_path) - 2)
                            Dir_path = Right(Dir_path, Len(Dir_path) - 1)
                        
                        End If
                    
                        If (Right_Needle_ON = False) Then
                            Z_High = ModifyOffsetZ
                            reference_ZHigh = True
                        Else
                            R_Z_High = ModifyOffsetZ
                            reference_R_ZHigh = True
                        End If
                            
                        First_Line = True
                    
                    End If
                    
                    
                Else
                    If StrComp(words(0), "reference", vbTextCompare) = 0 Then
                        'Read X,Y and Z reference
                        Call GettingPosition(line)
                        referenceX = ModifyOffsetX
                        referenceY = ModifyOffsetY
                        referenceZ = ModifyOffsetZ
                        referenceSet = True
                        
                    ElseIf StrComp(words(0), "repeat", vbTextCompare) = 0 Then
                        words2() = Split(line, ";")
                        Dir_path = Trim(words2(1))
                        Dir_path = Left(Dir_path, Len(Dir_path) - 2)
                        Dir_path = Right(Dir_path, Len(Dir_path) - 1)
                    End If
                End If
                editorForm.lstPattern.AddItem (line)
            End If
        Loop
        A.Close

        editorForm.Caption = File1.Path & "\" & File1.FileName
        
        If (Dir_path <> "") Then
            editorForm.PathFileName.Text = Dir_path
        End If
        'That is for not to do the first element whenever we load a new program     'XW
        'doTrack (editorForm.lstPattern.List(0) & vbNewLine)    'Origin
        'selectNodeIndex = -1                                   'Origin
        'editorForm.lstPattern.Selected(0) = True               'Origin
        'selectNodeIndex = 0                                    'Origin

        'Removed based on 280904 evaluation
        'If referenceSet = False Then
            'referenceWarning.Show (vbModal)
        'End If

    End If
    
    Unload Me
End Sub

'This procedure will take the robort's position, (X, Y and Z) from the whold string (XW)
Private Sub GettingPosition(ByVal txtString As String)
    Dim n, m, i As Integer
    Dim step, number, Count As Long
    Dim char, check1, check2 As String
            
    char = ""
    check1 = ""
    check2 = ""
    n = 0
    m = 0
    ModifyOffsetX = 0
    ModifyOffsetY = 0
    ModifyOffsetZ = 0
    One = False
    step = 1
    number = Len(txtString)
            
     For Count = 1 To number
        char = char & Mid(txtString, step, 1)
        check1 = Right(char, 2)
        check2 = Right(char, 1)
        If (check1 = "x=") Or (check1 = "y=") Or (check1 = "z=") Then
            n = step
            step = step + 1
        ElseIf check2 = "," Or check2 = ")" Or (check2 = ";") Then
            m = step
            step = step + 1
            
            If (Mid(char, n - 1, 2) = "x=") Then
                ModifyOffsetX = Val(Right(char, m - n))
            ElseIf (Mid(char, n - 1, 2) = "y=") Then
                ModifyOffsetY = Val(Right(char, m - n))
            ElseIf (Mid(char, n - 1, 2) = "z=") Then
                ModifyOffsetZ = Val(Right(char, m - n))
                i = step
            End If
            
            If (check2 = ";") Or (check2 = ")") Then
                Exit Sub
            End If
        Else
            step = step + 1
        End If
    Next Count
End Sub
