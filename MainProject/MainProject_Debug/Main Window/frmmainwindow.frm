VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmmainwindow 
   BorderStyle     =   0  'None
   Caption         =   "Main Window"
   ClientHeight    =   11505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15345
   Icon            =   "frmmainwindow.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmmainwindow.frx":08CA
   ScaleHeight     =   11505
   ScaleMode       =   0  'User
   ScaleWidth      =   15345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   3240
      Picture         =   "frmmainwindow.frx":44DE
      ScaleHeight     =   1278.635
      ScaleMode       =   0  'User
      ScaleWidth      =   10214.39
      TabIndex        =   9
      Top             =   10320
      Width           =   10425
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   360
      Picture         =   "frmmainwindow.frx":E2C6
      ScaleHeight     =   991.358
      ScaleMode       =   0  'User
      ScaleWidth      =   2535
      TabIndex        =   8
      Top             =   10255
      Width           =   2535
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   0
      Picture         =   "frmmainwindow.frx":FF60
      ScaleHeight     =   1278.635
      ScaleMode       =   0  'User
      ScaleWidth      =   15035
      TabIndex        =   6
      Top             =   10200
      Width           =   15345
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000D&
      Caption         =   "Production"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5520
      TabIndex        =   5
      Top             =   6600
      Width           =   4500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000D&
      Caption         =   "Programming"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5520
      TabIndex        =   4
      Top             =   4980
      Width           =   4500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Maintenance"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   5520
      TabIndex        =   3
      Top             =   3480
      Width           =   4500
   End
   Begin VB.Timer ShellWaitTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   5280
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   13920
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   12000
      TabIndex        =   1
      Text            =   "Production Run"
      Top             =   7080
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Shutdown"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   0
      Top             =   9600
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmmainwindow.frx":16E65
      Top             =   5760
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   10095
      Left            =   0
      Picture         =   "frmmainwindow.frx":17099
      ScaleHeight     =   10095
      ScaleWidth      =   15375
      TabIndex        =   7
      Top             =   0
      Width           =   15375
   End
End
Attribute VB_Name = "frmmainwindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim J As Integer
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Sub cmdExit_Click()
Shell "shutdown -s -f -t 00"
End Sub

Private Sub cmdok_Click()
If Combo1.Text = "Maintenance" Then
    cmdmaintenance_Click
ElseIf Combo1.Text = "Profile Editor" Then
    cmdprofileeditor_Click
Else
    cmdproductionrun_Click
End If
End Sub

Private Sub cmdmaintenance_Click()

cmdok.Enabled = False
cmdExit.Enabled = False
Combo1.Enabled = False
Dim line As String
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists("C:\Desktop Access\LogIn_Maintenance") Then
    Set A = fs.OpenTextFile("C:\Desktop Access\LogIn_Maintenance", 1, False)
        Do While A.AtEndOfStream <> True
            line = A.ReadLine
        Loop
            If (line = "") Then
                frmregisteration.Show vbModal
                If exitregister = True Then
                    cmdok.Enabled = True
                    cmdExit.Enabled = True
                    Combo1.Enabled = True
                Else
                    cmdmaintenance_Click
                End If
            Else
                Command1.Enabled = False      'ask
                Command2.Enabled = False
                Command3.Enabled = False
                lTaskID = Shell("C:\MainProject\maintenance\maintenance.exe", vbHide)
                ShellWaitTimer.Enabled = True
                ''Get process handle
                'lPID = OpenProcess(PROCESS_ALL_ACCESS, True, lTaskID)
                'If lPID Then
                '    'Wait for process to finish
                '    Call WaitForSingleObject(lPID, INFINITE)
                '    lTaskID = CloseHandle(lPID)
                'End If
            End If
Else
    frmregisteration.Show vbModal
    If exitregister = True Then
        cmdok.Enabled = True
        cmdExit.Enabled = True
        Combo1.Enabled = True
    Else
        cmdmaintenance_Click
    End If
End If

End Sub

Private Sub cmdprofileeditor_Click()
cmdok.Enabled = False
cmdExit.Enabled = False
Combo1.Enabled = False
Dim line As String
Set fs = CreateObject("Scripting.FileSystemObject")
If fs.FileExists("C:\Desktop Access\LogIn_Profile") Then
    Set A = fs.OpenTextFile("C:\Desktop Access\LogIn_Profile", 1, False)
        Do While A.AtEndOfStream <> True
            line = A.ReadLine
        Loop
            If (line = "") Then
                frmregisteration.Show vbModal
                If exitregister = True Then
                    cmdok.Enabled = True
                    cmdExit.Enabled = True
                    Combo1.Enabled = True
                Else
                    cmdprofileeditor_Click
                End If
            Else
                Command1.Enabled = False      'ask
                Command2.Enabled = False
                Command3.Enabled = False
                lTaskID = Shell("C:\MainProject\ProductionEditor4\ProfileEditor.exe", vbHide)
                ShellWaitTimer.Enabled = True
                ''Get process handle
                'lPID = OpenProcess(PROCESS_ALL_ACCESS, True, lTaskID)
                'If lPID Then
                '    'Wait for process to finish
                '    Call WaitForSingleObject(lPID, INFINITE)
                '    lTaskID = CloseHandle(lPID)
                'End If
            End If
        
Else
    frmregisteration.Show vbModal
    If exitregister = True Then
        cmdok.Enabled = True
        cmdExit.Enabled = True
        Combo1.Enabled = True
    Else
        cmdprofileeditor_Click
    End If
End If

End Sub

Private Sub cmdproductionrun_Click()
cmdok.Enabled = False
cmdExit.Enabled = False
Combo1.Enabled = False

Command1.Enabled = False      'ask
Command2.Enabled = False
Command3.Enabled = False

lTaskID = Shell("C:\MainProject\ProductionRunEngine4\ProductionRunEngine.exe", vbHide)
ShellWaitTimer.Enabled = True

    'one method
    'Get process handle
    'lPID = OpenProcess(PROCESS_ALL_ACCESS, True, lTaskID)
    'If lPID Then
    '    'Wait for process to finish
    '    Call WaitForSingleObject(lPID, INFINITE)
    '    lTaskID = CloseHandle(lPID)
    '    Combo1.Enabled = True
    '   cmdok.Enabled = True
    '    cmdexit.Enabled = True
    '    Me.Enabled = True
    'End If

End Sub

Private Sub Command1_Click()
    Combo1.Text = "Maintenance"
    cmdmaintenance_Click
End Sub

Private Sub Command2_Click()
    Combo1.Text = "Profile Editor"
    cmdprofileeditor_Click
End Sub

Private Sub Command3_Click()
    Combo1.Text = "Production Run"
    cmdproductionrun_Click
End Sub


Private Sub Form_load()
With Me
.Width = Screen.Width
.Height = Screen.Height
.Top = 0
.Left = 0
End With

Skin1.LoadSkin ("C:\MainProject\maintenance\skin\epoxySkin.skn")
Skin1.ApplySkin Me.hWnd

Combo1.AddItem "Maintenance"
Combo1.AddItem "Profile Editor"
Combo1.AddItem "Production Run"

End Sub


Private Sub ShellWaitTimer_Timer()
Dim hWnd1 As Long, hWnd2 As Long, hWnd3 As Long
Dim hWnd11 As Long, hWnd22 As Long

If Combo1.Text = "Maintenance" Then
    hWnd1 = FindWindow(vbNullString, "Desktop Setup Panel")
    hWnd11 = FindWindow(vbNullString, "Login to Maintenance")
    
    If hWnd1 = 0 And hWnd11 = 0 Then
        cmdok.Enabled = True
        cmdExit.Enabled = True
        Combo1.Enabled = True
        Command1.Enabled = True     'ask
        Command2.Enabled = True
        Command3.Enabled = True
        ShellWaitTimer.Enabled = False
    End If
ElseIf Combo1.Text = "Profile Editor" Then
    hWnd2 = FindWindow(vbNullString, "Profile Editor")
    hWnd22 = FindWindow(vbNullString, "Login to Profile Editor")
    If hWnd2 = 0 And hWnd22 = 0 Then
        cmdok.Enabled = True
        cmdExit.Enabled = True
        Combo1.Enabled = True
        Command1.Enabled = True        'ask
        Command2.Enabled = True
        Command3.Enabled = True
        ShellWaitTimer.Enabled = False
    End If
Else
    hWnd3 = FindWindow(vbNullString, "File Load")
    If hWnd3 = 0 Then
        cmdok.Enabled = True
        cmdExit.Enabled = True
        Combo1.Enabled = True
        Command1.Enabled = True         'ask
        Command2.Enabled = True
        Command3.Enabled = True
        ShellWaitTimer.Enabled = False
    End If
End If

End Sub

