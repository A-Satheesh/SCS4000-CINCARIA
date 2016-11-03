VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmGlobalTimeSpeed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Dispense Time/ Travel Speed"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel lblTravelSpeed 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "frmGlobalTimeSpeed.frx":0000
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblDispenseTime 
      Height          =   255
      Left            =   1320
      OleObjectBlob   =   "frmGlobalTimeSpeed.frx":0076
      TabIndex        =   4
      Top             =   600
      Width           =   1215
   End
   Begin VB.TextBox txtTravelSpeed 
      Height          =   375
      Left            =   2760
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtDispenseTime 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmGlobalTimeSpeed.frx":00EE
      Top             =   120
   End
End
Attribute VB_Name = "frmGlobalTimeSpeed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Time As Double, Speed As Integer            'To check whether the user change the default value or not.
Dim TimeFlag As Boolean, SpeedFlag As Boolean

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
    
    txtDispenseTime.TabIndex = 0
    txtTravelSpeed.TabIndex = 1
    Command1.TabIndex = 2
    Command2.TabIndex = 3
    
    'Do the initialization
    GlobalTravelSpeed = 0
End Sub

Private Sub Command1_Click()
    If (txtDispenseTime.Text = "") And (txtTravelSpeed.Text = "") Then
        FlagGlobalTimeSpeed = False
        Unload Me
        Exit Sub
    End If
    
    'This is doing the checking not to do anything if the default value of the text box is not changed.
    If (TimeFlag = True) Then
        TimeFlag = False
        GlobalDispenseTime = txtDispenseTime.Text
    End If
    
    If (SpeedFlag = True) Then
        SpeedFlag = False
        GlobalTravelSpeed = txtTravelSpeed.Text
    End If
    
    FlagGlobalTimeSpeed = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub txtDispenseTime_Validate(cancel As Boolean)
    If (txtDispenseTime.Text <> "") Then
        If Not IsNumeric(txtDispenseTime.Text) Then
            MsgBox "Please kye the numeric value."
            txtDispenseTime.Text = ""
            cancel = True
        Else
            If (CDbl(txtDispenseTime.Text) < 0) Then
                MsgBox "DispenseTime shouldn't be minus value."
                txtDispenseTime.Text = ""
                cancel = True
            Else
                TimeFlag = True
            End If
        End If
    End If
End Sub

Private Sub txtTravelSpeed_Validate(cancel As Boolean)
    If (txtTravelSpeed.Text <> "") Then
        If Not IsNumeric(txtTravelSpeed.Text) Then
            MsgBox "Please kye the numeric value."
            txtTravelSpeed.Text = ""
            cancel = True
        Else
            If (CInt(txtTravelSpeed.Text) <= 0) Then
                MsgBox "TravelSpeed shouldn't be less than zero."
                txtTravelSpeed.Text = ""
                cancel = True
            ElseIf (CInt(txtTravelSpeed.Text) > 500) Then
                MsgBox "TravelSpeed shouldn't be more than '500'."
                txtTravelSpeed.Text = ""
                cancel = True
            Else
                SpeedFlag = True
            End If
        End If
   End If
End Sub

