VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmGlobalOffset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Global Offset"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel lblOffsetZ 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "frmGlobalOffset.frx":0000
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblOffsetY 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "frmGlobalOffset.frx":007C
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel lblOffsetX 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "frmGlobalOffset.frx":00F8
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtZ 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Text            =   "0"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtY 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txtX 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Text            =   "0"
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2760
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmGlobalOffset.frx":0174
      Top             =   120
   End
End
Attribute VB_Name = "frmGlobalOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Command1_Click()
    GlobalOffsetX = txtX.Text
    GlobalOffsetY = txtY.Text
    GlobalOffsetZ = txtZ.Text
    FlagGlobalOffset = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub txtX_Validate(cancel As Boolean)
    If (txtX.Text <> "") Then
        If Not IsNumeric(txtX.Text) Then
            MsgBox "Please kye the numeric value."
            txtX.Text = ""
            cancel = True
        Else
            OffSetX = txtX.Text
        End If
    Else
        MsgBox "X_Offset value could not be blank."
        cancel = True
    End If
End Sub

Private Sub txtY_Validate(cancel As Boolean)
    If (txtY.Text <> "") Then
        If Not IsNumeric(txtY.Text) Then
            MsgBox "Please kye the numeric value."
            txtY.Text = ""
            cancel = True
        Else
            OffSetY = txtY.Text
        End If
    Else
        MsgBox "Y_Offset value could not be blank."
        cancel = True
    End If
End Sub

Private Sub txtZ_Validate(cancel As Boolean)
    If (txtZ.Text <> "") Then
        If Not IsNumeric(txtZ.Text) Then
            MsgBox "Please kye the numeric value."
            txtZ.Text = ""
            cancel = True
        Else
            OffSetZ = txtZ.Text
        End If
    Else
        MsgBox "Z_Offset value could not be blank."
        cancel = True
    End If
End Sub

