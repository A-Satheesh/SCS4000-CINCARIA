VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmOffset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Individual Offset"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtOffsetY 
      Height          =   405
      Left            =   2400
      TabIndex        =   4
      Text            =   "0"
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtOffsetZ 
      Height          =   405
      Left            =   2400
      TabIndex        =   3
      Text            =   "0"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox txtOffsetX 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   345
      Left            =   2400
      TabIndex        =   2
      Text            =   "0"
      Top             =   390
      Width           =   1335
   End
   Begin VB.CommandButton cmdOffset 
      Caption         =   "&Apply"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetX 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frmOffset.frx":0000
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetY 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frmOffset.frx":0078
      TabIndex        =   5
      Top             =   960
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.SkinLabel OffSetZ 
      Height          =   255
      Left            =   840
      OleObjectBlob   =   "frmOffset.frx":00F0
      TabIndex        =   6
      Top             =   1440
      Width           =   1455
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "frmOffset.frx":0168
      Top             =   2520
   End
End
Attribute VB_Name = "frmOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOffset_Click()
    If (txtOffsetX <> "0") Or (txtOffsetY <> "0") Or (txtOffsetZ <> "0") Then
        GlobalOffsetX = txtOffsetX.Text
        GlobalOffsetY = txtOffsetY.Text
        GlobalOffsetZ = txtOffsetZ.Text
        FlagGlobalOffset = True
    Else
        FlagGlobalOffset = False
    End If
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

Private Sub txtOffsetX_Validate(cancel As Boolean)
    If Not IsNumeric(txtOffsetX.Text) Then
        MsgBox "Please kye the numeric value."
        txtOffsetX.Text = "0"
        cancel = True
    End If
End Sub

Private Sub txtOffsetY_Validate(cancel As Boolean)
    If Not IsNumeric(txtOffsetX.Text) Then
        MsgBox "Please kye the numeric value."
        txtOffsetY.Text = "0"
        cancel = True
    End If
End Sub

Private Sub txtOffsetZ_Validate(cancel As Boolean)
    If Not IsNumeric(txtOffsetX.Text) Then
        MsgBox "Please kye the numeric value."
        txtOffsetZ.Text = "0"
        cancel = True
    End If
End Sub
