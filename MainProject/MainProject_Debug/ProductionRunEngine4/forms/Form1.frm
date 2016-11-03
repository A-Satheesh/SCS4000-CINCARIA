VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form frmReset 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ResetForm(Reset.frm)"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Quit"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   360
      OleObjectBlob   =   "Form1.frx":0000
      Top             =   1920
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   495
      Left            =   600
      OleObjectBlob   =   "Form1.frx":0234
      TabIndex        =   1
      Top             =   720
      Width           =   3135
   End
End
Attribute VB_Name = "frmReset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim resetValue, delay As Long
    Unload Me
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, resetValue))
    
    resetValue = resetValue Or &H200
    
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
    
    Call Sleep(0.5)
        
    resetValue = resetValue And &HFDFF
        
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
  
    Call Sleep(1)
    
    Call moveToHome
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionRunEngine4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Sub Form_Unload(cancel As Integer)
    Unload Me
End Sub


Private Sub Form_Activate()
    SetWindowOnTop Me, True    '@$K
End Sub

Private Sub Form_Deactivate()
    SetWindowOnTop Me, False    '@$K
End Sub

