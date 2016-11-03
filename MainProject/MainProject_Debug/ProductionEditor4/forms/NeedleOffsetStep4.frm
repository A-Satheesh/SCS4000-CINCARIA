VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form NeedleOffsetStep4 
   Caption         =   "Determine Needle Offset Step 4 of 4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   375
      Left            =   480
      OleObjectBlob   =   "NeedleOffsetStep4.frx":0000
      TabIndex        =   3
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtLightIntensity 
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Determine Offset"
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
   Begin MSComctlLib.Slider SliderLightIntensity 
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      _Version        =   393216
      Max             =   255
      TickFrequency   =   25
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "NeedleOffsetStep4.frx":007C
      Top             =   0
   End
End
Attribute VB_Name = "NeedleOffsetStep4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim offsetX, offsetY As Double
    
    returncode = VdeFindNeedleOffset(offsetX, offsetY)
    
    If returncode = 1 Then
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "XOff", offsetX
        SaveStringSetting "EpoxyDispenser", "NeedleOffset", "YOff", offsetY
    Else
        ErrorOffset.Show (vbModal)
    End If
    Unload Me

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (".\skin\epoxySkin.skn")

    Skin1.ApplySkin Me.hWnd
    
    
    VdeSelectCamera 1
    VdeReadSettings ("VisionSetup.txt")
    SliderLightIntensity = VdeGetLightIntensity()
    txtLightIntensity = SliderLightIntensity
End Sub

Private Sub Form_Unload(cancel As Integer)
    VdeSelectCamera 2
End Sub

Private Sub SliderLightIntensity_Scroll()
    VdeSetLightIntensity SliderLightIntensity
    txtLightIntensity = SliderLightIntensity
End Sub
