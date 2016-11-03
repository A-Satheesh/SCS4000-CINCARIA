Attribute VB_Name = "VisionModule"
Option Explicit

'change for login system (nno)
Public Declare Function VdeInitializeVision Lib "EpoxyVision2.dll" Alias "?VdeInitializeVision@@YGHPAUHWND__@@HHH@Z" _
        (ByVal hWnd As Long, ByVal sizeX As Integer, ByVal sizeY As Integer, ByVal Setupfiledir As Integer) As Integer
'Public Declare Function VdeInitializeVision Lib "EpoxyVision2.dll" Alias "?VdeInitializeVision@@YGHPAUHWND__@@HH@Z" _
'        (ByVal hWnd As Long, ByVal sizeX As Integer, ByVal sizeY As Integer) As Integer
       
Public Declare Function VdeInitializeVision1 Lib "EpoxyVision2.dll" Alias "?VdeInitializeVision1@@YGHPAUHWND__@@@Z" _
        (ByVal hWnd As Long) As Integer
        
Public Declare Sub VdeReleaseVision Lib "EpoxyVision2.dll" Alias "?VdeReleaseVision@@YGXXZ" _
        ()
        
Public Declare Sub VdeGetVisionMsg Lib "EpoxyVision2.dll" Alias "?VdeGetVisionMsg@@YGXPAD@Z" _
        (ByRef msg As String)

Public Declare Sub VdeSelectCamera Lib "EpoxyVision2.dll" Alias "?VdeSelectCamera@@YGXH@Z" _
        (ByVal iCamera As Integer)
        
Public Declare Sub VdeCameraLive Lib "EpoxyVision2.dll" Alias "?VdeCameraLive@@YGXH@Z" _
        (ByVal bLive As Integer)
        
Public Declare Sub VdeCameraSnap Lib "EpoxyVision2.dll" Alias "?VdeCameraSnap@@YGXXZ" _
        ()
        
Public Declare Function VdeIsGrabbing Lib "EpoxyVision2.dll" Alias "?VdeIsGrabbing@@YGHXZ" _
        (ByVal bLive As Integer) As Integer

Public Declare Sub VdeShowOverlay Lib "EpoxyVision2.dll" Alias "?VdeShowOverlay@@YGXH@Z" _
        (ByVal bLive As Integer)

Public Declare Function VdeIsShowOverlay Lib "EpoxyVision2.dll" Alias "?VdeIsShowOverlay@@YGHXZ" _
        () As Integer

Public Declare Sub VdeSetCameraSetting1 Lib "EpoxyVision2.dll" Alias "?VdeSetCameraSetting1@@YGXH@Z" _
        (ByVal Value As Integer)
        
Public Declare Sub VdeSetCameraSetting2 Lib "EpoxyVision2.dll" Alias "?VdeSetCameraSetting2@@YGXH@Z" _
        (ByVal Value As Integer)
        
Public Declare Sub VdeSetLightIntensity Lib "EpoxyVision2.dll" Alias "?VdeSetLightIntensity@@YGXH@Z" _
        (ByVal intensity As Integer)
        
Public Declare Function VdeGetLightIntensity Lib "EpoxyVision2.dll" Alias "?VdeGetLightIntensity@@YGHXZ" _
       () As Integer

Public Declare Function VdeOnLButtonDown Lib "EpoxyVision2.dll" Alias "?VdeOnLButtonDown@@YGHIJJ@Z" _
       (ByVal nFlags As Long, ByVal x As Long, ByVal y As Long) As Integer
       
Public Declare Function VdeOnMouseMove Lib "EpoxyVision2.dll" Alias "?VdeOnMouseMove@@YGHIJJ@Z" _
       (ByVal nFlags As Long, ByVal x As Long, ByVal y As Long) As Integer
       
Public Declare Function VdeOnLButtonUp Lib "EpoxyVision2.dll" Alias "?VdeOnLButtonUp@@YGHIJJ@Z" _
       (ByVal nFlags As Long, ByVal x As Long, ByVal y As Long) As Integer
       
Public Declare Function VdeScreenToActualPos Lib "EpoxyVision2.dll" Alias "?VdeScreenToActualPos@@YGHNNAAN0@Z" _
       (ByVal ptX As Double, ByVal ptY As Double, ByRef posX As Double, ByRef posY As Double) As Integer
       
Public Declare Function AVdectualToScreenPos Lib "EpoxyVision2.dll" Alias "?VdeActualToScreenPos@@YGHNNAAH0@Z" _
       (ByVal posX As Double, ByVal posY As Double, ByRef ptX As Double, ByRef ptY As Double) As Integer
       
Public Declare Function VdeGetCurrVisionMode Lib "EpoxyVision2.dll" Alias "?VdeGetCurrVisionMode@@YGHXZ" _
       () As Integer

Public Declare Function VdeCalibrationDlg Lib "EpoxyVision2.dll" Alias "?VdeCalibrationDlg@@YGHW4dlgMode@@PAD@Z" _
       (ByVal dlgMode As Integer, ByVal msg As String) As Integer
       
Public Declare Function VdeFindNeedleOffset Lib "EpoxyVision2.dll" Alias "?VdeFindNeedleOffset@@YGHAAN0@Z" _
       (ByRef OffSetX As Double, ByRef OffSetY As Double) As Integer

Public Declare Function VdeTeachRefPtDlg Lib "EpoxyVision2.dll" Alias "?VdeTeachRefPtDlg@@YGHW4dlgMode@@PAD@Z" _
       (ByVal dlgMode As Integer, ByVal msg As String) As Integer

Public Declare Sub VdeSetRefPtFilename Lib "EpoxyVision2.dll" Alias "?VdeSetRefPtFilename@@YGXPAD@Z" _
       (ByVal msg As String)

Public Declare Function VdeGetRefPtPos Lib "EpoxyVision2.dll" Alias "?VdeGetRefPtPos@@YGHAAN000@Z" _
       (ByRef x1 As Double, ByRef y1 As Double, ByRef x2 As Double, ByRef y2 As Double) As Integer

Public Declare Function VdeFindRefPt Lib "EpoxyVision2.dll" Alias "?VdeFindRefPt@@YGHPADNNNNAAN11@Z" _
       (ByVal msg As String, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
       ByRef dx As Double, ByRef dy As Double, ByRef da As Double) As Integer

Public Declare Function VdeReadSettings Lib "EpoxyVision2.dll" Alias "?VdeReadSettings@@YGHPAD@Z" _
       (ByVal msg As String) As Integer
       
Public Declare Function VdeWriteSettings Lib "EpoxyVision2.dll" Alias "?VdeWriteSettings@@YGHPAD@Z" _
       (ByVal msg As String) As Integer

Public Const VisionDlgInit = 0
Public Const VisionDlgNext = 1
Public Const VisionDlgBack = -1
Public Const VisionDlgCancel = -2
Public Const VisionDlgToFinish = -3
Public Const VisionDlgFinish = -4
Public Const VisionDlgOnTimer = -5

Public Declare Function VdeFindNeedleOffset1 Lib "EpoxyVision2.dll" Alias "?VdeFindNeedleOffset1@@YGHAAN0@Z" _
       (ByRef OffSetX As Double, ByRef OffSetY As Double) As Integer
Public Declare Function VdeFindCameraOffset1 Lib "EpoxyVision2.dll" Alias "?VdeFindCameraOffset1@@YGHAAN0@Z" _
       (ByRef OffSetX As Double, ByRef OffSetY As Double) As Integer




