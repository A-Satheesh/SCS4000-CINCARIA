Attribute VB_Name = "P1750"
Global Const MaxDev = 255                   ' max.devices
Global Const MaxPort = 3
Global Const REMOTE = 1
Global Const NONPROG = 0
Global Const PROG = REMOTE
Global Const INTERNAL = 0
Global Const EXTERNAL = 1
Global Const AAC = &H0                      'Define board vendor ID
Global Const NONE = &H0                     'Define DAS I/O CardType ID.
Global Const BD_PCI1750 = AAC Or &H5E       'Advantech CardType ID
'Global Const SUCCESS = 0                    'Define Status Code

Global Const CFG_DeviceNumber = &H0
Global Const CFG_BoardID = &H1
Global Const CFG_SwitchID = &H2
Global Const CFG_BaseAddress = &H3
Global Const CFG_SlotNumber = &H6

Global Const CFG_DiPortCount = &H3C84           'Get DI Port Count. Max available DI Port count on the card.
Global Const CFG_DoPortCount = &H4808           'Get Do Port Count. Max available DO Port count on the card.
Global Const DIO_ChannelDir_DI = &H0            'Get/Set DIO Channel Direction ( IN / OUT ).
Global Const DIO_ChannelDir_DO = &HFF
Global Const DI_DataWidth_Byte = &H0            'Get DI data width. The optimized data width when Reading.
Global Const DI_DataWidth_Word = &H1
Global Const DI_DataWidth_DWORD = &H2
Global Const DO_DataWidth_Byte = &H0            'Get DO data width. The optimized data width when writing.
Global Const DO_DataWidth_Word = &H1
Global Const DO_DataWidth_Dword = &H2

Type PT_DEVLIST
    dwDeviceNum  As Long
    szDeviceName(0 To 49) As Byte
    nNumOfSubdevices As Integer
End Type

'**************************************************************************
'    Function Declaration for ADSAPI32
'**************************************************************************
Declare Function DRV_SelectDevice Lib "adsapi32.dll" (ByVal hCaller As Long, ByVal GetModule As Boolean, DeviceNum As Long, ByVal Description As String) As Long
Declare Function DRV_DeviceGetNumOfList Lib "adsapi32.dll" (NumOfDevices As Integer) As Long
Declare Function DRV_DeviceGetList Lib "adsapi32.dll" (ByVal devicelist As Long, ByVal MaxEntries As Integer, nOutEntries As Integer) As Long
Declare Function DRV_DeviceOpen Lib "adsapi32.dll" (ByVal DeviceNum As Long, DriverHandle As Long) As Long
Declare Function DRV_DeviceClose Lib "adsapi32.dll" (DriverHandle As Long) As Long
Declare Function DRV_DeviceSetProperty Lib "adsapi32.dll" (ByVal DriverHandle As Long, ByVal nID As Integer, ByRef pBuffer As Any, ByVal dwLength As Long) As Long
Declare Function DRV_DeviceGetProperty Lib "adsapi32.dll" (ByVal DriverHandle As Long, ByVal nID As Integer, ByRef pBuffer As Any, ByRef pLength As Long) As Long
Declare Sub DRV_GetErrorMessage Lib "adsapi32.dll" (ByVal lError As Long, ByVal lpszszErrMsg As String)
Declare Function DRV_GetAddress Lib "adsapi32.dll" (lpVoid As Any) As Long
Declare Function AdxDioReadDiPorts Lib "adsapi32.dll" (ByVal DriverHandle As Long, ByVal dwPortStart As Long, ByVal dwPortCount As Long, ByRef pBuffer As Byte) As Long
Declare Function AdxDioWriteDoPorts Lib "adsapi32.dll" (ByVal DriverHandle As Long, ByVal dwPortStart As Long, ByVal dwPortCount As Long, ByRef pBuffer As Byte) As Long
Declare Function AdxDioGetCurrentDoPortsState Lib "adsapi32.dll" (ByVal DriverHandle As Long, ByVal dwPortStart As Long, ByVal dwPortCount As Long, ByRef pBuffer As Byte) As Long



