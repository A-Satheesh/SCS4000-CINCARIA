Attribute VB_Name = "P1240"
Option Explicit
'****************************************************************************
'    Constant Definition
'****************************************************************************
Global Const MaxBoard = 16  'Maximum boards support in PCI-1240 utility
Global Const MaxAxis = 4    'Maximum axises support in one board
Global Const IPO_L2 = &H30
Global Const IPO_L3 = &H31
Global Const IPO_CW = &H32
Global Const IPO_CCW = &H134

'****************************************************************************
'    Constant definition for
'        P1240MotRdReg, P1240MotRdMutiReg, P1240MotWrReg, P1240MotwrMutiReg
'****************************************************************************
Global Const Rcnt = &H100   'Real position counter
Global Const Lcnt = &H101   'Logical position counter
Global Const Pcmp = &H102   'P direction compare register
Global Const Ncmp = &H103   'N direction compare register
Global Const Pnum = &H104   'Pulse Number
Global Const CurV = &H105   'Current V value
Global Const CurAC = &H106  'Current AC value

Global Const RR0 = &H200    'PCI-1240 control register 0 for read
Global Const RR1 = &H202    'PCI-1240 control register 1 for read
Global Const RR2 = &H204    'PCI-1240 control register 2 for read
Global Const RR3 = &H206    'PCI-1240 control register 3 for read
Global Const RR4 = &H208    'PCI-1240 control register 4 for read
Global Const RR5 = &H20A    'PCI-1240 control register 5 for read
Global Const RR6 = &H20C    'PCI-1240 control register 6 for read
Global Const RR7 = &H20E    'PCI-1240 control register 7 for read

Global Const WR0 = &H210    'PCI-1240 control register 0 for write
Global Const WR1 = &H212    'PCI-1240 control register 0 for write
Global Const WR2 = &H214    'PCI-1240 control register 0 for write
Global Const WR3 = &H216    'PCI-1240 control register 0 for write
Global Const WR4 = &H218    'PCI-1240 control register 0 for write
Global Const WR5 = &H21A    'PCI-1240 control register 0 for write
Global Const WR6 = &H21C    'PCI-1240 control register 0 for write
Global Const WR7 = &H21E    'PCI-1240 control register 0 for write

Global Const RG = &H300     'Internal range code
Global Const SV = &H301     'Initial speed
Global Const DV = &H302     'Const drive speed
Global Const MDV = &H303    'Maxium drive speed
Global Const AC = &H304     'Acceleration speed
Global Const DC = &H305     'Deacceleration speed (same as acceleration)
Global Const AK = &H306     'Acceleration rate
Global Const PLmt = &H307   'Maxium length for P or CW or + direction
Global Const NLmt = &H308   'Maxium length for N or CCW or - direction
Global Const HomeOffset = &H309 'Home offset from hardware home to logical home
Global Const HomeMode = &H30A   'Setting vaule for return home direction

Global Const HomeType = &H30B
Global Const HomeP0Dir = &H30C
Global Const HomeP0Speed = &H30D
Global Const HomeP1Dir = &H30E
Global Const HomeP1Speed = &H30F
Global Const HomeP2Dir = &H310
Global Const HomeOffsetSpeed = &H311

'***************************************************************************
'    Constant define for Operation Axis
'***************************************************************************
Global Const X_axis = &H1
Global Const Y_axis = &H2
Global Const Z_axis = &H4
Global Const U_axis = &H8
Global Const XY_axis = &H3
Global Const XZ_axis = &H5
Global Const XU_axis = &H9
Global Const YZ_axis = &H6
Global Const YU_axis = &HA
Global Const ZU_axis = &HC
Global Const XYZ_axis = &H7
Global Const XYU_axis = &HB
Global Const YZU_axis = &HE
Global Const XYZU_axis = &HF
'**************************************************************************
'    Error Message define
'**************************************************************************
Global Const Success = 0
Global Const BoardNumErr = &H1
Global Const CreateVxdFail = &H2
Global Const CallVxdFail = &H3
Global Const RegistryOpenFail = &H4
Global Const RegistryReadFail = &H5
Global Const AxisNumErr = &H6
Global Const UnderRGErr = &H7
Global Const OverRGErr = &H8
Global Const UnderSVErr = &H9
Global Const OverSVErr = &HA
Global Const OverMDVErr = &HB
Global Const UnderDVErr = &HC
Global Const OverDVErr = &HD
Global Const UnderACErr = &HE
Global Const OverACErr = &HF
Global Const UnderAKErr = &H10
Global Const OverAKErr = &H11
Global Const OverPLmtErr = &H12
Global Const OverNLmtErr = &H13
Global Const MaxMoveDistErr = &H14
Global Const AxisDrvBusy = &H15
Global Const RegItemErr = &H16
Global Const ParaValueErr = &H17
Global Const ParaValueOverRange = &H18
Global Const ParaValueUnderRange = &H19
Global Const AxisHomeBusy = &H1A
Global Const AxisExtBusy = &H1B
Global Const RegistryWriteFail = &H1C
Global Const ParaValueOverErr = &H1D
Global Const ParaValueUnderErr = &H1E
Global Const OverDCErr = &H1F
Global Const UnderDCErr = &H20
Global Const UnderMDVErr = &H21
Global Const RegistryCreateFail = &H22
Global Const CreateThreadErr = &H23
Global Const HomeSwStop = &H24
Global Const ChangeSpeedErr = &H25
Global Const DOPortAsDriverStatus = &H26

Global Const OpenEventFail = &H30
Global Const DeviceCloseErr = &H32

Global Const HomeEMGStop = &H40
Global Const HomeLMTPStop = &H41
Global Const HomeLMTNStop = &H42
Global Const HomeALARMStop = &H43

Global Const AllocateBufferFail = &H50
Global Const BufferReAllocate = &H51
Global Const FreeBufferFail = &H52
Global Const FirstPointNumberFail = &H53
Global Const PointNumExceedAllocatedSize = &H54
Global Const BufferNoneAllocate = &H55
Global Const SequenceNumberErr = &H56
Global Const PathTypeErr = &H57
Global Const PathTypeMixErr = &H60
Global Const BufferDataNotEnough = &H61


'**************************************************************************
'    Defnie Continue Motion data struct
'**************************************************************************
Type ContiPathData
    EndPoint_1 As Long  'End position for 1'st axis
    EndPoint_2 As Long  'End position for 2'nd axis
    EndPoint_3 As Long  'End position for 3'rd axis
    CenPoint_1 As Long  'Center position for 1'st axis
    CenPoint_2 As Long  'Center position for 2'rd axis
    PointNum As Long    'Serial number for current data
    PathType As Integer 'L2,L3,CWarc,CCWarc
    
    TempB As Integer    'For internal using
    TempA As Long       'For internal using
End Type

'**************************************************************************
'    DLL Function Declaration for PCI-1240
'**************************************************************************
Declare Function P1240MotDevAvailable Lib "ads1240.dll" (ByRef BoardAvailable As Long) As Long
Declare Function P1240MotDevOpen Lib "ads1240.dll" (ByVal Bid As Byte) As Long
Declare Function P1240MotDevClose Lib "ads1240.dll" (ByVal Bid As Byte) As Long
Declare Function P1240MotAxisParaSet Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal TS As Byte, ByVal ulSV As Long, ByVal ulDV As Long, ByVal ulMDV As Long, ByVal ulAC As Long, ByVal ulAK As Long) As Long
Declare Function P1240MotCmove Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal Dir As Byte) As Long
Declare Function P1240MotHome Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte) As Long
Declare Function P1240MotChgDV Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal Spd As Long) As Long
Declare Function P1240MotPtp Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal RA As Byte, ByVal PulseX As Long, ByVal PulseY As Long, ByVal PulseZ As Long, ByVal PulseU As Long) As Long
Declare Function P1240MotLine Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal RA As Byte, ByVal posX As Long, ByVal posY As Long, ByVal PosZ As Long, ByVal PosU As Long) As Long
Declare Function P1240MotArc Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal RA As Byte, ByVal Dir As Byte, ByVal Cen1 As Long, ByVal Cen2 As Long, ByVal End1 As Long, ByVal End2 As Long) As Long
Declare Function P1240MotArcTheta Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal RA As Byte, ByVal Cen1 As Long, ByVal Cen2 As Long, ByVal ArcTheta As Double) As Long
Declare Function P1240MotChgLineArcDV Lib "ads1240.dll" (ByVal Bid As Byte, ByVal DVValue As Long) As Long
Declare Function P1240MotStop Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal StopMode As Byte) As Long
Declare Function P1240MotRdReg Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal port As Integer, ByRef Value As Long) As Long
Declare Function P1240MotWrReg Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal port As Integer, ByVal Value As Long) As Long
Declare Function P1240MotRdMutiReg Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal port As Integer, ByRef valueX As Long, ByRef valueY As Long, ByRef valueZ As Long, ByRef ValueU As Long) As Long
Declare Function P1240MotWrMutiReg Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal port As Integer, ByVal valueX As Long, ByVal valueY As Long, ByVal valueZ As Long, ByVal ValueU As Long) As Long
Declare Function P1240MotSavePara Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte) As Long
Declare Function P1240MotExtMode Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal ExtMode As Byte, ByVal HandWheelPulse As Long) As Long
Declare Function P1240MotEnableEvent Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal EventX As Byte, ByVal EventYX As Byte, ByVal EventZ As Byte, ByVal EventU As Byte) As Long
Declare Function P1240MotCheckEvent Lib "ads1240.dll" (ByVal Bid As Byte, ByRef IrqEvent As Long, ByVal TimeOut As Long) As Long
Declare Function P1240MotDI Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByRef DIValue) As Long
Declare Function P1240MotDO Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByVal DOValue) As Long
Declare Function P1240MotHomeStatus Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte, ByRef HomeStatus) As Long
Declare Function P1240MotAxisBusy Lib "ads1240.dll" (ByVal Bid As Byte, ByVal axis As Byte) As Long
Declare Function P1240MotReset Lib "ads1240.dll" (ByVal Bid As Byte) As Long
Declare Function P1240InitialContiBuf Lib "ads1240.dll" (ByVal Bid As Byte, ByVal APNum As Long) As Long
Declare Function P1240SetContiData Lib "ads1240.dll" (ByVal Bid As Byte, ByRef pathdata As ContiPathData, ByVal PtNum As Long) As Long
Declare Function P1240StartContiDrive Lib "ads1240.dll" (ByVal Bid As Byte, ByVal MoveAxis As Byte, ByVal bufid As Byte) As Long
Declare Function P1240FreeContiBuf Lib "ads1240.dll" (ByVal bufid As Byte) As Long
Declare Function P1240GetCurContiNum Lib "ads1240.dll" (ByVal bufid As Byte, ByRef PresentSequenceNum As Long) As Long

