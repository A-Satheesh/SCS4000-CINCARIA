Attribute VB_Name = "definitions"
'Definition constants for translator state machine

Global Const UnKnownState = 0
Global Const StartState = 1
Global Const linksArcRestartState = 2
Global Const ArcState = 3
Global Const ArcStopState = 4
Global Const linksArcEndState = 5
Global Const Line3DState = 6
Global Const Stop3DState = 7
Global Const End3DState = 8
Global Const linksArcStartState = 9
Global Const DotState = 10
Global Const PotState = 11
Global Const ArcStartState = 12
Global Const ArcEndState = 13

'Definitions constants for application
Global Const boardNum = 0
Global Const ZGearRatio = 1000
'Original Gear ratio
'Global Const ZGearRatio = 800 '250 pulses per mm New Motor Old motor is 5000

'Global Const txtCGTFilePath1 = ".\\translatortemplate\\newTranslator.cgt"
'Global Const txtCGTFilePath2 = ".\\translatortemplate\\newEpoxy.cgt"
Global Const txtCGTFilePath1 = "C:\MainProject\\ProductionEditor4\\translatortemplate\\newTranslator.cgt" 'NNO
Global Const txtCGTFilePath2 = "C:\MainProject\\ProductionEditor4\\translatortemplate\\newEpoxy.cgt"      'NNO

Public XYU_Interpolate As Boolean
Public ballScrew As Integer
Public fs, A
Public referenceSet, updateDispensePt, updateXDevOnly, updateYDevOnly, updateMoveHeightOnly, updateWithDrawalHeightOnly, readyStatus, busyStatus, errorStatus As Boolean
Public tempX, tempY, tempZ, referenceX, referenceY, referenceZ, systemTrackMoveHeight, zSystemTravelSpeed, xySystemTravelSpeed, SystemMoveHeight, systemHomeX, systemHomeY, systemHomeZ, SolventPosX, SolventPosY, SolventPosZ As Long
Public PreviousState As Integer
Public PrevPrevX, PrevPrevY, PrevPrevZ As Long
Public PrevX, PrevY, PrevZ As Long
Public error As Boolean
Public ContiBufferLines, SegmentPropertyLines As String
Public SegmentSeqNum As Long
Public xCen, yCen, ccw As Long
Public firstTime As Boolean
Public txtDataFilePath As String

Public glbOffsetX, glbOffsetY, glbOffsetZ As Long
Public glbOffsetChg As Boolean
Public offsetstk As OffsetStack
Public estopValue As Boolean
Public proceed As Boolean
Public stopValue As Boolean
Public abortValue As Boolean
Public prevDispenserValue As Boolean
Public fileDirty As Boolean
Public proceedWithAction As Boolean
Public selectNodeIndex As Long
Public tempCutPasteString As String
Public StartVelocity, MaxVelocity, AccelSpeed, AccelRate As Long
Public XYGearRatio As Long
Public directSoftHomeOption, externalDispenserControl As Boolean
Public needleOffsetX, needleOffsetY, needleOffsetZ_L, needleOffsetZ_R As Long
Public xOrgFid, yOrgFid As Double
Public deltaX, deltaY, deltaAngle As Double
Public s As String * 4096
Public step As Long
Public prevStep As Long
Public visionRetry As Boolean
Public doneFudicial As Boolean
Public systemMoveHeightDotPot As Long
Public presentX As Long, presentY As Long
Public origX As Long, origY As Long
Public firstFudicial As Boolean
Public camera2LightSetting As Long
Public printErrorLimit As Boolean
Public doingHome As Boolean
Public doingModify As Boolean

Public home_limit_flag As Boolean        'xu long
Public D_A_row_pitch As Double           'xu long
Public D_A_column_pitch As Double        'xu long
Public add_row_pitch As Double           'xu long
Public add_column_pitch As Double        'xu long
Public D_A_rows As Integer               'xu long
Public D_A_columns As Integer            'xu long

Public cmd_test_str As String            'xu long

Public Close_Emg, Modify, SingleClick As Boolean         'XW
Public Emergency_Stop, Ext, arcFlag As Boolean           'XW
Public Indicator, Reflector, NodeTypeNoChange As Boolean 'XW
Public StartingPoint As Long                             'XW
Public MovingMouse, NumLock As Boolean                   'XW
Public KeyOne, KeySeven As Boolean                       'XW
Public Click As Integer                                  'XW
Public RemoveSingleClick As Boolean                                              'XW
Public TotalLine As Integer                                                      'XW
Public CalculateX, CalculateY, CalculateZ                                        'XW
Public NoChange, NoChange2, NoChange3, Change As Boolean                         'XW
Public TravelSpeed, Last, ArcDelay As String                                     'XW

Public Previous_U As Long

Public ModifyOffsetX, ModifyOffsetY, ModifyOffsetZ As Double                     'XW
Public ExpandX, ExpandY, ExpandZ As Long    'This will be used while doing expanding
Public ExpandWithDrawSpeed As Long          'Use while doing expanding because X,Y,Z and this parameters will not be changing when we do the single click
Public FirstLineSelect As Boolean           'Check whether the first line is selected or not
Public ClickExpand As Boolean               'Flag the expanded button is pressed
Public NoEndArray As Boolean                'EndArray is deleted by the user
Public SaveX, SaveY, SaveZ As Long          'To regenerate array elements because we don't know x,y and z positions when we click on "dotArray"
Public Rows, Columns As Integer             'rows and colums of selected items
Public RepeatPattern As Boolean             'For part Array 3D circle
Public ReadRepeatString                     'For part array
Public SubArray, ErrorPartArray As Boolean
Public SingleLineSelected As Boolean
Public LeftDirection, RightDirection, UpDirection, DownDirection, UpLeftDirection, UpRightDirection, DownLeftDirection, DownRightDirection As Boolean
Public ContinuousLine As Boolean            'For 3D line dispense porcedure
Public ErrorKeyIn As Boolean                'Do a flag if the user keys in invalid value
Public StartLineDelay As Double             'Travel delay
Public FlagStartingPoint As Boolean         'To flag the first node
Public ActualSystemMoveHeight As Long       'To compare with 'systemHomeZ'

'Global Offset/Global "Dispense Time" and "Travel Speed"
Public FlagGlobalOffset As Boolean          'Just give a flag when the user click "apply"
Public FlagGlobalTimeSpeed As Boolean       'Just give a flag when the user click "apply"
Public GlobalOffsetX As Double             'Global Offset Value X
Public GlobalOffsetY As Double              'Global Offset Value Y
Public GlobalOffsetZ As Double             'Global Offset Value Z
Public GlobalDispenseTime As Double         'Global Dispensing Time
Public GlobalTravelSpeed As Integer         'Global TravelSpeed

'No_Fill_Area
Public Area_P1(0 To 2) As Double
Public Area_P2(0 To 2) As Double
Public Area_P3(0 To 2) As Double
Public Middle_Pt_OnOff As String            'Save dispensing valve ON/OFF for rectangel
Public Right_Needle_ON As Boolean           'Flage the last option for needles after loading the old pattern file
Public First_3D_Line As Boolean             'Flag for first 3D_Line

Public Red_Light As Boolean, Yellow_Light As Boolean, Green_Light As Boolean        'Show Tower Light
Public File_Load_Cancel As Boolean          'Flag if the user doesn't load a file.


Public originSpeed As Double        'Testing

'Just for testing (Not real use)
Public Pt_X As Long, Pt_Y As Long, Pt_Z As Long         'Save Previous Position

'''''''''''''''''''''''''''''''''''''''''
'   Offset between Camera and Needle    '
'''''''''''''''''''''''''''''''''''''''''

Public Travel_Speed_Rec As String, Last_Rec As String   'Save the travel speed for first rectangle
Public No_Start_Stop As Boolean                         'Flag not to do "Start" and "Stop" procedure for "L-Needle"
Public Offset_DistanceX_Camera_L_Needle As Long, Offset_DistanceY_Camera_L_Needle As Long   'Save the offset distance between Camera and L_Needle
Public Offset_DistanceX_Camera_R_Needle As Long, Offset_DistanceY_Camera_R_Needle As Long   'Save the offset distance between Camera and R_Needle

Public Remove_DispOff As Boolean                        'Flage the start of the line for "Start/Stop procedure"
Public reference_ZHigh As Boolean                       'Teach the z_high first before adding the position
Public Z_High As Long                                   'reference z_high
Public reference_R_ZHigh As Boolean                     'Teach the r_z_high first before adding the position
Public R_referenceZ As Long                             'Pattern referenceZ for right needle
Public R_Z_High As Long                                 'reference z_high for right needle
Public pointX As Long, pointY As Long, pointZ As Long

'''''''''''''''''''
'      login      '
'''''''''''''''''''
'NNO
Public mCancel As Boolean       'click cancel flag on login form
Public fPassword As String      'password parameter
Public loginsuccessful As Boolean   'login fail or pass?
Public confirmreset As Boolean      'flag for reset

Public leftside, rightside As Boolean
