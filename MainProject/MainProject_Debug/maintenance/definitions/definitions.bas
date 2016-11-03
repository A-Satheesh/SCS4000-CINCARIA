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
Global Const ZGearRatio = 1000 '250 pulses per mm New Motor Old motor is 5000 (Change)
Global Const txtCGTFilePath1 = ".\\translatortemplate\\newTranslator.cgt"
Global Const txtCGTFilePath2 = ".\\translatortemplate\\newEpoxy.cgt"

Public ballScrew As Integer
Public fs, A
Public referenceSet, updateDispensePt, updateXDevOnly, updateYDevOnly, updateMoveHeightOnly, updateWithDrawalHeightOnly, readyStatus, busyStatus, errorStatus As Boolean
Public tempX, tempY, tempZ, referenceX, referenceY, referenceZ, systemTrackMoveHeight, systemMoveHeight, systemHomeX, systemHomeY, systemHomeZ As Long
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
Public needleOffsetX, needleOffsetY As Long
Public xOrgFid, yOrgFid As Double
Public deltaX, deltaY, deltaAngle As Double
Public s As String * 4096
Public step As Long
Public prevStep As Long
Public visionRetry As Boolean
Public home_limit_flag As Boolean        'xu long
Public Emergency_Stop, Ext As Boolean        'XW
Public Indicator, Reflector As Boolean       'XW
Public StartingPoint As Long                 'XW
Public MovingMouse, NumLock As Boolean       'XW
Public KeyOne, KeySeven As Boolean           'XW
Public LeftDirection, RightDirection, UpDirection, DownDirection, UpLeftDirection, UpRightDirection, DownLeftDirection, DownRightDirection As Boolean
Public ErrorKeyIn As Boolean                'Do a flag if the user keys in invalid value

Public Red_Light As Boolean, Yellow_Light As Boolean, Green_Light As Boolean    'Show Tower Light
Public loadVisionCalibration As Boolean

'''''''''''''''''''
'      login      '
'''''''''''''''''''
'NNO
Public mCancel As Boolean
Public fPassword As String
Public loginsuccessful As Boolean
Public confirmreset As Boolean

Public Master_LN_Position As Long    'Master Z position for left needle
Public Master_RN_Position As Long    'Master Z position for right needle
