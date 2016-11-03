VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Offsets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private x, y, z As Long

Public Sub setOffsets(ByVal xValue As Long, ByVal yValue As Long, ByVal zValue As Long)

    x = xValue
    y = yValue
    z = zValue
    
End Sub
Public Sub getOffsets(ByRef xValue As Long, ByRef yValue As Long, ByRef zValue As Long)

    xValue = x
    yValue = y
    zValue = z
    
End Sub

Public Function getOffsetX() As Long
    
    getOffsetX = x

End Function

Public Function getOffsetY() As Long
    
    getOffsetY = y

End Function
Public Function getOffsetZ() As Long
    
    getOffsetZ = z

End Function


