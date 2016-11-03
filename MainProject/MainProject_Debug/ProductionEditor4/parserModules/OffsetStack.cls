VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OffsetStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MemberList() As Offsets
Private memberCount As Long


Friend Property Let Count(Value As Long)

If Value < 1 Then
      Erase MemberList
      memberCount = 0
   Else
      ReDim Preserve MemberList(0 To Value - 1)     'Change the size of the array
      memberCount = Value
   End If

End Property

Public Sub Clear()

   Erase MemberList
   memberCount = 0

End Sub

Property Get Count() As Long

   Count = memberCount

End Property



Property Get Member(ByVal Index As Long) As Offsets

   If Index >= 0 And Index < memberCount Then
      Set Member = MemberList(Index)
   Else
      Set Member = Nothing
   '   MsgBox "Nothing!"
   End If
End Property

Property Set Member(ByVal Index As Long, TheOffsets As Offsets)
   
   If Index >= 0 And Index < memberCount Then
      Set MemberList(Index) = TheOffsets
   End If

End Property

Public Sub Push(TheOffsets As Offsets)

   memberCount = memberCount + 1
   ReDim Preserve MemberList(0 To memberCount - 1)     'Change the size of the array
   Set MemberList(memberCount - 1) = TheOffsets

End Sub






Public Function Pop() As Offsets

    'Modified 12/11/2001
    If memberCount >= 1 Then
        Set Pop = MemberList(memberCount - 1)       'Indexing from 0
        Set MemberList(memberCount - 1) = Nothing   'List no longer points to the token
        memberCount = memberCount - 1
    Else
        Set Pop = Nothing
    End If

End Function

Public Function Top() As Offsets

   If memberCount >= 1 Then
      Set Top = MemberList(memberCount - 1)
   End If

End Function

