Attribute VB_Name = "userAuthentication"
Option Explicit

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Function Get_User_Name() As String

    Dim lpBuff As String * 25
    Dim ret As Long, UserName As String

    ' Get the user name minus any trailing spaces found in the name.
    ret = GetUserName(lpBuff, 25)
    UserName = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

    ' Return the User Name
    Get_User_Name = UserName
     
End Function
