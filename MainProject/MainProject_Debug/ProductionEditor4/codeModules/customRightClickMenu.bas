Attribute VB_Name = "customRightClickMenu"
Option Explicit

Private Const GWL_WNDPROC = (-4)
Private Const WM_RBUTTONUP = &H205

Private Declare Function GetWindowLong Lib "user32" Alias _
"GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias _
"SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As _
Long, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias _
"CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd _
As Long, ByVal msg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long

Private m_lWndProc As Long
Private OriginalWndProc As Long

Public Sub WindowHook(hWnd As Long)
    m_lWndProc = SetWindowLong(hWnd, GWL_WNDPROC, AddressOf MessageCenter)
    OriginalWndProc = hWnd
End Sub

Public Sub WindowFree(hWnd As Long)

    SetWindowLong hWnd, GWL_WNDPROC, m_lWndProc
End Sub

Private Function MessageCenter(ByVal hWnd As Long, ByVal msg As Long, _
ByVal wParam As Long, ByVal lParam As Long) As Long
 
'Only suppress context menu for hooked control

    If msg = WM_RBUTTONUP Then
         If hWnd = OriginalWndProc Then
            editorForm.PopupMenu editorForm.mnuCustom
            MessageCenter = 0
           Exit Function
        End If
    End If
    
'Pass it on to the normal window processor.
MessageCenter = CallWindowProc(m_lWndProc, hWnd, msg, wParam, lParam)
End Function



