Attribute VB_Name = "LightIntensity"
Option Explicit

Public Sub Initialize_LightIntensity_Com()
    With editorForm.mscomLighIntensity
    
        .Settings = "9600,N,8,1"
        
        .DTREnable = False
        .InBufferCount = 0
        .OutBufferCount = 0
        .InputLen = 1
        .SThreshold = 1
        .RThreshold = 1
        .InputMode = 0
    
        If (.PortOpen = False) Then
            .CommPort = 1
            .PortOpen = True
        End If
    End With
End Sub

Public Sub SetLightIntensity(ByVal LightIntensity)
    With editorForm.mscomLighIntensity
        Turn_On_LightIntensity
        .Output = Chr(184)                   '1011 1000
        .Output = Chr(LightIntensity)
        Sleep (0.01)
    End With
End Sub

Public Sub Turn_On_LightIntensity()
    Dim Read_Value As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, Read_Value))
    Read_Value = Read_Value And &H400
    
    If Read_Value = &H0 Then
        checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, Read_Value))
        Read_Value = Read_Value Or &H400
        checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, Read_Value))
    End If
    
End Sub

Public Sub Turn_Off_LightIntensity()
    Dim Read_Value As Long
    Call SetLightIntensity(0)
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, Read_Value))
    Read_Value = Read_Value And &H400
    
    If Read_Value = &H400 Then
        Read_Value = Read_Value And &HFBFF
        checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, Read_Value))
    End If
End Sub




