Attribute VB_Name = "dispenserRoutines"
Public Sub readRegistryOptions()

    Dim tempStr As String

    'Made the ballScrew to be "1"
    tempStr = 1
    'tempStr = GetStringSetting("EpoxyDispenser", "Setup", "BallScrew")

    If tempStr = "0" Then
        ballScrew = 0
    Else
        ballScrew = 1
    End If
    
    tempStr = GetStringSetting("EpoxyDispenser", "Setup", "DirectSoftHome")

    If tempStr = "0" Then
        directSoftHomeOption = False
    Else
        directSoftHomeOption = True
    End If

    tempStr = GetStringSetting("EpoxyDispenser", "Setup", "externalDispenserControl")

    If tempStr = "0" Then
        externalDispenserControl = False
    Else
        externalDispenserControl = True
    End If

End Sub

Public Sub determineProfile()

    If ballScrew = 1 Then
        XYGearRatio = 1000 '250 pulses per mm (New Motor) Old motor is 5000
        StartVelocity = 1000
        MaxVelocity = 2000000
        '0.2G
        'AccelSpeed = 2000000
        'AccelRate = 15264000
       
        AccelSpeed = 5000000
        AccelRate = 30000000
        'origin
        'XYGearRatio = 250 '250 pulses per mm (New Motor) Old motor is 5000
        'StartVelocity = 250
        'MaxVelocity = 2000000
        'AccelSpeed = 260000
        'AccelRate = 500000
    Else
        XYGearRatio = 50000 / 75 '50000/75 pulses per mm
        StartVelocity = 1000
        MaxVelocity = 800000
        AccelSpeed = 5000000
        AccelRate = 800000000
    End If
    
    'XW
    needleOffsetX = convertToPulses(GetStringSetting("EpoxyDispenser", "NeedleOffset", "XOff", "0"), X_axis)
    needleOffsetY = convertToPulses(GetStringSetting("EpoxyDispenser", "NeedleOffset", "YOff", "0"), Y_axis)
    needleOffsetY = needleOffsetY * (-1)
End Sub

Public Sub PTPToXYZ(ByVal X As Long, ByVal Y As Long, ByVal z As Long)
    readyStatus = False
    busyStatus = True
    
    'To get the +ve direction   'XW
    Y = Y * (-1)
    z = z * (-1)
    
    'Exit the task that hasn't finished
    If Emergency_Stop = True Then
        readyStatus = True
        busyStatus = False
        Exit Sub
    Else
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, systemTrackMoveHeight, 0))
    End If
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
        DoEvents
    Loop
    
    'Exit the task that hasn't finished
    If Emergency_Stop = True Then
        readyStatus = True
        busyStatus = False
        Exit Sub
    Else
        checkSuccess (P1240MotPtp(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, X, Y, 0, 0))
    End If
    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS) 'Loop while XY motor is still spinning
        DoEvents
    Loop
    
    'Exit the task that hasn't finished
    If Emergency_Stop = True Then
        readyStatus = True
        busyStatus = False
        Exit Sub
    Else
        checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, z, 0))
    End If
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
        DoEvents
    Loop
    
    Emergency_Stop = False
    readyStatus = True
    busyStatus = False
End Sub

Public Sub setSpeed(Speed As Long)
    
    Dim AccelSpeed, AccelSpeedZ As Double
    Dim AccelRate, AccelRateZ As Double
    Dim factor, factorZ As Double
    
    If ballScrew = 1 Then
        '0.12G
        AccelSpeedZ = 1200000 '1200000
        AccelRateZ = 9158400 '9158400

        '0.5G
        AccelSpeed = 5000000
        AccelRate = 30000000
        
'        AccelSpeedZ = 260000       'origin
'        AccelRateZ = 500000
'        AccelSpeed = 260000
'        AccelRate = 500000
    
        'If Speed <= 10 Then
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            'AccelRateZ = AccelSpeedZ / 0.05

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.1
            'AccelRate = AccelSpeed / 0.03
        'ElseIf Speed <= 90 Then
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            'AccelRateZ = AccelSpeedZ / 0.05

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.2
            'AccelRate = AccelSpeed / 0.06
        'Else
            'factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.8
            'AccelRateZ = AccelSpeedZ / 0.08

            'factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            'AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.7
            'AccelRate = AccelSpeed / 0.2

        'End If
    Else
    
        If Speed <= 10 Then
            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.1
            AccelRateZ = AccelSpeedZ / 0.05
            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.01
            AccelRate = AccelSpeed / 0.05

        Else
            factorZ = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeedZ = (convertSpeed(Speed, Z_axis) - (convertSpeed(Speed, Z_axis) / factorZ)) / 0.4
            AccelRateZ = AccelSpeedZ / 0.05
            factor = (Exp(-1 * CLng(Speed) / 100) ^ 1.1) * 100
            AccelSpeed = (convertSpeed(Speed, X_axis Or Y_axis) - (convertSpeed(Speed, X_axis Or Y_axis) / factor)) / 0.05
            AccelRate = AccelSpeed / 0.05
        End If
    
    End If
    
    checkSuccess (P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(Speed, Z_axis), 2000000, AccelSpeedZ, AccelRateZ))
    checkSuccess (P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(Speed, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))
    'For U_axis (startig velocity, speed,max speed, acceleration, rate)
    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 500, 8300, 8300, 53000, 9000000))
    'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 1000, 16600, 16600, 106000, 18000000))
    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 3000, 10000, 50000, 20000))
    
End Sub

Public Sub displayCoOrds()

    Dim xValue, yValue, zValue, uValue As Long

    checkSuccess (P1240MotRdMutiReg(boardNum, Z_axis Or Y_axis Or X_axis, Lcnt, xValue, yValue, zValue, uValue))
    
    'To get the +ve direction   'XW
    yValue = yValue * (-1)
    zValue = zValue * (-1)
    
    With mainForm
        .xCoOrd.Text = convertToMM(xValue, X_axis)
        .yCoOrd.Text = convertToMM(yValue, Y_axis)
        .zCoOrd.Text = convertToMM(zValue, Z_axis)
        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, RR2, xValue, yValue, zValue, uValue))
        xValue = xValue And &HC
        yValue = yValue And &HC
        zValue = zValue And &HC
        uValue = uValue And &H8
    
        'If (xValue <> 0) Or (yValue <> 0) Or (zValue <> 0) Then        'origin
        If ((xValue <> 0) Or (yValue <> 0) Or (zValue <> 0) Or (uValue <> 0)) And home_limit_flag = False Then 'xu long, do not display limit errors while homing
            .LimitReachedLabel.Visible = True
        Else
            .LimitReachedLabel.Visible = False
        End If
    End With
End Sub

Public Function unInitializeBoard() As Boolean
    'Close board
    checkSuccess (P1240MotDevClose(boardNum))
End Function

Public Function initializeBoard() As Boolean

    Dim resetValue As Long
    Dim ValueX, ValueY, ValueZ As Long
    Dim DriverXYZ As Long
    
    initializeBoard = False
    busyStatus = True
    
    If (checkSuccess(P1240MotDevAvailable(BoardAvailable))) Then
        'Open board
        If (checkSuccess(P1240MotDevOpen(boardNum))) Then
            If checkSuccess(P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR2, ValueX, ValueY, ValueZ, ValueU)) Then
                ValueX = (ValueX Or CLng("&HA8C0"))
                ValueY = (ValueY Or CLng("&HA8C0"))
                ValueZ = (ValueZ Or CLng("&HA8C0"))
                ValueU = (ValueU Or CLng("&H8C0"))
                
                'If checkSuccess((P1240MotWrMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR2, ValueX, ValueY, ValueZ, ValueU))) Then
                'If checkSuccess((P1240MotWrMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR2, &HA8C0, &HA8C0, &HA8C0, &H8C0))) Then
                If checkSuccess((P1240MotWrMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR2, &H800, &H800, &H800, &H800))) Then
                    If checkSuccess(P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR3, ValueX, ValueY, ValueZ, ValueU)) Then
                        ValueX = (ValueX And &HFF7F)
                        ValueY = (ValueY And &HFF7F)
                        ValueZ = (ValueZ And &HFF7F)
                        ValueU = (ValueU And &HFF7F)
                        
                        If checkSuccess((P1240MotWrMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR3, ValueX, ValueY, ValueZ, ValueU))) Then
                            ValueX = &H0
                            ValueY = &H0
                            ValueZ = &H0
                            ValueU = &H0
                                
                            If checkSuccess(P1240MotWrMutiReg(boardNum, X_axis Or Y_axis Or Z_axis Or U_axis, WR1, ValueX, ValueY, ValueZ, ValueU)) Then
                                If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(30, Z_axis), 500000, 1500000, 15000000))) Then
                                    If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(30, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
                                        checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 1500, 10000, 50000, 20000))
        
                                        'Check whether the e-stop release or not
                                        'if not, it will show a message to user.
                                        If (checkSuccess(P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, ValueX, ValueY, ValueZ, 0))) Then
                                            ValueX = (ValueX And &H20)
                                            ValueY = (ValueY And &H20)
                                            ValueZ = (ValueZ And &H20)
                                            
                                            If (ValueX <> 0) Or (ValueY <> 0) Or (ValueZ <> 0) Then
                                                Do While (ValueX <> 0) Or (ValueY <> 0) Or (ValueZ <> 0)
                                            
                                                    'Enable Alarm
                                                    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, ValueU))
                                                    ValueU = ValueU Or &H800
                                                    checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, ValueU))
                                                
                                                    frmRelease.Show (vbModal)
                                                
                                                    Call P1240MotRdMutiReg(boardNum, X_axis Or Y_axis Or Z_axis, RR2, ValueX, ValueY, ValueZ, 0)
                                                    ValueX = (ValueX And &H20)
                                                    ValueY = (ValueY And &H20)
                                                    ValueZ = (ValueZ And &H20)
                                                Loop
        
                                                'Close the alarm
                                                checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, ValueU))
                                                ValueU = ValueU And &HF7FF
                                                checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, ValueU))
                                            
                                                Call Sleep(0.02)
                                            
                                                'Do reset
                                                checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, resetValue))
                                                resetValue = resetValue Or &H200
                                                checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
        
                                                Call Sleep(0.3)
                                            
                                                resetValue = resetValue And &HFDFF
                                                checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
                                    
                                                Call Sleep(1)
                                                
                                                Homing_Sequence
                                                
                                                Servo_On
                                        
                                                initializeBoard = True
                                            Else
                                                Homing_Sequence
                                                
                                                Servo_On
                                        
                                                initializeBoard = True
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox "Board open fail!"
        End If
    Else
        MsgBox "There is no avalibale board. Please check the hardware!"
    End If
    
    busyStatus = False
    
End Function

Public Sub Servo_On()
    Dim DriverXYZ As Long
    
    'Servo ON (XW)
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, DriverXYZ))
    DriverXYZ = (DriverXYZ Or &H700)
    checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, DriverXYZ))
    
    Call Sleep(0.6)
End Sub

Public Sub Servo_Off()
    Dim DriverXYZ As Long
    
    'Servo OFF (XW)
    checkSuccess (P1240MotRdReg(boardNum, X_axis, WR3, DriverXYZ))
    DriverXYZ = (DriverXYZ And &HF8FF)
    checkSuccess (P1240MotWrReg(boardNum, X_axis, WR3, DriverXYZ))
                
    Call Sleep(0.03)
End Sub

Public Function moveToHome() As Boolean
    Dim value, ValueX, ValueY, ValueZ, ValueU As Long
   
    Ext = False
    moveToHome = False
    readyStatus = False
    busyStatus = True
    mainForm.SetFocusTimer.Enabled = False
    
    'Left slider go up first before do homing
    Call Leftslider_go_up
    
    'If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(30, Z_axis), 500000, 1500000, 15000000))) Then

        If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
                DoEvents
                
                If (Emergency_Stop = True) And (Ext = True) Then
                    readyStatus = True
                    busyStatus = False
                    Emergency_Stop = False
                    Exit Function
                End If
            Loop
            
            
            '''''''''''''''''''''
            '   U_axis (Homing) '
            '''''''''''''''''''''
    '        If (checkSuccess(P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 200, 10000, 50000, 20000))) Then
               
                checkSuccess (P1240MotHome(boardNum, U_axis))
                Do While (P1240MotAxisBusy(boardNum, U_axis) <> SUCCESS)
                    DoEvents
                    If (Emergency_Stop = True) And (Ext = True) Then
                        readyStatus = True
                        busyStatus = False
                        Emergency_Stop = False
                        Exit Function
                    End If
                Loop
                
    '            If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(20, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
    '                If (checkSuccess(P1240MotCmove(boardNum, X_axis, 1))) Or (checkSuccess(P1240MotCmove(boardNum, Y_axis, 0))) Then
    '                'Move X and Y motors in clockwise direction and anti-clockwise direction
    '                'If (checkSuccess(P1240MotCmove(boardNum, X_axis Or Y_axis, 0 Or 1))) Then
    '                    ValueX = 0
    '                    ValueY = 0
    '                    Do While (((ValueX And &H8) <> &H8) Or ((ValueY And &H4) <> &H4))
    '                       checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, ValueX, ValueY, ValueZ, ValueU))
    '                        If ((ValueY And &H4) = &H4) Then
    '                            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    '                        End If
    '                        If ((ValueX And &H8) = &H8) Then
    '                            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    '                        End If
    '                        DoEvents
    '                        If (Emergency_Stop = True) And (Ext = True) Then
    '                            readyStatus = True
    '                            busyStatus = False
    '                            Emergency_Stop = False
    '                            Exit Function
    '                        End If
    '                    Loop
                        
    '                    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
    '                    Loop
                        
                        If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
                            Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
                                DoEvents
                                If (Emergency_Stop = True) And (Ext = True) Then
                                    readyStatus = True
                                    busyStatus = False
                                    Emergency_Stop = False
                                    Exit Function
                                End If
                            Loop
                            moveToHome = True
                        End If
    '                End If
    '            End If
    '        'End If
        End If
    'End If
    
    ''Change from T_curve to S_curve
    ''If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(30, Z_axis), 2000000, 1200000, 9158400))) Then
    'If (checkSuccess(P1240MotAxisParaSet(boardNum, Z_axis, Z_axis, 1000, convertSpeed(20, Z_axis), 200000, 1000000, 9158400))) Then
    '    If checkSuccess(P1240MotCmove(boardNum, Z_axis, 0)) Then  'Move Z in clockwise direction
    '        value = 0
    
    '        Do While ((value And &H4) <> &H4)   'Do loop if Z Limit switch still not reached
    '            checkSuccess (P1240MotRdReg(boardNum, Z_axis, RR2, value))
    
    '            If ((value And &H4) = &H4) Then 'Do immediate stop on Z axis
    '                checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
    '            End If
    
    '            DoEvents
    '            If (Emergency_Stop = True) And (Ext = True) Then
    '                readyStatus = True
    '                busyStatus = False
    '                Emergency_Stop = False
    '                Exit Function
    '            End If
    '        Loop
    '        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS)  'Loop while Z motor is still spinning
    '        Loop
    '        If (checkSuccess(P1240MotHome(boardNum, Z_axis))) Then
    '            Do While (P1240MotAxisBusy(boardNum, Z_axis) <> SUCCESS) 'Loop while Z motor is still spinning
    '                DoEvents
    '                If (Emergency_Stop = True) And (Ext = True) Then
    '                    readyStatus = True
    '                    busyStatus = False
    '                    Emergency_Stop = False
    '                    Exit Function
    '                End If
    '            Loop
    
    '            '''''''''''''''''''''
    '            '   U_axis (Homing) '
    '            '''''''''''''''''''''
    '            'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 300, 300, 8300, 53000, 9000000))
    '            'checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 600, 600, 16600, 106000, 18000000))
    '            checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 200, 10000, 50000, 20000))
    '            checkSuccess (P1240MotCmove(boardNum, U_axis, 8))
    '            ValueU = 0
    '
    '            Do While ((ValueU And &H8) <> &H8)
    '                checkSuccess (P1240MotRdReg(boardNum, U_axis, RR2, ValueU))
   
    '                If ((ValueU And &H8) = &H8) Then 'Do immediate stop on U axis
    '                    checkSuccess (P1240MotStop(boardNum, U_axis, 8))
    '                End If
    '                DoEvents
    '                If (Emergency_Stop = True) And (Ext = True) Then
    '                    readyStatus = True
    '                    busyStatus = False
    '                    Emergency_Stop = False
    '                    Exit Function
    '                End If
    '            Loop
    '            Do While (P1240MotAxisBusy(boardNum, U_axis) <> SUCCESS)
    '            Loop
    
    '            checkSuccess (P1240MotHome(boardNum, U_axis))
    '            Do While (P1240MotAxisBusy(boardNum, U_axis) <> SUCCESS)
    '                DoEvents
    '                If (Emergency_Stop = True) And (Ext = True) Then
    '                    readyStatus = True
    '                    busyStatus = False
    '                    Emergency_Stop = False
    '                    Exit Function
    '                End If
    '            Loop
    
    '            If (checkSuccess(P1240MotAxisParaSet(boardNum, X_axis Or Y_axis, X_axis Or Y_axis, StartVelocity, convertSpeed(20, X_axis Or Y_axis), MaxVelocity, AccelSpeed, AccelRate))) Then
    '                If (checkSuccess(P1240MotCmove(boardNum, X_axis, 1))) Or (checkSuccess(P1240MotCmove(boardNum, Y_axis, 0))) Then
    '                'Move X and Y motors in clockwise direction and anti-clockwise direction
    '                'If (checkSuccess(P1240MotCmove(boardNum, X_axis Or Y_axis, 0 Or 1))) Then
    '                    ValueX = 0
    '                    ValueY = 0
    '                    Do While (((ValueX And &H8) <> &H8) Or ((ValueY And &H4) <> &H4))
    '                        checkSuccess (P1240MotRdMutiReg(boardNum, X_axis Or Y_axis, RR2, ValueX, ValueY, ValueZ, ValueU))
    '                        If ((ValueY And &H4) = &H4) Then
    '                            checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
    '                        End If
    '                        If ((ValueX And &H8) = &H8) Then
    '                            checkSuccess (P1240MotStop(boardNum, X_axis, 1))
    '                        End If
    '                        DoEvents
    '                        If (Emergency_Stop = True) And (Ext = True) Then
    '                            readyStatus = True
    '                            busyStatus = False
    '                            Emergency_Stop = False
    '                            Exit Function
    '                        End If
    '                    Loop
    '                    Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
    '                    Loop
    '                    If (checkSuccess(P1240MotHome(boardNum, X_axis Or Y_axis))) Then
    '                        Do While (P1240MotAxisBusy(boardNum, X_axis Or Y_axis) <> SUCCESS)
    '                            DoEvents
    '                            If (Emergency_Stop = True) And (Ext = True) Then
    '                                readyStatus = True
    '                                busyStatus = False
    '                                Emergency_Stop = False
    '                                Exit Function
    '                            End If
    '                        Loop
    '                        moveToHome = True
    '                   End If
    '                End If
    '            End If
    '        End If
    '    End If
    'End If
    
    'For U_axis
    checkSuccess (P1240MotAxisParaSet(boardNum, U_axis, U_axis, 200, 3000, 10000, 50000, 20000))
    
    readyStatus = True
    busyStatus = False
    mainForm.SetFocusTimer.Enabled = True
End Function

Public Sub Homing_Sequence()
    
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, PLmt, 1000000, 1000000, 1000000, 1000000))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, NLmt, -1000000, -1000000, -1000000, -1000000))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, HomeType, 4, 4, 4, 4))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, HomeP0Dir, 1, 0, 0, 1))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, HomeP0Speed, 30000, 30000, 30000, 500))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, HomeOffsetSpeed, 10000, 10000, 10000, 100))
    checkSuccess (P1240MotWrMutiReg(boardNum, XYZU_axis, HomeOffset, 15000, -15000, -10000, 50))
        
End Sub

Public Function convertSpeed(ByVal Speed As Long, ByVal axis As Long) As Long
    If axis = Z_axis Then
        convertSpeed = CLng(Speed) * ZGearRatio
    Else
        convertSpeed = CLng(Speed) * XYGearRatio
    End If
End Function

Public Function checkSuccess(ByVal returncode As Long) As Boolean
    If returncode = 0 Then
        checkSuccess = True
    Else
        errorStatus = True
        Beep
        
        If (returncode = 48) Then
            ShowError (returncode)
        End If
        
        checkSuccess = False
        
    End If
End Function

'**************************************************************************
'    Error status checking
'**************************************************************************
Sub ShowError(ByVal Err As Long)
    Dim ErrMsg As String
    
    If Err = 0 Then
        Exit Sub
    End If
    
    Select Case Err
        Case BoardNumErr
            ErrMsg = "Error 0x0001: Board Number Error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case CreateDriverFail
            ErrMsg = "Error 0x0002: System return error when open kernel driver"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case CallDriverFail
            ErrMsg = "Error 0x0003: System return error when call kernel driver"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case RegistryOpenFail
            ErrMsg = "Error 0x0004: Open Registry file error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case RegistryReadFail
            ErrMsg = "Error 0x0005: Read Registry file error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case AxisNumErr
            ErrMsg = "Error 0x0006: Axis Number Error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderRGErr
            ErrMsg = "Error 0x0007: None RG value can suit for MDV,AC and AK value"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverRGErr
            ErrMsg = "Error 0x0008: None RG value can suit for MDV,AC and AK value"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderSVErr
            ErrMsg = "Error 0x0009: SV value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverSVErr
            ErrMsg = "Error 0x000A: SV value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverMDVErr
            ErrMsg = "Error 0x000B: MDV value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderDVErr
            ErrMsg = "Error 0x000C: DV value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverDVErr
            ErrMsg = "Error 0x000D: DV value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderACErr
            ErrMsg = "Error 0x000E: AC value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverACErr
            ErrMsg = "Error 0x000F: AC value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderAKErr
            ErrMsg = "Error 0x0010: AK value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverAKErr
            ErrMsg = "Error 0x0011: AK value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverPLmtErr
            ErrMsg = "Error 0x0012: Over Maximum P direction limited"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverNLmtErr
            ErrMsg = "Error 0x0013: Over Maximum N direction limited"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case MaxMoveDistErr
            ErrMsg = "Error 0x0014: Moving distance is over single maximum distance"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case AxisDrvBusy
            ErrMsg = "Error 0x0015: One of the assignment axis is busy"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case RegItemErr
            ErrMsg = "Error 0x0016: The assigned Register have no defined"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ParaValueErr
            ErrMsg = "Error 0x0017: The parameter out of range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ParaValueOverRange
            ErrMsg = "Error 0x0018: The drive speed parameter out of range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ParaValueUnderRange
            ErrMsg = "Error 0x0019: The drive speed parameter out of range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case AxisHomeBusy
            ErrMsg = "Error 0x001A: The Axis is in Home process busy"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case AxisExtBusy
            ErrMsg = "Error 0x001B: The Operating Axis is in external input controlling mode"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case RegistryWriteFail
            ErrMsg = "Error 0x001C: Write Registry file error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ParaValueOverErr
            ErrMsg = "Error 0x001D: One of Motion function parameter over range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ParaValueUnderErr
            ErrMsg = "Error 0x001E: One of the assigned parameter value is under valid value"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OverDCErr
            ErrMsg = "Error 0x001F: DC value is over valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderDCErr
            ErrMsg = "Error 0x0020: DC value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case UnderMDVErr
            ErrMsg = "Error 0x0021: MDV value is under valid range"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case RegistryCreateFail
            ErrMsg = "Error 0x0022: Create Registry file fail"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case CreateThreadErr
            ErrMsg = "Error 0x0023: Create internal Thread error"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case HomeSwStop
            ErrMsg = "Error 0x0024: Return Home process be stopped by P1240MotStop function"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case ChangeSpeedErr
            ErrMsg = "Error 0x0025: The speed parameter is invalid"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case DOPortAsDriverStatus
            ErrMsg = "Error 0x0026: Output DO when DO type is assigned to indicator function"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case OpenEventFail
            ErrMsg = "Error 0x0030: System return error when create IRQ event"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case DeviceCloseErr
            ErrMsg = "Error 0x0032: System return error when in P1240MotDevClose process"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case HomeEMGStop
            ErrMsg = "Error 0x0040: Return home process be stopped by hardware emergency stop input"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case HomeLMTPStop
            ErrMsg = "Error 0x0041: Return home process be stopped by hardware positive limited switch"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case HomeLMTNStop
            ErrMsg = "Error 0x0042: Return home process be stopped by hardware negative limited switch"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case HomeALARMStop
            ErrMsg = "Error 0x0043: Return home process be stopped by software limited switch"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case AllocateBufferFail
            ErrMsg = "Error 0x0050: System return error when buffer allocate"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case BufferReAllocate
            ErrMsg = "Error 0x0051: The assigned buffer have been allocated and try to reallocate again"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case FreeBufferFail
            ErrMsg = "Error 0x0052: The handle of assigned Buffer is NULL or has been freed before"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case FirstPointNumberFail
            ErrMsg = "Error 0x0053: The first data hasn't been set and try to set other number data"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case PointNumExceedAllocatedSize
            ErrMsg = "Error 0x0054: Current buffer is full"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case BufferNoneAllocate
            ErrMsg = "Error 0x0055: The assigned Buffer number hasn't created"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case SequenceNumberErr
            ErrMsg = "Error 0x0056: The point number is not continue number of previous setting data"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case PathTypeErr
            ErrMsg = "Error 0x0057: Path type is not IPO_L2, IPO_L3, IPO_CW or IPO_CCW."
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case PathTypeMixErr
            ErrMsg = "Error 0x0060: Continue Data buffer have mixed IPO_L2 and IPO_L3 or IPO_CW/IPO_CCW and IPO_L3 path type"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case BufferDataNotEnough
            ErrMsg = "Error 0x0061: Buffer contain only one data"
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
        Case Else
            If Err > &HF Then
               ErrMsg = "Error 0x00" + Hex(Err)
            Else
               ErrMsg = "Error 0x000" + Hex(Err)
            End If
            errCode = MsgBox(ErrMsg, vbOKOnly, "PCI-1240 Message")
    End Select

End Sub



