Attribute VB_Name = "guiRoutines"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

'''''''''''''''''''''''''''''
'   Left slinder go down    '
'''''''''''''''''''''''''''''
Public Sub Leftslider_go_down()
    Dim ReadValue As Long
    
    'Left slider
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.5)
    
    'Left_Down_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Down_Sensor)
    L_Down_Sensor = L_Down_Sensor And &H4
    
    If Not (L_Down_Sensor = 0) Then
        MsgBox "Left Cylinder have some problem!"
    End If
End Sub

'''''''''''''''''''''''''''''
'   Left slinder go up      '
'''''''''''''''''''''''''''''
Public Sub Leftslider_go_up()
    Dim ReadValue As Long
    Dim L_Up_Sensor As Byte
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(1.5)
    
    'Left_Up_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Up_Sensor)
    L_Up_Sensor = L_Up_Sensor And &H1
    
    If Not (L_Up_Sensor = 0) Then
        MsgBox "Left Cylinder have some problem!"
    End If
End Sub

Public Sub initializeInputParams()
    editorForm.dispensePtX.Text = ""
    editorForm.dispensePtY.Text = ""
    editorForm.dispensePtZ.Text = ""
    editorForm.dispenseTime.Text = "1.0"
    editorForm.potDepth.Text = "0"
    editorForm.depthSpeed.Text = "10"
    editorForm.endDispenseHeight.Text = "0"
    editorForm.delay.Text = "1.0"
    editorForm.DispenseSpeed.Text = "10"
    editorForm.dispenseOnOff.Value = Checked
    editorForm.retractDelay.Text = "1.0"
    editorForm.withdrawalSpeed.Text = "10"
    editorForm.WithDrawalZ.Text = "0"
    editorForm.moveHeight.Text = convertToMM(systemMoveHeight, Z_axis)
    editorForm.xRepeatNum.Text = "1"
    editorForm.yRepeatNum.Text = "1"
    editorForm.xDev.Text = "10"
    editorForm.yDev.Text = "10"
    editorForm.PathFileName.Text = ""
    editorForm.jogSpeedSlider.Value = 28
    referenceSet = False
End Sub
Public Sub disableAllInputParams()
    With editorForm
        .dispensePtZ.Enabled = False
        .dispensePtZLabel.Enabled = False
        .dispenseTime.Enabled = False
        .DispenseTimeLabel.Enabled = False
        .UpDownDispenseTime.Enabled = False
        .potDepth.Enabled = False
        .UpDownPotDepth.Enabled = False
        .PotDepthLabel.Enabled = False
        .depthSpeed.Enabled = False
        .UpDownDepthSpeed.Enabled = False
        .depthSpeedLabel.Enabled = False
        .PotDepthLabel.Enabled = False
        .endDispenseHeight.Enabled = False
        .UpDownEndDispenseHeight.Enabled = False
        .endDispenseHeightLabel.Enabled = False
        .delay.Enabled = False
        .UpDownDelay.Enabled = False
        .delayLabel.Enabled = False
        .DispenseSpeed.Enabled = False
        .UpDownDispenseSpeed.Enabled = False
        .dispenseSpeedLabel.Enabled = False
        .dispenseOnOff.Enabled = False
        .retractDelay.Enabled = False
        .UpDownRetractDelay.Enabled = False
        .retractDelayLabel.Enabled = False
        .withdrawalSpeed.Enabled = False
        .UpDownWithDrawalSpeed.Enabled = False
        .withDrawalSpeedLabel.Enabled = False
        .WithDrawalZ.Enabled = False
        .decideWithdrawalHeight.Enabled = False
        .withDrawalZLabel.Enabled = False
        .moveHeight.Enabled = False
        .decideMoveHeight.Enabled = False
        .moveHeightLabel.Enabled = False
        .xRepeatNum.Enabled = False
        .UpDownXRepeatNum.Enabled = False
        .xRepeatNumLabel.Enabled = False
        .yRepeatNum.Enabled = False
        .UpDownYRepeatNum.Enabled = False
        .yRepeatNumLabel.Enabled = False
        .xDev.Enabled = False
        .decideXYDev.Enabled = False
        .xDevLabel.Enabled = False
        .yDev.Enabled = False
        .yDevLabel.Enabled = False
        .PathFileName.Enabled = False
        .loadPartArray.Enabled = False
        .pathFileNameLabel.Enabled = False
    End With
End Sub
Public Sub enableInputParams()

    'Reference node
    If editorForm.NodeType.ListIndex = 0 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
    End If
     
    'Dot node
    If editorForm.NodeType.ListIndex = 2 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        
        editorForm.xDev.Enabled = True
        editorForm.xDevLabel.Enabled = True
        editorForm.decideXYDev.Enabled = True
        editorForm.yDev.Enabled = True
        editorForm.yDevLabel.Enabled = True
        
        editorForm.dispenseTime.Enabled = True
        editorForm.DispenseTimeLabel.Enabled = True
        editorForm.UpDownDispenseTime.Enabled = True
        
        editorForm.xRepeatNum.Enabled = True
        editorForm.UpDownXRepeatNum.Enabled = True
        editorForm.xRepeatNumLabel.Enabled = True
        editorForm.yRepeatNum.Enabled = True
        editorForm.UpDownYRepeatNum.Enabled = True
        editorForm.yRepeatNumLabel.Enabled = True
        
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
    End If

    'dot Potting node
    If editorForm.NodeType.ListIndex = 3 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        
        editorForm.xDev.Enabled = True
        editorForm.xDevLabel.Enabled = True
        editorForm.decideXYDev.Enabled = True
        editorForm.yDev.Enabled = True
        editorForm.yDevLabel.Enabled = True
        
        editorForm.xRepeatNum.Enabled = True
        editorForm.UpDownXRepeatNum.Enabled = True
        editorForm.xRepeatNumLabel.Enabled = True
        editorForm.yRepeatNum.Enabled = True
        editorForm.UpDownYRepeatNum.Enabled = True
        editorForm.yRepeatNumLabel.Enabled = True
        
        editorForm.dispenseTime.Enabled = True
        editorForm.DispenseTimeLabel.Enabled = True
        editorForm.UpDownDispenseTime.Enabled = True
        
        editorForm.potDepth.Enabled = True
        editorForm.PotDepthLabel.Enabled = True
        editorForm.UpDownPotDepth.Enabled = True
        editorForm.depthSpeed.Enabled = True
        editorForm.depthSpeedLabel.Enabled = True
        editorForm.UpDownDepthSpeed.Enabled = True
        
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
    End If
    
    'line Potting node
    If editorForm.NodeType.ListIndex = 5 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        
        editorForm.xDev.Enabled = True
        editorForm.xDevLabel.Enabled = True
        editorForm.decideXYDev.Enabled = True
        editorForm.yDev.Enabled = True
        editorForm.yDevLabel.Enabled = True
        
        editorForm.xRepeatNum.Enabled = True
        editorForm.UpDownXRepeatNum.Enabled = True
        editorForm.xRepeatNumLabel.Enabled = True
        editorForm.yRepeatNum.Enabled = True
        editorForm.UpDownYRepeatNum.Enabled = True
        editorForm.yRepeatNumLabel.Enabled = True
        
        editorForm.potDepth.Enabled = True
        editorForm.PotDepthLabel.Enabled = True
        editorForm.UpDownPotDepth.Enabled = True
        editorForm.depthSpeed.Enabled = True
        editorForm.depthSpeedLabel.Enabled = True
        editorForm.UpDownDepthSpeed.Enabled = True
        editorForm.endDispenseHeight.Enabled = True
        editorForm.endDispenseHeightLabel.Enabled = True
        editorForm.UpDownEndDispenseHeight.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
    End If
    
    'line Start node
    If editorForm.NodeType.ListIndex = 6 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'line End node
    If editorForm.NodeType.ListIndex = 7 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
        
    End If
        
    'arc Start Node
    If editorForm.NodeType.ListIndex = 9 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'arc Point Node
    If editorForm.NodeType.ListIndex = 10 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'arc End Node
    If editorForm.NodeType.ListIndex = 11 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
    End If
    
    'links Line Point Node
    If editorForm.NodeType.ListIndex = 13 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
    
    'links Arc Restart Node
    If editorForm.NodeType.ListIndex = 14 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
        
    'links Arc Start Node
    If editorForm.NodeType.ListIndex = 15 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
        
    'links Arc End node
    If editorForm.NodeType.ListIndex = 16 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
    
    'rectCorner1 Node
    If editorForm.NodeType.ListIndex = 18 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'rectCorner2 Node
    If editorForm.NodeType.ListIndex = 19 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'rectCorner3 Node
    If editorForm.NodeType.ListIndex = 20 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
        editorForm.retractDelay.Enabled = True
        editorForm.UpDownRetractDelay.Enabled = True
        editorForm.retractDelayLabel.Enabled = True
        editorForm.withdrawalSpeed.Enabled = True
        editorForm.UpDownWithDrawalSpeed.Enabled = True
        editorForm.withDrawalSpeedLabel.Enabled = True
        editorForm.WithDrawalZ.Enabled = True
        editorForm.decideWithdrawalHeight.Enabled = True
        editorForm.withDrawalZLabel.Enabled = True
        editorForm.moveHeight.Enabled = True
        editorForm.decideMoveHeight.Enabled = True
        editorForm.moveHeightLabel.Enabled = True
        
    End If
    
    'Repeat node
    If editorForm.NodeType.ListIndex = 22 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        
        editorForm.PathFileName.Enabled = True
        editorForm.pathFileNameLabel.Enabled = True
        editorForm.loadPartArray.Enabled = True
        
        editorForm.xDev.Enabled = True
        editorForm.xDevLabel.Enabled = True
        editorForm.decideXYDev.Enabled = True
        editorForm.yDev.Enabled = True
        editorForm.yDevLabel.Enabled = True
        
        editorForm.xRepeatNum.Enabled = True
        editorForm.UpDownXRepeatNum.Enabled = True
        editorForm.xRepeatNumLabel.Enabled = True
        editorForm.yRepeatNum.Enabled = True
        editorForm.UpDownYRepeatNum.Enabled = True
        editorForm.yRepeatNumLabel.Enabled = True
         
        editorForm.xDev.Text = "10"
        editorForm.yDev.Text = "10"
        editorForm.xRepeatNum.Text = "1"
        editorForm.yRepeatNum = "1"
    End If

End Sub

Public Sub initializeNodeTypeItems()
    Call editorForm.NodeType.AddItem("Reference", 0)
    Call editorForm.NodeType.AddItem("----------------------", 1)
    Call editorForm.NodeType.AddItem("Dot", 2)
    Call editorForm.NodeType.AddItem("Dot Potting", 3)
    Call editorForm.NodeType.AddItem("----------------------", 4)
    Call editorForm.NodeType.AddItem("Line Potting", 5)
    Call editorForm.NodeType.AddItem("Line Start", 6)
    Call editorForm.NodeType.AddItem("Line End", 7)
    Call editorForm.NodeType.AddItem("----------------------", 8)
    Call editorForm.NodeType.AddItem("Arc Start", 9)
    Call editorForm.NodeType.AddItem("Arc Point", 10)
    Call editorForm.NodeType.AddItem("Arc End", 11)
    Call editorForm.NodeType.AddItem("----------------------", 12)
    Call editorForm.NodeType.AddItem("Links Line Point", 13)
    Call editorForm.NodeType.AddItem("Links Arc Restart", 14)
    Call editorForm.NodeType.AddItem("Links Arc Start", 15)
    Call editorForm.NodeType.AddItem("Links Arc End", 16)
    Call editorForm.NodeType.AddItem("----------------------", 17)
    ' Rectangle Node Type   (XW)
    Call editorForm.NodeType.AddItem("RectC1", 18)
    Call editorForm.NodeType.AddItem("RectC2", 19)
    Call editorForm.NodeType.AddItem("RectC3", 20)
    Call editorForm.NodeType.AddItem("----------------------", 21)
    Call editorForm.NodeType.AddItem("Part Array", 22)
    editorForm.NodeType.Selected(0) = True
End Sub

Public Sub drawStatus()
    Dim PW, PH
   
    With executionForm.PictureReady

    .FillStyle = vbFSSolid
    If readyStatus = True Then
        .FillColor = QBColor(10)
    Else
        .FillColor = QBColor(2)
    End If
    PW = .ScaleWidth
    PH = .ScaleHeight
    ' Draw circle
    executionForm.PictureReady.Circle (PW / 2, PH / 2), PH / 3
    End With

    With executionForm.PictureBusy

    .FillStyle = vbFSSolid
    
    If busyStatus = True Then
        .FillColor = QBColor(14)
    Else
        .FillColor = QBColor(6)
    End If
    
    ' Draw circle
    executionForm.PictureBusy.Circle (PW / 2, PH / 2), PH / 3
    End With
    
    With executionForm.PictureError

    .FillStyle = vbFSSolid
    If errorStatus = True Then
        .FillColor = QBColor(12)
        errorStatus = False         'To switch off the red light (XW)
    Else
        .FillColor = QBColor(4)
    End If
    ' Draw circle
    executionForm.PictureError.Circle (PW / 2, PH / 2), PH / 3
    End With
End Sub

Public Sub validateNumber(ByVal str As String, ByVal cap As String)
    If Not IsNumeric(str) Then
        MsgBox ("Please enter a numberic value for " & cap & " !")
        'XW
        ErrorKeyIn = True
    End If
End Sub

Public Function processAddNode() As String
    With editorForm
        Select Case .NodeType.ListIndex
        Case 0
            processAddNode = "reference(x=" & convertToPulses(.dispensePtX.Text, X_axis) & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) & ")"
            referenceX = convertToPulses(.dispensePtX.Text, X_axis)
            referenceY = convertToPulses(.dispensePtY.Text, Y_axis)
            referenceZ = convertToPulses(.dispensePtZ.Text, Z_axis)
        
        Case 2
            If .xRepeatNum = 0 Then
                .xDev.Text = 0
            End If
            If .yRepeatNum = 0 Then
                .yDev.Text = 0
            End If
                
            If ClickExpand = False Then
                If .xRepeatNum > 1 Or .yRepeatNum > 1 Then
                    processAddNode = "dotArray(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(CLng(.WithDrawalZ.Text), Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                Else
                    processAddNode = "dot(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(CLng(.WithDrawalZ.Text), Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                End If
            Else
                'processAddNode = "   dot(x=" & ExpandX + add_column_pitch + referenceX & ", y=" & ExpandY + add_row_pitch + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(CLng(.WithDrawalZ.Text), Z_axis) & "; sp=" & ExpandWithDrawSpeed & "; " & Format(.dispenseTime.Text, "####0.000") & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                processAddNode = "   dot(x=" & ExpandX + add_column_pitch + referenceX & ", y=" & ExpandY + add_row_pitch + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(CLng(.WithDrawalZ.Text), Z_axis) & "; sp=" & ExpandWithDrawSpeed & "; " & Format(.dispenseTime.Text, "####0.000") & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            End If
        Case 3
            If .xRepeatNum = 0 Then
                .xDev.Text = 0
            End If
            If .yRepeatNum = 0 Then
                .yDev.Text = 0
            End If
            
            If ClickExpand = False Then
                If .xRepeatNum > 1 Or .yRepeatNum > 1 Then
                    processAddNode = "dotPottingArray(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                Else
                    processAddNode = "dotPotting(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                End If
            Else
                processAddNode = "   dotPotting(x=" & ExpandX + add_column_pitch + referenceX & ", y=" & ExpandY + add_row_pitch + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            End If
        Case 5
            If .xRepeatNum = 0 Then
                .xDev.Text = 0
            End If
            If .yRepeatNum = 0 Then
                .yDev.Text = 0
            End If
            
            If ClickExpand = False Then
                If .xRepeatNum > 1 Or .yRepeatNum > 1 Then
                    processAddNode = "linePottingArray(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.delay.Text, "####0.000") & "; z=" & convertToPulses(.endDispenseHeight.Text, Z_axis) & "; sp=" & .DispenseSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                Else
                    processAddNode = "linePotting(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.delay.Text, "####0.000") & "; z=" & convertToPulses(.endDispenseHeight.Text, Z_axis) & "; sp=" & .DispenseSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
                End If
            Else
                processAddNode = "   linePotting(x=" & ExpandX + add_column_pitch + referenceX & ", y=" & ExpandY + add_row_pitch + referenceY & ", z=" & ExpandZ + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & 1 & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & 1 & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.delay.Text, "####0.000") & "; z=" & convertToPulses(.endDispenseHeight.Text, Z_axis) & "; sp=" & .DispenseSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            End If
        Case 6
            processAddNode = "lineStart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
        Case 7
            processAddNode = "lineEnd(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
        Case 9
            processAddNode = "arcStart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 10
            processAddNode = "       arcPoint(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 11
            processAddNode = "arcEnd(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            
        Case 13
            processAddNode = "   linksLinePoint(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & ")"
        
        Case 14
            processAddNode = "   linksArcRestart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & ")"
          
        Case 15
            processAddNode = "   linksArcStart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & ")"
            
        Case 16
            processAddNode = "   linksArcEnd(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & ")"
        
        Case 18
            processAddNode = "rectC1(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 19
            processAddNode = "   rectC2(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 20
            processAddNode = "rectC3(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.Value & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            
        Case 22
            processPartArray
            processAddNode = ""
        End Select
    End With
End Function

Public Function doPartArrayY(ByVal x As Long, ByVal y As Long, ByVal Z As Long, ByVal yDev As Long, ByVal yRepeatNum As Long) As Long
    
    Dim yTemp, counter As Long

    yTemp = y

    For counter = 1 To yRepeatNum
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & x + referenceX & ", y=" & yTemp + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            Call editorForm.lstPattern.AddItem("repeat(x=" & x + referenceX & ", y=" & yTemp + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
        End If
        yTemp = yTemp + yDev
    Next counter

    doPartArrayY = yTemp
End Function


Public Function doPartArrayX(ByVal x As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xRepeatNum As Long) As Long
    
    Dim xTemp, counter As Long

    xTemp = x

    For counter = 1 To xRepeatNum
    
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & xTemp + referenceX & ", y=" & y + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            'XW
            If editorForm.lstPattern.SelCount = 1 Then
                Call editorForm.lstPattern.AddItem("repeat(x=" & xTemp + referenceX & ", y=" & y + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
                'Adding "+ 1" is to incerase the index of list box
                editorForm.lstPattern.ListIndex = editorForm.lstPattern.ListIndex + 1
            Else
                editorForm.lstPattern.AddItem ("repeat(x=" & xTemp + referenceX & ", y=" & y + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
            End If
        End If

        xTemp = xTemp + xDev
    Next counter

    doPartArrayX = xTemp
End Function

Public Function doSinglePartArray(ByVal x As Long, ByVal y As Long, ByVal Z As Long) As Long
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & x + referenceX & ", y=" & y + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            Call editorForm.lstPattern.AddItem("repeat(x=" & x + referenceX & ", y=" & y + referenceY & ", z=" & Z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
        End If
        doSinglePartArray = 0
End Function

Public Sub processPartArray()

    Dim xTemp, yTemp, zTemp, xdevtemp, ydevtemp, xcounter As Long
    Dim ycounter As Long    'XW
    
    With editorForm
    
        xTemp = convertToPulses(.dispensePtX.Text, X_axis)
        yTemp = convertToPulses(.dispensePtY.Text, Y_axis)
        zTemp = convertToPulses(.dispensePtZ.Text, Z_axis)
        ydevtemp = convertToPulses(.yDev.Text, Y_axis) - yTemp
        xdevtemp = convertToPulses(.xDev.Text, X_axis) - xTemp
        
        If (xdevtemp = 0 And ydevtemp = 0) Then
            xTemp = doSinglePartArray(xTemp, yTemp, zTemp)
        ElseIf (ydevtemp = 0) Then
            xTemp = doPartArrayX(xTemp, yTemp, zTemp, xdevtemp, .xRepeatNum.Text)
        ElseIf (xdevtemp = 0) Then
            yTemp = doPartArrayY(xTemp, yTemp, zTemp, ydevtemp, .yRepeatNum.Text)
        Else
            'For xcounter = 1 To .xRepeatNum.Text
            '    yTemp = doPartArrayY(xTemp, yTemp, zTemp, ydevtemp, .yRepeatNum.Text)
            '    ydevtemp = -ydevtemp
            '    yTemp = yTemp + ydevtemp
            '    xTemp = xTemp + xdevtemp
            'Next xcounter
                
            'To be matched with other dotting array-pattern     'XW
            For ycounter = 1 To .yRepeatNum.Text
                xTemp = doPartArrayX(xTemp, yTemp, zTemp, xdevtemp, .xRepeatNum.Text)
                xdevtemp = -xdevtemp
                xTemp = xTemp + xdevtemp
                yTemp = yTemp + ydevtemp
            Next ycounter
        End If
    End With


End Sub


Public Sub doTrack(ByVal patternStr As String)

    Dim errorString As String
    Dim Response As GPMessageConstants
    Dim Parser   As New GOLDParser
    Dim Done, error As Boolean                                    'Controls when we leave the loop
   
    Dim ReductionNumber As Integer                         'Just for information
    Dim n As Integer, Text As String
            
    If Parser.LoadCompiledGrammar(txtCGTFilePath1) Then
        Parser.OpenTextString (patternStr)
        Parser.TrimReductions = True
                        
        Done = False
        error = False
        Do Until Done
            Response = Parser.Parse()
              
            Select Case Response
            Case gpMsgLexicalError
                errorString = "Line " & Parser.CurrentLineNumber & ": Lexical Error: Cannot recognize token: " & Parser.CurrentToken.Data
                MsgBox (errorString)
                Done = True
                error = True
            Case gpMsgSyntaxError
                Text = ""
                For n = 0 To Parser.TokenCount - 1
                    Text = Text & " " & Parser.Tokens(n).Name
                Next
                errorString = "Line " & Parser.CurrentLineNumber & ": Syntax Error: Expecting the following tokens: " & LTrim(Text)
                MsgBox (errorString)
                Done = True
                error = True
              
            Case gpMsgReduction
                ReductionNumber = ReductionNumber + 1
                Parser.CurrentReduction.Tag = ReductionNumber   'Mark the reduction
              
            Case gpMsgAccept
                '=== Success!
                Call processSelectedNode(Parser.CurrentReduction)
                Done = True
              
            Case gpMsgTokenRead
              
            Case gpMsgInternalError
                errorString = "Internal Error, " & "Something is horribly wrong, " & Parser.CurrentLineNumber
                MsgBox (errorString)
                Done = True
                error = True
              
            Case gpMsgNotLoadedError
                '=== Due to the if-statement above, this case statement should never be true
                errorString = "Not Loaded Error, Compiled Gramar Table not loaded"
                MsgBox (errorString)
                Done = True
              
            Case gpMsgCommentError
                errorString = "Comment Error, Unexpected end of file at line number: " & Parser.CurrentLineNumber
                MsgBox (errorString)
                Done = True
                error = True
            End Select
           
        Loop
    Else
        MsgBox "Could not load the CGT file", vbCritical
        error = True
    End If
                
End Sub

Private Sub processSelectedNode(TheReduction As Reduction)

Dim n As Integer

For n = 0 To TheReduction.TokenCount - 1
    Select Case TheReduction.Tokens(n).Kind
        Case SymbolTypeNonterminal
            Call processSelectedNode(TheReduction.Tokens(n).Data)
        Case Else
            With editorForm
            Select Case LCase(TheReduction.Tokens(n).Data)
                Case "reference"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    referenceX = CLng(.dispensePtX.Text)
                    referenceY = CLng(.dispensePtY.Text)
                    referenceZ = CLng(.dispensePtZ.Text)
                    referenceSet = True
                Case "repeat"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .PathFileName.Text = TheReduction.Tokens(6).Data.Tokens(0).Data
                    .PathFileName.Text = Left(.PathFileName.Text, Len(.PathFileName.Text) - 1)
                    .PathFileName.Text = Right(.PathFileName.Text, Len(.PathFileName.Text) - 1)
                Case "linksarcrestart"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(4).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(6).Data.Tokens(0).Data
                Case "arc"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                Case "start"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .delay.Text = TheReduction.Tokens(6).Data.Tokens(0).Data
                Case "arcstart"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .delay.Text = TheReduction.Tokens(6).Data.Tokens(0).Data
                Case "line3d"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(6).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(8).Data.Tokens(0).Data
                Case "end3d"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(6).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(8).Data.Tokens(0).Data
                    .retractDelay.Text = TheReduction.Tokens(10).Data.Tokens(0).Data
                    .WithDrawalZ.Text = convertToMM(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .withdrawalSpeed.Text = TheReduction.Tokens(14).Data.Tokens(2).Data
                    .moveHeight.Text = convertToMM(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                Case "arcend"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(6).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(8).Data.Tokens(0).Data
                    .retractDelay.Text = TheReduction.Tokens(10).Data.Tokens(0).Data
                    .WithDrawalZ.Text = convertToMM(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .withdrawalSpeed.Text = TheReduction.Tokens(14).Data.Tokens(2).Data
                    .moveHeight.Text = convertToMM(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                Case "linksarcstart"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(6).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(8).Data.Tokens(0).Data
                Case "linksarcend"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    .DispenseSpeed.Text = TheReduction.Tokens(6).Data.Tokens(2).Data
                    .dispenseOnOff.Value = TheReduction.Tokens(8).Data.Tokens(0).Data
                Case "dot"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandX = .dispensePtX.Text
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    ExpandY = .dispensePtY.Text
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandZ = .dispensePtZ.Text
                    .xDev.Text = convertToMM(TheReduction.Tokens(6).Data.Tokens(0).Data, X_axis)
                    .xRepeatNum.Text = TheReduction.Tokens(8).Data.Tokens(0).Data
                    .yDev.Text = convertToMM(TheReduction.Tokens(10).Data.Tokens(0).Data, Y_axis)
                    .yRepeatNum.Text = TheReduction.Tokens(12).Data.Tokens(0).Data
                    .WithDrawalZ.Text = convertToMM(TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .DispenseSpeed.Text = TheReduction.Tokens(16).Data.Tokens(2).Data
                    .dispenseTime.Text = TheReduction.Tokens(18).Data.Tokens(0).Data
                    .retractDelay.Text = TheReduction.Tokens(20).Data.Tokens(0).Data
                    .moveHeight.Text = convertToMM(TheReduction.Tokens(22).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                Case "pottype1"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandX = .dispensePtX.Text
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    ExpandY = .dispensePtY.Text
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandZ = .dispensePtZ.Text
                    .xDev.Text = convertToMM(TheReduction.Tokens(6).Data.Tokens(0).Data, X_axis)
                    .xRepeatNum.Text = TheReduction.Tokens(8).Data.Tokens(0).Data
                    .yDev.Text = convertToMM(TheReduction.Tokens(10).Data.Tokens(0).Data, Y_axis)
                    .yRepeatNum.Text = TheReduction.Tokens(12).Data.Tokens(0).Data
                    .potDepth.Text = convertToMM(TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .depthSpeed.Text = TheReduction.Tokens(16).Data.Tokens(2).Data
                    .dispenseTime.Text = TheReduction.Tokens(18).Data.Tokens(0).Data
                    .WithDrawalZ.Text = convertToMM(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .withdrawalSpeed.Text = TheReduction.Tokens(22).Data.Tokens(2).Data
                    .retractDelay.Text = TheReduction.Tokens(24).Data.Tokens(0).Data
                    .moveHeight.Text = convertToMM(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                Case "pottype2"
                    .dispensePtX.Text = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandX = .dispensePtX.Text
                    .dispensePtY.Text = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    ExpandY = .dispensePtY.Text
                    .dispensePtZ.Text = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    ExpandZ = .dispensePtZ.Text
                    .xDev.Text = convertToMM(TheReduction.Tokens(6).Data.Tokens(0).Data, X_axis)
                    .xRepeatNum.Text = TheReduction.Tokens(8).Data.Tokens(0).Data
                    .yDev.Text = convertToMM(TheReduction.Tokens(10).Data.Tokens(0).Data, Y_axis)
                    .yRepeatNum.Text = TheReduction.Tokens(12).Data.Tokens(0).Data
                    .potDepth.Text = convertToMM(TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .depthSpeed.Text = TheReduction.Tokens(16).Data.Tokens(2).Data
                    .dispenseTime.Text = TheReduction.Tokens(18).Data.Tokens(0).Data
                    .endDispenseHeight.Text = convertToMM(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .DispenseSpeed.Text = TheReduction.Tokens(22).Data.Tokens(2).Data
                    .retractDelay.Text = TheReduction.Tokens(24).Data.Tokens(0).Data
                    .WithDrawalZ.Text = convertToMM(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
                    .withdrawalSpeed.Text = TheReduction.Tokens(28).Data.Tokens(2).Data
                    .moveHeight.Text = convertToMM(TheReduction.Tokens(30).Data.Tokens(2).Data.Tokens(0).Data, Z_axis)
            End Select
            End With
    End Select
Next
End Sub

Public Function convertToMM(ByVal pulseCount As Long, ByVal axis As Long) As Double

    Dim factor As Long

    If (axis = Z_axis) Then
        factor = ZGearRatio
    Else
        factor = XYGearRatio
    End If

    convertToMM = Round(pulseCount / factor, 3)

End Function


Public Function convertToPulses(ByVal measurement As Double, ByVal axis As Long) As Long

    Dim factor As Long

    If (axis = Z_axis) Then
        factor = ZGearRatio
    Else
        factor = XYGearRatio
    End If

    convertToPulses = measurement * factor


End Function

'''''''''''''''''''''''''''''
'   Testing two valve (XW)  '
'''''''''''''''''''''''''''''
Public Sub LeftNeedleValve()
    'Change to left-valve
    Dim ReadValue As Long
    Dim L_Down_Sensor As Byte
    
    'Left-needle will be gone down
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.5)
    
     'Left_Down_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Down_Sensor)
    L_Down_Sensor = L_Down_Sensor And &H4
    
    If (L_Down_Sensor <> 0) And (L_Down_Sensor <> 4) Then
        MsgBox "Left Cylinder have some problem!"
    End If
End Sub

Public Sub RightNeedleValve()
    'Change to right-valve
    
    Dim ReadValue As Long
    Dim L_Up_Sensor As Byte
    
    'Left-needle will be gone up.
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.5)
    
    'Left_Up_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Up_Sensor)
    L_Up_Sensor = L_Up_Sensor And &H1
    
    If Not (L_Up_Sensor = 0) Then
        MsgBox "Left Cylinder have some problem!"
    End If
End Sub

Public Sub Sleep(ByVal DelayTime As Double)
    '''''''''''''''''
    '   Do delay    '
    '''''''''''''''''
    Dim CurrentTime
    
    CurrentTime = Timer
    
    Do While (Timer < CurrentTime + DelayTime)
        If (Timer < CurrentTime) Then
            CurrentTime = (86400 - CurrentTime)
        End If
        
        DoEvents
    Loop
End Sub

Public Sub Open_Valve1()
    Dim ReadValue As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    'Wait for a few (m sec)
    Call Sleep(0.5)
End Sub

Public Sub Open_Valve2()
    Dim ReadValue As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H100
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    'Wait for a few (m sec)
    Call Sleep(0.5)
End Sub

Public Sub Close_Valve1()
    Dim ReadValue As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HF7FF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    'Wait for a few (m sec)
    Call Sleep(0.5)
End Sub

Public Sub Close_Valve2()
    Dim ReadValve As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HFEFF
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    'Wait for a few (m sec)
    Call Sleep(0.5)
End Sub

Public Sub Tilt_ON()
    Dim ReadValve As Long, ReadValue2 As Byte
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H100
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.5)
    
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, ReadValue2)
    ReadValue2 = ReadValue2 And &H2
    
    If (ReadValue2 = &H2) Then
        MsgBox "Tilting has problem, please check hardware!"
    End If
End Sub

Public Sub Tilt_Off()
    Dim ReadValve As Long, ReadValue2 As Byte
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, ReadValue))
    ReadValue = ReadValue And &HFEFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.5)
    
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, ReadValue2)
    ReadValue2 = ReadValue2 And &H4
    
    If (ReadValue2 = &H4) Then
        MsgBox "Tilting has problem, please check hardware!"
    End If
End Sub

Public Function Turnning_Angle(ByVal Rot_angle As String)
    Dim String_Line As String
    
    If (RepeatPattern = True) Then
'        If (Rot_angle = "0") Or (Rot_angle = "-360") Then
'            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=36363; 55555.000; 55555.000; z=55555)"
'        ElseIf (Rot_angle = "-90") Then
'            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=99999; 55555.000; 55555.000; z=55555)"
'        ElseIf (Rot_angle = "-180") Then
'            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=18181; 55555.000; 55555.000; z=55555)"
'        ElseIf (Rot_angle = "-270") Then
'            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=27272; 55555.000; 55555.000; z=55555)"
'        ElseIf (Rot_angle = "None") Then
'            'No tilt
'            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)"
'        End If
        
        '@$K
        If (Rot_angle = "None") Then
            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)"
        Else
            String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp= " & CLng(Rot_angle) & "; 55555.000; 55555.000; z=55555)"
        End If
        
        String_Line = String_Line & vbNewLine
        ReadRepeatString = ReadRepeatString & String_Line
        RepeatPattern = False
    Else
'        If (Rot_angle = "0") Or (Rot_angle = "-360") Then
'            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=36363; 55555.000; 55555.000; z=55555)")
'        ElseIf (Rot_angle = "-90") Then
'            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=99999; 55555.000; 55555.000; z=55555)")
'        ElseIf (Rot_angle = "-180") Then
'            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=18181; 55555.000; 55555.000; z=55555)")
'        ElseIf (Rot_angle = "-270") Then
'            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=27272; 55555.000; 55555.000; z=55555)")
'        ElseIf (Rot_angle = "None") Then
'            'No tilt
'            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)")
'        End If

        '@$K
        If (Rot_angle = "None") Then
            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)")
        Else
            A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp= " & CLng(Rot_angle) & "; 55555.000; 55555.000; z=55555)")
        End If
    End If
End Function

Public Function Turnning_Line_Angle(ByVal Line_Angle As String)
    Dim String_Line As String
    
    If (RepeatPattern = True) Then
'        If (Line_Angle = "0") Or (Line_Angle = "-360") Then
'            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)"
'        ElseIf (Line_Angle = "-90") Then
'            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)"
'        ElseIf (Line_Angle = "-180") Then
'            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)"
'        ElseIf (Line_Angle = "-270") Then
'            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)"
'        ElseIf (Line_Angle = "None") Then
'            'No tilt
'            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=10101; 1)"
'        End If

        '@$K
        If (Line_Angle = "None") Then
            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=10101; 1)"
        Else
            String_Line = "line3D(x=-55555, y=-55555, z=-55555; sp=" & CLng(Line_Angle) & "; 1)"
        End If
        
        String_Line = String_Line & vbNewLine
        ReadRepeatString = ReadRepeatString & String_Line
        RepeatPattern = False
    Else
'        If (Line_Angle = "0") Or (Line_Angle = "-360") Then
'            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)")
'        ElseIf (Line_Angle = "-90") Then
'            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)")
'        ElseIf (Line_Angle = "-180") Then
'            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)")
'        ElseIf (Line_Angle = "-270") Then
'            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)")
'        ElseIf (Line_Angle = "None") Then
'            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=10101; 1)")
'        End If

        '@$K
        If (Line_Angle = "None") Then
            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=10101; 1)")
        Else
            A.writeline ("line3D(x=-55555, y=-55555, z=-55555; sp=" & CLng(Line_Angle) & "; 1)")
        End If
    End If
End Function

Public Function Turnning_Arc_Angle(ByVal x1 As Double, ByVal x2 As Double, ByVal x3 As Double, ByVal y1 As Double, ByVal y2 As Double, ByVal y3 As Double, ByVal z1 As Double, ByVal z2 As Double, ByVal z3 As Double, ByVal Arc_St As Boolean, ByVal Arc_End As String, ByVal Rot_angle As String) As Boolean
    Dim String_Line As String
    
    If (x1 = x2) And (x1 = x3) And (y1 = y2) And (y1 = y3) Then
        Turnning_Arc_Angle = True
        Exit Function
    End If
    
    If (x1 < x2) And (y1 > y2) And (x2 < x3) And (y2 < y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                'If the previous node is no "Tilting", there will not do the rotating although the system is running the spray valve
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=27272; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-270")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-270")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("0")
            End If
            A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-90")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 > x2) And (y1 > y2) And (x2 > x3) And (y2 < y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=99999; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-90")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-90")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("0")
            End If
            A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-270")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 < x2) And (y1 < y2) And (x2 < x3) And (y2 > y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=27272; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-270")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-270")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-180")
            End If
            A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-90")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 > x2) And (y1 < y2) And (x2 > x3) And (y2 > y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=99999; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-90")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-90")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-180")
            End If
            A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-270")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 > x2) And (y1 < y2) And (x2 < x3) And (y2 < y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=36363; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("0")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("0")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-270")
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-180")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 < x2) And (y1 < y2) And (x2 > x3) And (y2 < y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=36363; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("0")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("0")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-90")
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-180")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 > x2) And (y1 > y2) And (x2 < x3) And (y2 > y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=18181; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=27272; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-180")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-180")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-270")
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("0")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    ElseIf (x1 < x2) And (y1 > y2) And (x2 > x3) And (y2 > y3) Then
        If (RepeatPattern = True) Then
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    String_Line = "dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=18181; 55555.000; 55555.000; z=55555)" & vbNewLine
                End If
                String_Line = String_Line & "start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")" & vbNewLine
            Else
                String_Line = "line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                If (Rot_angle <> "None") Then
                    String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=18181; 1)" & vbNewLine
                End If
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=99999; 1)" & vbNewLine
            End If
            String_Line = String_Line & "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
            If (Rot_angle <> "None") Then
                String_Line = String_Line & "line3D(x=-55555, y=-55555, z=-55555; sp=36363; 1)" & vbNewLine
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    StringLine = "end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")" & vbNewLine
                Else
                    String_Line = String_Line & "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")" & vbNewLine
                End If
            End If
            ReadRepeatString = ReadRepeatString & String_Line
            RepeatPattern = False
        Else
            If (Arc_St = True) Then
                If (Rot_angle <> "None") Then
                    Turnning_Angle ("-180")
                End If
                A.writeline ("start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
            Else
                A.writeline ("line3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & ")")
                If (Rot_angle <> "None") Then
                    Turnning_Line_Angle ("-180")
                End If
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("-90")
            End If
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z2 * 1000) & "; " & TravelSpeed & ")")
            If (Rot_angle <> "None") Then
                Turnning_Line_Angle ("0")
            End If
            If (NoChange3 = False) Then
                If (Arc_End = "arcEnd") Then
                    A.writeline ("end3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                Else
                    A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; " & TravelSpeed & ")")
                End If
            End If
        End If
    End If
    
    Turnning_Arc_Angle = False
End Function

Public Sub Tower_Light()
    Dim Tower_Light_Value As Long, Light_Value As Long      'Use for tower's lights
    
    If (Red_Light = True) Then
        'Indicate Red_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value Or &H800
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    Else
        'Disable Red_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value And &HF7FF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    End If
    
    If (Yellow_Light = True) Then
        'Indicate Yellow_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value Or &H400
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    Else
        'Disable Yellow_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value And &HFBFF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    End If
    
    If (Green_Light = True) Then
        'Indicate Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value Or &H200
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    Else
        'Disable Green_Light
        checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
        Light_Value = Tower_Light_Value And &HFDFF
        checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Light_Value))
    End If
End Sub

Public Sub Close_TowerLight()
    Dim Tower_Light_Value As Long
    
    'Disable Red_Light,Yellow_light and Green_Light
    checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, Tower_Light_Value))
    Tower_Light_Value = Tower_Light_Value And &HF1FF
    checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, Tower_Light_Value))
                
    Call Sleep(0.03)
End Sub

'''''''''''''''''''''''''''''
'   both slinder go down    '
'''''''''''''''''''''''''''''
Public Sub Both_Go_Down()
    Dim ReadValue As Long
    
    'Left
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    'Sleep (0.5)
    
'    'Right
'    checkSuccess (P1240MotRdReg(boardNum, U_axis, WR3, ReadValue))
'    ReadValue = ReadValue Or &H100
'
'    checkSuccess (P1240MotWrReg(boardNum, U_axis, WR3, ReadValue))
End Sub

Public Sub Move_To_Zero()
    setSpeed (xySystemTravelSpeed)
        
    checkSuccess (P1240MotPtp(boardNum, Z_axis, Z_axis, 0, 0, 0, 0))
    Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success)
        DoEvents
    Loop
End Sub

Public Sub ResetDriver()
    Dim resetValue As Long
    
    checkSuccess (P1240MotRdReg(boardNum, Y_axis, WR3, resetValue))
    resetValue = resetValue Or &H200
    
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
    
    Call Sleep(0.3)
        
    resetValue = resetValue And &HFDFF
    checkSuccess (P1240MotWrReg(boardNum, Y_axis, WR3, resetValue))
  
    Call Sleep(0.5)
End Sub

Public Sub SetWindowOnTop(f As Form, bAlwaysOnTop As Boolean)   '@$K
   Dim iFlag As Long
   iFlag = IIf(bAlwaysOnTop, HWND_TOPMOST, HWND_NOTOPMOST)
   SetWindowPos f.hWnd, iFlag, f.Left / Screen.TwipsPerPixelX, f.Top / Screen.TwipsPerPixelY, f.Width / Screen.TwipsPerPixelX, f.height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW
End Sub

