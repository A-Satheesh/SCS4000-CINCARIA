Attribute VB_Name = "guiRoutines"
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40

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
    editorForm.dispenseOnOff.value = Checked
    editorForm.retractDelay.Text = "1.0"
    editorForm.withdrawalSpeed.Text = "10"
    editorForm.WithDrawalZ.Text = "0"
    editorForm.moveHeight.Text = convertToMM(systemMoveHeight, Z_axis)
    editorForm.xRepeatNum.Text = "1"
    editorForm.yRepeatNum.Text = "1"
    editorForm.xDev.Text = "10"
    editorForm.yDev.Text = "10"
    editorForm.PathFileName.Text = ""
    editorForm.jogSpeedSlider.value = 28
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

    'Pot type1 node
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
    
    'Pot type1 node
    If editorForm.NodeType.ListIndex = 4 Then
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
    
    'Start node
    If editorForm.NodeType.ListIndex = 6 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
        
    'Arc Start Node
    If editorForm.NodeType.ListIndex = 7 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.delay.Enabled = True
        editorForm.delayLabel.Enabled = True
        editorForm.UpDownDelay.Enabled = True
    End If
    
    'Arc node no need to re-enable!
    
    'Line node
    If editorForm.NodeType.ListIndex = 10 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
        
    'Links Arc Start node
    If editorForm.NodeType.ListIndex = 11 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
        
    'Links Arc Restart node
    If editorForm.NodeType.ListIndex = 12 Then
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
        
    'Links Arc End node
    If editorForm.NodeType.ListIndex = 13 Then
        editorForm.dispensePtZ.Enabled = True
        editorForm.dispensePtZLabel.Enabled = True
        editorForm.DispenseSpeed.Enabled = True
        editorForm.UpDownDispenseSpeed.Enabled = True
        editorForm.dispenseSpeedLabel.Enabled = True
        editorForm.dispenseOnOff.Enabled = True
    End If
           
    'End node
    If editorForm.NodeType.ListIndex = 15 Then
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
    
    'Arc end node
    If editorForm.NodeType.ListIndex = 16 Then
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
    If editorForm.NodeType.ListIndex = 18 Then
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
    Call editorForm.NodeType.AddItem("Pot Dot", 3)
    Call editorForm.NodeType.AddItem("Pot Line", 4)
    Call editorForm.NodeType.AddItem("----------------------", 5)
    Call editorForm.NodeType.AddItem("Start Line", 6)
    Call editorForm.NodeType.AddItem("Start Arc", 7)
    Call editorForm.NodeType.AddItem("----------------------", 8)
    Call editorForm.NodeType.AddItem("Arc Point", 9)
    Call editorForm.NodeType.AddItem("Links Line Point", 10)
    Call editorForm.NodeType.AddItem("Links Arc Start", 11)
    Call editorForm.NodeType.AddItem("Links Arc Restart", 12)
    Call editorForm.NodeType.AddItem("Links Arc End Node", 13)
    Call editorForm.NodeType.AddItem("----------------------", 14)
    Call editorForm.NodeType.AddItem("End Line", 15)
    Call editorForm.NodeType.AddItem("End Arc", 16)
    Call editorForm.NodeType.AddItem("----------------------", 17)
    Call editorForm.NodeType.AddItem("Part Array", 18)
    editorForm.NodeType.Selected(0) = True
End Sub

Public Sub drawStatus()
    Dim PW, PH
   
    With editorForm.PictureReady

    .FillStyle = vbFSSolid
    If readyStatus = True Then
        .FillColor = QBColor(10)
    Else
        .FillColor = QBColor(2)
    End If
    PW = .ScaleWidth
    PH = .ScaleHeight
    ' Draw circle
    editorForm.PictureReady.Circle (PW / 2, PH / 2), PH / 3
    End With

    With editorForm.PictureBusy

    .FillStyle = vbFSSolid
    
    If busyStatus = True Then
        .FillColor = QBColor(14)
    Else
        .FillColor = QBColor(6)
    End If
    
    ' Draw circle
    editorForm.PictureBusy.Circle (PW / 2, PH / 2), PH / 3
    End With
    
    With editorForm.PictureError

    .FillStyle = vbFSSolid
    If errorStatus = True Then
        .FillColor = QBColor(12)
    Else
        .FillColor = QBColor(4)
    End If
    
    ' Draw circle
    editorForm.PictureError.Circle (PW / 2, PH / 2), PH / 3
    End With
End Sub

Public Sub validateNumber(ByVal str As String, ByVal cap As String)
    If Not IsNumeric(str) Then
        MsgBox ("Please enter a numberic value for " & cap & " !")
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
            
            processAddNode = "dot(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(CLng(.WithDrawalZ.Text), Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
        Case 3
            If .xRepeatNum = 0 Then
                .xDev.Text = 0
            End If
            If .yRepeatNum = 0 Then
                .yDev.Text = 0
            End If
            
            processAddNode = "potType1(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.dispenseTime.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            
        Case 4
            If .xRepeatNum = 0 Then
                .xDev.Text = 0
            End If
            If .yRepeatNum = 0 Then
                .yDev.Text = 0
            End If
            
            processAddNode = "potType2(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & convertToPulses(.xDev.Text, X_axis) & ", " & .xRepeatNum.Text & "; " & convertToPulses(.yDev.Text, Y_axis) & ", " & .yRepeatNum.Text & "; z=" & convertToPulses(.potDepth.Text, Z_axis) & "; sp=" & .depthSpeed.Text & "; " & Format(.delay.Text, "####0.000") & "; z=" & convertToPulses(.endDispenseHeight.Text, Z_axis) & "; sp=" & .DispenseSpeed.Text & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
            
        Case 6
            processAddNode = "start(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 7
            processAddNode = "arcStart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; " & Format(.delay.Text, "####0.000") & ")"
            
        Case 9
            processAddNode = "       arc(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ")"
        Case 10
            processAddNode = "   line3D(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & ")"
        Case 11
            processAddNode = "   linksArcStart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & ")"
            
        Case 12
            processAddNode = "   linksArcRestart(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & ")"
            
        Case 13
            processAddNode = "   linksArcEnd(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & ")"
            
        Case 15
            processAddNode = "end3D(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
        
        Case 16
            processAddNode = "arcEnd(x=" & convertToPulses(.dispensePtX.Text, X_axis) + referenceX & ", y=" & convertToPulses(.dispensePtY.Text, Y_axis) + referenceY & ", z=" & convertToPulses(.dispensePtZ.Text, Z_axis) + referenceZ & "; sp=" & .DispenseSpeed.Text & "; " & .dispenseOnOff.value & "; " & Format(.retractDelay.Text, "####0.000") & "; z=" & convertToPulses(.WithDrawalZ.Text, Z_axis) & "; sp=" & .withdrawalSpeed.Text & "; z=" & convertToPulses(.moveHeight.Text, Z_axis) & ")"
        
        Case 18
            processPartArray
            processAddNode = ""
        
        End Select
    End With
End Function

Public Function doPartArrayY(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal yDev As Long, ByVal yRepeatNum As Long) As Long
    
    Dim yTemp, counter As Long

    yTemp = Y

    For counter = 1 To yRepeatNum
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & X + referenceX & ", y=" & yTemp + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            Call editorForm.lstPattern.AddItem("repeat(x=" & X + referenceX & ", y=" & yTemp + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
        End If
        yTemp = yTemp + yDev
    Next counter

    doPartArrayY = yTemp
End Function


Public Function doPartArrayX(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xRepeatNum As Long) As Long
    
    Dim xTemp, counter As Long

    xTemp = X

    For counter = 1 To xRepeatNum
    
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & xTemp + referenceX & ", y=" & Y + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            Call editorForm.lstPattern.AddItem("repeat(x=" & xTemp + referenceX & ", y=" & Y + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
        End If

        xTemp = xTemp + xDev
    Next counter

    doPartArrayX = xTemp
End Function

Public Function doSinglePartArray(ByVal X As Long, ByVal Y As Long, ByVal z As Long) As Long
        If (editorForm.lstPattern.ListIndex = -1) Then
            editorForm.lstPattern.AddItem ("repeat(x=" & X + referenceX & ", y=" & Y + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")")
        Else
            Call editorForm.lstPattern.AddItem("repeat(x=" & X + referenceX & ", y=" & Y + referenceY & ", z=" & z + referenceZ & "; " & Chr(34) & editorForm.PathFileName.Text & Chr(34) & ")", editorForm.lstPattern.ListIndex)
        End If
        doSinglePartArray = 0
End Function

Public Sub processPartArray()

    Dim xTemp, yTemp, zTemp, xdevtemp, ydevtemp, xcounter As Long

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
            For xcounter = 1 To .xRepeatNum.Text
                yTemp = doPartArrayY(xTemp, yTemp, zTemp, ydevtemp, .yRepeatNum.Text)
                ydevtemp = -ydevtemp
                yTemp = yTemp + ydevtemp
                xTemp = xTemp + xdevtemp
            Next xcounter
        End If
    End With
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

Public Sub LeftNeedleValve()
    'Change to left-valve
    
    Dim ReadValue As Long
    Dim L_Down_Sensor As Byte
    
    'Left-needle will be gone down
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(0.8)
    
    'Left_Down_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Down_Sensor)
    L_Down_Sensor = L_Down_Sensor And &H4
    
    If Not (L_Down_Sensor = 0) Then
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
    Call Sleep(0.8)
    
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
'   Left slinder go down    '
'''''''''''''''''''''''''''''
Public Sub Leftslider_go_down()
    Dim ReadValue As Long
    Dim L_Down_Sensor As Byte
    
    'Left slider
    checkSuccess (P1240MotRdReg(boardNum, Z_axis, WR3, ReadValue))
    ReadValue = ReadValue Or &H800
    
    checkSuccess (P1240MotWrReg(boardNum, Z_axis, WR3, ReadValue))
    
    'Wait for a few (m sec)
    Call Sleep(1)
    
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
    Call Sleep(1)
    
    'Left_Up_Sensor input
    Call AdxDioReadDiPorts(m_lDevHandle, nPortStart, 1, L_Up_Sensor)
    L_Up_Sensor = L_Up_Sensor And &H1
    
    If Not (L_Up_Sensor = 0) Then
        MsgBox "Left Cylinder have some problem!"
    End If
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
   SetWindowPos f.hWnd, iFlag, f.Left / Screen.TwipsPerPixelX, f.Top / Screen.TwipsPerPixelY, f.Width / Screen.TwipsPerPixelX, f.Height / Screen.TwipsPerPixelY, SWP_SHOWWINDOW
End Sub
