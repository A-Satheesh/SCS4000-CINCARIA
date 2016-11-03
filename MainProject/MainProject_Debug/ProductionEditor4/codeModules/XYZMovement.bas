Attribute VB_Name = "XYZMovement"
Public Sub Move_xPlus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, X_axis, 1))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, X_axis, 0, -convertToPulses(CDbl(editorForm.StepDistance.Text), X_axis), 0, 0, 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, X_axis, 0, -convertToPulses(CDbl(frmDistance.StepDistance.Text), X_axis), 0, 0, 0))
        End If
        
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_yPlus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 0))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, Y_axis, 0, 0, convertToPulses(CDbl(editorForm.StepDistance.Text), Y_axis), 0, 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, Y_axis, 0, 0, convertToPulses(CDbl(frmDistance.StepDistance.Text), Y_axis), 0, 0))
        End If
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success) 'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_zPlus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 0))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, Z_axis, 0, 0, 0, convertToPulses(CDbl(editorForm.StepDistance.Text), Z_axis), 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, Z_axis, 0, 0, 0, convertToPulses(CDbl(frmDistance.StepDistance.Text), Z_axis), 0))
        End If
        
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_xMinus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, X_axis, 0))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, X_axis, 0, convertToPulses(CDbl(editorForm.StepDistance.Text), X_axis), 0, 0, 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, X_axis, 0, convertToPulses(CDbl(frmDistance.StepDistance.Text), X_axis), 0, 0, 0))
        End If
        Do While (P1240MotAxisBusy(boardNum, X_axis) <> Success)  'Loop while X motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_yMinus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, Y_axis, 2))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, Y_axis, 0, 0, -convertToPulses(CDbl(editorForm.StepDistance.Text), Y_axis), 0, 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, Y_axis, 0, 0, -convertToPulses(CDbl(frmDistance.StepDistance.Text), Y_axis), 0, 0))
        End If
        
        Do While (P1240MotAxisBusy(boardNum, Y_axis) <> Success)  'Loop while Y motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_zMinus_mouseDown()
    If (editorForm.Jogging.value = True) Or (frmDistance.Jogging.value = True) Then
        checkSuccess (P1240MotCmove(boardNum, Z_axis, 4))
    ElseIf (editorForm.JoggingStep.value = True) Or (frmDistance.JoggingStep.value = True) Then
        If (frmDistance.JoggingStep.value = False) Then
            checkSuccess (P1240MotPtp(boardNum, Z_axis, 0, 0, 0, -convertToPulses(CDbl(editorForm.StepDistance.Text), Z_axis), 0))
        Else
            checkSuccess (P1240MotPtp(boardNum, Z_axis, 0, 0, 0, -convertToPulses(CDbl(frmDistance.StepDistance.Text), Z_axis), 0))
        End If
        
        Do While (P1240MotAxisBusy(boardNum, Z_axis) <> Success) 'Loop while Z motor is still spinning
        Loop
    End If
End Sub

Public Sub Move_xPlusMinus_mouseUp()
    checkSuccess (P1240MotStop(boardNum, X_axis, 1))
End Sub

Public Sub Move_yPlusMinus_mouseUp()
    checkSuccess (P1240MotStop(boardNum, Y_axis, 2))
End Sub

Public Sub Move_zPlusMinus_mouseUp()
    checkSuccess (P1240MotStop(boardNum, Z_axis, 4))
End Sub

