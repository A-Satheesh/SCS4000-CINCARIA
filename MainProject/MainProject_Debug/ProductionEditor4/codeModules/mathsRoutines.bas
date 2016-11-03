Attribute VB_Name = "mathsRoutines"
Public Function DoArcCalCen(ByVal X As Double, ByVal Y As Double) As Boolean

    Dim var1, var2, var3, var4, var5, var6, xc, yc, temp As Double

    'Delibrately intro slight error to compensate for arc in a straight line

    If PrevPrevX = PrevX Then
        PrevX = PrevX + 1
    End If

    If PrevPrevY = PrevY Then
        PrevY = PrevY + 1
    End If

    var1 = (PrevPrevX * PrevPrevX) - (PrevX * PrevX) + (PrevPrevY * PrevPrevY) - (PrevY * PrevY)
    var2 = (PrevX * PrevX) - (X * X) + (PrevY * PrevY) - (Y * Y)
    var3 = PrevPrevY - PrevY
    var4 = PrevX - X
    var5 = PrevPrevX - PrevX
    var6 = PrevY - Y

    If ((var4 = 0) Or (var5 = 0)) Then GoTo error

    yc = var1 / (2 * var5)
    yc = yc - (var2 / (2 * var4))
    temp = var3 / var5
    temp = temp - (var6 / var4)

error:

    If (temp = 0) Then
        DoArcCalCen = False
        Exit Function
    Else
        DoArcCalCen = True
    End If

    yc = yc / temp

    yCen = CLng(yc)

    xCen = CLng((var1 - (2 * var3 * yc)) / (2 * var5))

End Function

Public Function detAngle(ByVal X As Double, ByVal Y As Double) As Double

    Dim sinevalue, Result, radius As Double
        
        
    radius = Sqr((X * X) + (Y * Y))
        
    sinevalue = Y / radius
    
    If (sinevalue <> 1 And sinevalue <> -1) Then
        
        Result = Atn(sinevalue / Sqr(-sinevalue * sinevalue + 1))
    
        Result = Result * 180 / (Atn(1) * 4)
    
        If X < 0 Then
            Result = 180 - Result
        Else
            If Y < 0 Then
                Result = 360 + Result
            End If
        End If
    Else
        If (Y > 0) Then
            Result = 90
        Else
            Result = 270
        End If
    End If
    
    detAngle = Result

End Function

Public Function detCCW(ByVal X As Double, ByVal Y As Double) As Long

    Dim xStartCorrected, yStartCorrected, xMidCorrected, yMidCorrected, xEndCorrected, yEndCorrected As Double
    Dim radius, angle1, angle2, angle3, radAngle As Double

    xStartCorrected = PrevPrevX - xCen
    yStartCorrected = PrevPrevY - yCen
    xMidCorrected = PrevX - xCen
    yMidCorrected = PrevY - yCen
    xEndCorrected = X - xCen
    yEndCorrected = Y - yCen

    angle1 = detAngle(xStartCorrected, yStartCorrected)
    angle2 = detAngle(xMidCorrected, yMidCorrected)
    angle3 = detAngle(xEndCorrected, yEndCorrected)

    radius = Sqr(xStartCorrected * xStartCorrected + yStartCorrected * yStartCorrected)
    xStartCorrected = radius
    yStartCorrected = 0

    radAngle = ((angle2 + (360 - angle1)) / 180) * Atn(1) * 4

    xMidCorrected = radius * Cos(radAngle)
    yMidCorrected = radius * Sin(radAngle)

    radAngle = ((angle3 + (360 - angle1)) / 180) * Atn(1) * 4

    xEndCorrected = radius * Cos(radAngle)
    yEndCorrected = radius * Sin(radAngle)

    angle2 = detAngle(xMidCorrected, yMidCorrected)
    angle3 = detAngle(xEndCorrected, yEndCorrected)

    If (angle2 > angle3) Then
        detCCW = 0
    Else
        detCCW = 1
    End If
        
End Function

