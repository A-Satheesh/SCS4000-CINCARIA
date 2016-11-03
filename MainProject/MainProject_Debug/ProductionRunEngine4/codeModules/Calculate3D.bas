Attribute VB_Name = "Calculate3D"
Option Explicit

Public Function Data(line)
    Dim n, m, i, j, k As Integer
    Dim step, number, Count As Long
    Dim char, check1, check2 As String
    Dim One, Two, Three As Boolean
            
    char = ""
    check1 = ""
    check2 = ""
    n = 0
    m = 0
    CalculateX = 0
    CalculateY = 0
    CalculateZ = 0
    One = False
    step = 1
    number = Len(line)
            
     For Count = 1 To number
        char = char & Mid(line, step, 1)
        check1 = Right(char, 2)
        check2 = Right(char, 1)
        If (check1 = "x=") Or (check1 = "y=") Or (check1 = "z=") Then
            n = step
            step = step + 1
        ElseIf check2 = "," Or check2 = ")" Or (check2 = ";") Then
            m = step
            step = step + 1
            If One = False Then
                If (Mid(char, n - 1, 2) = "x=") Then
                    CalculateX = Val(Right(char, m - n))
                    'Not to include the gear ratio
                    CalculateX = Format(CalculateX, "####0.000") / 1000
                ElseIf (Mid(char, n - 1, 2) = "y=") Then
                    CalculateY = Val(Right(char, m - n))
                    CalculateY = Format(CalculateY, "####0.000") / 1000
                ElseIf (Mid(char, n - 1, 2) = "z=") Then
                    CalculateZ = Val(Right(char, m - n))
                    CalculateZ = Format(CalculateZ, "####0.000") / 1000
                    i = step
                End If
            End If
            If (Three = True) And (check2 = ")") Then
                Last = Right(char, m - j)
                Last = Left(Last, m - j - 1)
                Exit Function
            End If
            If (One = True) And (Two = True) And ((check2 = ";") Or (check2 = ")")) And (Three = False) Then
                TravelSpeed = Right(char, m - i)
                TravelSpeed = Left(TravelSpeed, m - i - 1)
                Three = True
                j = step
            End If
            If (One = True) And (check2 = ";") And (Two = False) And (Three = False) Then
                Two = True
            End If
            If (check2 = ";") And (Two = False) And (Three = False) Then
                One = True
            End If
            If check2 = ")" Then
                If (One = True) And (Two = False) And (Three = False) Then
                    k = step
                    ArcDelay = Right(char, k - i)
                    ArcDelay = Left(ArcDelay, k - i - 1)
                    Exit For
                Else
                    Exit For
                End If
            End If
        Else
            step = step + 1
        End If
    Next Count
End Function

Public Function Calculation3D(x1 As Variant, y1 As Variant, z1 As Variant, x2 As Variant, y2 As Variant, z2 As Variant, x3 As Variant, y3 As Variant, z3 As Variant) As Integer
    Dim Center(0 To 2), normal(0 To 2), angle(0 To 1), tmpt1(0 To 2), tmpt2(0 To 2), tmpt3(0 To 2), radius As Double
    Dim centerPt, normalVector, AngleReturn As Integer
    Dim height, direction, signal, same As Boolean
    Dim StringLine As String
    
    Dim gPROCISION As Double
    gPROCISION = 0.05 '0.02
    StringLine = ""
    height = False
    direction = False
    signal = False
    same = False
    
    If (z1 = z2) And (z1 = z3) And (z2 = z3) Then
        height = True
    End If
    
    Dim centerx, centery, centerZ, r As Double
    centerPt = CircleCenter(x1, y1, z1, x2, y2, z2, x3, y3, z3, Center(), radius)
    If centerPt < 0 Then
        Call MsgBox("Coordinate Error. Please Check again!", vbOKOnly)
        Calculation3D = -1
        Exit Function
    Else
        centerx = Center(0)
        centery = Center(1)
        centerZ = Center(2)
        r = radius
    End If
    
    Dim normalXX, normalYY, normalZZ As Double
    normalVector = Normal3D(x1, y1, z1, x2, y2, z2, x3, y3, z3, normal())
    If normalVector < 0 Then
        'Check it again
        MsgBox ("Coordinate Error. Please Check again!")
        Calculation3D = -1
        Exit Function
    Else
        normalXX = normal(0)
        normalYY = normal(1)
        normalZZ = normal(2)
    End If
        
    Dim angle1, angle2 As Double
    AngleReturn = NormalAngle(normalXX, normalYY, normalZZ, angle())
    If AngleReturn < 0 Then
        If normal(2) < 0 Then
            direction = True
        End If
        If (height = True) And (direction = True) Then
            angle1 = -1         'Testing (-1 or 0)
            angle2 = 0
            If (x1 >= x3) Then
                same = True
                direction = False
                signal = False
            Else
                signal = True
            End If
        Else
            angle1 = 1          'Can put 0 or 1
            angle2 = 0
        End If
    Else
        angle1 = angle(0)
        angle2 = angle(1)
    End If
    
    Call Rotation3D(x1, y1, z1, angle1, 2, tmpt1())
    Dim tmptX1, tmptY1, tmptZ1 As Double
    tmptX1 = tmpt1(0)
    tmptY1 = tmpt1(1)
    tmptZ1 = tmpt1(2)
    Call Rotation3D(tmptX1, tmptY1, tmptZ1, angle2, 1, tmpt1())
    tmptX1 = tmpt1(0)
    tmptY1 = tmpt1(1)
    tmptZ1 = tmpt1(2)
    
    Call Rotation3D(x2, y2, z2, angle1, 2, tmpt2())
    Dim tmptX2, tmptY2, tmptZ2 As Double
    tmptX2 = tmpt2(0)
    tmptY2 = tmpt2(1)
    tmptZ2 = tmpt2(2)
    Call Rotation3D(tmptX2, tmptY2, tmptZ2, angle2, 1, tmpt2())
    tmptX2 = tmpt2(0)
    tmptY2 = tmpt2(1)
    tmptZ2 = tmpt2(2)
    
    Call Rotation3D(x3, y3, z3, angle1, 2, tmpt3())
    Dim tmptX3, tmptY3, tmptZ3 As Double
    tmptX3 = tmpt3(0)
    tmptY3 = tmpt3(1)
    tmptZ3 = tmpt3(2)
    Call Rotation3D(tmptX3, tmptY3, tmptZ3, angle2, 1, tmpt3())
    tmptX3 = tmpt3(0)
    tmptY3 = tmpt3(1)
    tmptZ3 = tmpt3(2)
    
    Call Rotation3D(centerx, centery, centerZ, angle1, 2, Center())
    Dim tmptCenterX, tmptCenterY, tmptCenterZ As Double
    tmptCenterX = Center(0)
    tmptCenterY = Center(1)
    tmptCenterZ = Center(2)
    Call Rotation3D(tmptCenterX, tmptCenterY, tmptCenterZ, angle2, 1, Center())
    tmptCenterX = Center(0)
    tmptCenterY = Center(1)
    tmptCenterZ = Center(2)
    
    Dim TranDisX, TranDisY, TranDisZ As Double
    TranDisX = -Center(0)
    TranDisY = -Center(1)
    TranDisZ = -Center(2)
    Call Translation3D(tmptX1, tmptY1, tmptZ1, TranDisX, TranDisY, TranDisZ, tmpt1())
    Dim TranX1, TranY1, TranZ1 As Double
    TranX1 = tmpt1(0)
    TranY1 = tmpt1(1)
    TranZ1 = tmpt1(2)
    Call Translation3D(tmptX2, tmptY2, tmptZ2, TranDisX, TranDisY, TranDisZ, tmpt2())
    Dim TranX2, TranY2, TranZ2 As Double
    TranX2 = tmpt2(0)
    TranY2 = tmpt2(1)
    TranZ2 = tmpt2(2)
    Call Translation3D(tmptX3, tmptY3, tmptZ3, TranDisX, TranDisY, TranDisZ, tmpt3())
    Dim TranX3, TranY3, TranZ3 As Double
    TranX3 = tmpt3(0)
    TranY3 = tmpt3(1)
    TranZ3 = tmpt3(2)
    
    Dim start_angle As Double
    start_angle = DetalAngle(TranX1, TranY1)
    If start_angle < 0 Then
        start_angle = start_angle + 2 * (Atn(1) * 4)
    End If
    
    Dim end_angle As Double
    end_angle = DetalAngle(TranX3, TranY3)
    If end_angle < 0 Then
        end_angle = end_angle + 2 * (Atn(1) * 4)
    End If
    
    Dim angle_span As Double
    angle_span = 0
    If (signal = False) And (direction = False) Then
        If end_angle <= start_angle Then
            angle_span = end_angle + 2 * (Atn(1) * 4) - start_angle
        Else
            angle_span = end_angle - start_angle
        End If
        If same = True Then
            angle_span = (angle_span - 2 * (Atn(1) * 4))
            signal = True
            direction = True
        End If
    Else
        If end_angle <= start_angle Then
            angle_span = end_angle - start_angle
        Else
            angle_span = end_angle + 2 * (Atn(1) * 4) - start_angle
        End If
    End If
       
    If (gPROCISION / r > 1) Then
        'Define something?????
        'Calculation3D = -1
        'Exit Function
    End If
    
    Dim est_angle_step, Result As Double
    Result = 1 - gPROCISION / r
    est_angle_step = 2 * (Atn(-Result / Sqr(1 - Result * Result)) + Atn(1) * 2)
    
    Dim step_no As Integer
    step_no = CInt(angle_span / est_angle_step) + 1
    If (direction = True) And (signal = True) Then
        step_no = step_no * (-1)
    End If
    
    If (step_no > 450) Then     'Just define
        step_no = 450
    End If
    
    Dim actual_angle_step As Double
    If (step_no = 0) Then
        step_no = Abs(CInt(angle_span / 0.05))
        actual_angle_step = angle_span / step_no
    Else
        actual_angle_step = angle_span / step_no
    End If
    
    Dim x, y, Z As Double
    Dim i As Integer
    Z = TranZ1
    
    Dim prev_x, prev_y, prev_z As Double
    prev_x = TranX1
    prev_y = TranY1
    prev_z = TranZ1
    
    Dim ActualX, ActualY, ActualZ As Double
    Dim curAngle As Double
    curAngle = start_angle
    
    For i = 1 To step_no
        curAngle = curAngle + actual_angle_step
        x = r * Cos(curAngle)
        y = r * Sin(curAngle)
        
        Dim prev_p(0 To 2), p(0 To 2)
        Dim prev_PointX, prev_PointY, prev_PointZ As Double
        prev_PointX = prev_x
        prev_PointY = prev_y
        prev_PointZ = prev_z
        ActualX = x
        ActualY = y
        ActualZ = Z
        
        Call Rotation3D(prev_PointX, prev_PointY, prev_PointZ, -angle2, 1, prev_p())
        prev_PointX = prev_p(0)
        prev_PointY = prev_p(1)
        prev_PointZ = prev_p(2)
        Call Rotation3D(prev_PointX, prev_PointY, prev_PointZ, -angle1, 2, prev_p())
        prev_PointX = prev_p(0)
        prev_PointY = prev_p(1)
        prev_PointZ = prev_p(2)
        
        Call Rotation3D(ActualX, ActualY, ActualZ, -angle2, 1, p())
        ActualX = p(0)
        ActualY = p(1)
        ActualZ = p(2)
        Call Rotation3D(ActualX, ActualY, ActualZ, -angle1, 2, p())
        ActualX = p(0)
        ActualY = p(1)
        ActualZ = p(2)
        
        Call Translation3D(prev_PointX, prev_PointY, prev_PointZ, centerx, centery, centerZ, prev_p())
        Call Translation3D(ActualX, ActualY, ActualZ, centerx, centery, centerZ, p())
                
        prev_PointX = Format(prev_p(0), "####0.000")
        prev_PointY = Format(prev_p(1), "####0.000")
        prev_PointZ = Format(prev_p(2), "####0.000")
        
        ActualX = Format(p(0), "####0.000")
        ActualY = Format(p(1), "####0.000")
        ActualZ = Format(p(2), "####0.000")
        
        If i = 1 Then
            If (NoChange = True) Or (Change = True) Then
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '   Flag not to do "Start" and "Stop" procedure for "L-Needle"  '
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Just for "New Spray System"
                If (Change = False) Then
                    If RepeatPattern = True Then
                        StringLine = "line3D(x=-11111,y=-11111, z=-11111; sp=11111; 1)"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        A.writeline ("line3D(x=-11111, y=-11111, z=-11111; sp=11111; 1)")
                    End If
                End If
                
                If RepeatPattern = True Then
                    StringLine = "line3D(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & TravelSpeed & ")"
                    StringLine = StringLine & vbNewLine
                    ReadRepeatString = ReadRepeatString & StringLine
                Else
                    A.writeline ("line3D(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & TravelSpeed & ")")
                End If
                
                'This procedure will make wrong positionning. So, remove it.
                'For rotation (No tilting for Arc)
                'Turnning_Line_Angle ("None")
                
                NoChange = False
                Change = False
            Else
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                '   Flag not to do "Start" and "Stop" procedure for "L-Needle"  '
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                'Just for "New Spray System"
                If RepeatPattern = True Then
                    StringLine = "dot(x=11111, y=11111, z=11111; 11111, 11111; 11111, 11111; z=11111; sp=11111; 11111.000; 11111.000; z=11111)"
                    StringLine = StringLine & vbNewLine
                    ReadRepeatString = ReadRepeatString & StringLine
                Else
                    A.writeline ("dot(x=11111, y=11111, z=11111; 11111, 11111; 11111, 11111; z=11111; sp=11111; 11111.000; 11111.000; z=11111)")
                End If
                
                If RepeatPattern = True Then
                    StringLine = "start(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & ArcDelay & ")"
                    StringLine = StringLine & vbNewLine
                    ReadRepeatString = ReadRepeatString & StringLine
                Else
                    A.writeline ("start(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & ArcDelay & ")")
                End If
            End If
        ElseIf i > 1 And i <= step_no Then
            If RepeatPattern = True Then
                StringLine = "line3D(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & TravelSpeed & ")"
                StringLine = StringLine & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("line3D(x=" & CLng(prev_PointX * 1000) & ", y=" & CLng(prev_PointY * 1000) & ", z=" & CLng(prev_PointZ * 1000) & "; " & TravelSpeed & ")")
            End If
        End If
        
        prev_x = x
        prev_y = y
        prev_z = Z
    
    Next i
    
    If (NoChange2 = True) Then
        If RepeatPattern = True Then
            'StringLine = "line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & ")"
            'StringLine = StringLine & vbNewLine
            'Testing (Not to do a big sound)
            '11111 means not to close the valve
            StringLine = "end3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & "; 0.000; z=0; sp=11111; z=0)"
            StringLine = StringLine & vbNewLine
            ReadRepeatString = ReadRepeatString & StringLine
            StringLine = "start(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; 0.000)"
            StringLine = StringLine & vbNewLine
            ReadRepeatString = ReadRepeatString & StringLine
        Else
            'A.writeline ("line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & ")")
            'Testing (Not to do a big sound)
            A.writeline ("end3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & "; 0.000; z=0; sp=11111; z=0)")
            A.writeline ("start(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; 0.000)")
        End If
        NoChange2 = False
    ElseIf (NoChange3 = True) Then
        'Because this is the same position. So, jump one step
        NoChange3 = False
        Exit Function
    Else
        If RepeatPattern = True Then
            StringLine = "end3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & "; " & Last & ")"
            StringLine = StringLine & vbNewLine
            ReadRepeatString = ReadRepeatString & StringLine
        Else
            A.writeline ("end3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; " & TravelSpeed & "; " & Last & ")")
        End If
    End If
    
    RepeatPattern = False
    Calculation3D = 0
End Function
Public Function Translation3D(ByVal pointX As Double, ByVal pointY As Double, ByVal pointZ As Double, ByVal OffSetX As Double, ByVal OffSetY As Double, ByVal OffSetZ As Double, ByRef newpt() As Variant)
    
    newpt(0) = pointX + OffSetX
    newpt(1) = pointY + OffSetY
    newpt(2) = pointZ + OffSetZ
    
End Function

Public Function Rotation3D(ByVal pointX As Double, ByVal pointY As Double, ByVal pointZ As Double, ByVal angle As Double, ByVal axis As Integer, ByRef newpt() As Variant)
    
    Dim x, y, Z As Double
    x = pointX
    y = pointY
    Z = pointZ
    
    Select Case (axis)
        Case 0   'X
            newpt(0) = x
            newpt(1) = y * Cos(angle) + Z * Sin(angle)
            newpt(2) = -y * Sin(angle) + Z * Cos(angle)
        Case 1  'Y
            newpt(0) = x * Cos(angle) - Z * Sin(angle)
            newpt(1) = y
            newpt(2) = x * Sin(angle) + Z * Cos(angle)
        Case 2  'Z
            newpt(0) = x * Cos(angle) + y * Sin(angle)
            newpt(1) = -x * Sin(angle) + y * Cos(angle)
            newpt(2) = Z
    End Select
    
End Function

Public Function CircleCenter(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByRef Center() As Variant, ByRef radius As Double) As Integer
    
    Dim CirCenA, CirCenB, CirCenC, CirCenD As Double
    CirCenA = 0
    CirCenB = 0
    CirCenC = 0
    CirCenD = 0
    Call PlaneCoefficient(x1, y1, z1, x2, y2, z2, x3, y3, z3, CirCenA, CirCenB, CirCenC, CirCenD)
    
    Dim midpointX, midpointY, midpointZ As Double
    midpointX = (x1 + x2) / 2
    midpointY = (y1 + y2) / 2
    midpointZ = (z1 + z2) / 2
    Dim a1, b1, c1, d1 As Double
    a1 = x2 - x1
    b1 = y2 - y1
    c1 = z2 - z1
    d1 = -(a1 * midpointX) - (b1 * midpointY) - (c1 * midpointZ)
    
    midpointX = (x2 + x3) / 2
    midpointY = (y2 + y3) / 2
    midpointZ = (z2 + z3) / 2
    Dim a2, b2, c2, d2 As Double
    a2 = x3 - x2
    b2 = y3 - y2
    c2 = z3 - z2
    d2 = -(a2 * midpointX) - (b2 * midpointY) - (c2 * midpointZ)
    
    Dim CrossProduct_Matrix As Double
    CrossProduct_Matrix = ((CirCenA * b1 * c2) + (a1 * b2 * CirCenC) + (CirCenB * c1 * a2) - (a2 * b1 * CirCenC) - (c1 * b2 * CirCenA) - (a1 * CirCenB * c2))
    
    If CrossProduct_Matrix <= 0 Then
        CircleCenter = -1
        Exit Function
    End If
    
    Dim a11, a12, a13, a21, a22, a23, a31, a32, a33 As Double
    a11 = (b1 * c2 - c1 * b2) / CrossProduct_Matrix
    a12 = (CirCenC * b2 - CirCenB * c2) / CrossProduct_Matrix
    a13 = (CirCenB * c1 - CirCenC * b1) / CrossProduct_Matrix
    a21 = (c1 * a2 - a1 * c2) / CrossProduct_Matrix
    a22 = (CirCenA * c2 - CirCenC * a2) / CrossProduct_Matrix
    a23 = (CirCenC * a1 - CirCenA * c1) / CrossProduct_Matrix
    a31 = (a1 * b2 - b1 * a2) / CrossProduct_Matrix
    a32 = (CirCenB * a2 - CirCenA * b2) / CrossProduct_Matrix
    a33 = (CirCenA * b1 - a1 * CirCenB) / CrossProduct_Matrix
    
    Center(0) = -(a11 * CirCenD + a12 * d1 + a13 * d2)
    Center(1) = -(a21 * CirCenD + a22 * d1 + a23 * d2)
    Center(2) = -(a31 * CirCenD + a32 * d1 + a33 * d2)

    radius = Sqr((x1 - Center(0)) * (x1 - Center(0)) + (y1 - Center(1)) * (y1 - Center(1)) + (z1 - Center(2)) * (z1 - Center(2)))
    
    'If (radius > 10000 Or radius < 0.005) Then     'Define the minimum and maximum radius
    '   CircleCenter = -1
    'End If
    
    CircleCenter = 0
    
End Function

Public Function PlaneCoefficient(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByRef A As Variant, ByRef B As Variant, ByRef c As Variant, ByRef D As Variant)
    
    Dim x11, x21, x31, y11, y21, y31, z11, z21, z31 As Double
    x11 = x1
    x21 = x2 - x1
    x31 = x3 - x1
    y11 = y1
    y21 = y2 - y1
    y31 = y3 - y1
    z11 = z1
    z21 = z2 - z1
    z31 = z3 - z1
    
    A = y21 * z31 - y31 * z21
    B = z21 * x31 - x21 * z31
    c = x21 * y31 - x31 * y21
    D = (-y21 * z31 + y31 * z21) * x11 + (-z21 * x31 + x21 * z31) * y11 + (-x21 * y31 + x31 * y21) * z11
    
End Function

Public Function Normal3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByRef normal() As Variant) As Integer
    
    Dim a1, a2, a3, b1, b2, b3, c1, c2, c3, NormalD As Double
    a1 = x2 - x1
    a2 = y2 - y1
    a3 = z2 - z1

    b1 = x3 - x1
    b2 = y3 - y1
    b3 = z3 - z1

    c1 = a2 * b3 - a3 * b2
    c2 = a3 * b1 - a1 * b3
    c3 = a1 * b2 - a2 * b1
    
    'If (Abs(c1) < 0.0001 And Abs(c2) < 0.0001 And Abs(c3) < 0.0001) Then   'Must deifne
    '    Normal3D = -1
    'End If
    
    If (Abs(c1) <= 0) And (Abs(c2) <= 0) And (Abs(c3) <= 0) Then
        Normal3D = -1
        Exit Function
    End If
    
    NormalD = Sqr(c1 * c1 + c2 * c2 + c3 * c3)
    
    normal(0) = c1 / NormalD
    normal(1) = c2 / NormalD
    normal(2) = c3 / NormalD
    
    Normal3D = 0
    
End Function

Public Function NormalAngle(ByVal normalX As Double, ByVal normalY As Double, ByVal normalZ As Double, ByRef angle() As Variant) As Integer
    
    Dim r As Double
    r = Sqr(normalX * normalX + normalY * normalY + normalZ * normalZ)
    
    'If (r < 0.0001) Then       'Must define
    '    NormalAngle = -1
    'End If
    
    If (Abs(normalX) < 0.0001 And Abs(normalY) < 0.0001) Then 'Checking for whithout height
        NormalAngle = -1
        Exit Function
    End If

    angle(0) = DetalAngle(normalX, normalY)
    
    Dim Value As Double
    Value = normalZ / r
    angle(1) = Atn(-Value / Sqr(1 - Value * Value)) + 2 * Atn(1)
    
    NormalAngle = 0
    
End Function
Public Function DetalAngle(ByVal x As Double, ByVal y As Double) As Double
    Dim sinevalue, Result, radius As Double
    Dim Pi As Double
    Pi = 3.14159265359
    
    radius = Sqr((x * x) + (y * y))
    sinevalue = y / radius
    
    If (sinevalue <> 1 And sinevalue <> -1) Then
        Result = Atn(sinevalue / Sqr(-sinevalue * sinevalue + 1))
        If x < 0 Then
            Result = Pi - Result
        Else
            If y < 0 Then
                Result = 2 * Pi + Result
            End If
        End If
    Else
        If y > 0 Then
            Result = Pi / 2
        Else
            Result = 3 * (Pi / 2)
        End If
    End If
    DetalAngle = Result
End Function




