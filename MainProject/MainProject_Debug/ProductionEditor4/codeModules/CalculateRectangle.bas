Attribute VB_Name = "CalculateRectangle"
Option Explicit

Dim Last_Pair As Boolean

Public Function Spray_CalculationRectangle_Xplus(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double, ByVal Pitch_Dis As Double, ByVal No_Fill_Area As Integer) As Integer
    Dim r As Integer, C As Integer, Next_Row As Integer
    Dim x4, y4, z4 As Double, Point4(0 To 2) As Double, old_value_y As Double
    Dim Delta_X As Double, New_PositionX As Double, New_PositionY As Double
    Dim StringLine As String
    
    Call FindPoint4(x1, y1, z1, x2, y2, z2, x3, y3, z3, Point4())
    x4 = Point4(0)
    y4 = Point4(1)
    z4 = Point4(2)
    
    If (No_Fill_Area = 0) Then
        New_PositionX = x2
        New_PositionY = y2
    Else
        New_PositionX = x1
        New_PositionY = y1
    End If
    Delta_X = x2 - x1
    
    '''''''''''''''''''''
    '   Drawing "Up"    '
    '''''''''''''''''''''
    If (y2 > y3) Then
        If (y3 > (y2 - Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Xplus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
  
            r = 2
            Do While (r)
                New_PositionY = New_PositionY - Pitch_Dis
            
                If (New_PositionY < y3) Then
                    New_PositionY = New_PositionY + Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
            
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY + Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY + Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    'If (c = 1) Then
                    '    If RepeatPattern = True Then
                    '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                    '        StringLine = StringLine & vbNewLine
                    '        ReadRepeatString = ReadRepeatString & StringLine
                    '    Else
                    '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                    '    End If
                    'End If
                    
                    ''''''''''''''''''''''''
                    '   Always "ON/OFF"    '
                    ''''''''''''''''''''''''
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
    
            Do While (r)
                If (New_PositionY < y3) Then
                    r = r - 1
                    New_PositionY = New_PositionY + Pitch_Dis
                
                    If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(0) > New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(0) < New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
        
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionY = New_PositionY + Pitch_Dis
                        old_value_y = New_PositionY
                    
                        If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(0) > New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(0) < New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        '''''''''''''''''''''
                        '   Always "OFf"    '
                        '''''''''''''''''''''
                        'If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                        '    'If (Area_P2(0) > New_PositionX) Then
                        '    If (Area_P1(1) > New_PositionY) And (Area_P1(1) < (New_PositionY + Pitch_Dis)) Then
                        '        If RepeatPattern = True Then
                        '            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                        '            StringLine = StringLine & vbNewLine
                        '            ReadRepeatString = ReadRepeatString & StringLine
                        '        Else
                        '            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        '        End If
                        '    End If
                        'End If
                        r = r + 1
                        New_PositionY = New_PositionY - Pitch_Dis
                    End If
                
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(0) > New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(0) > New_PositionX) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                                A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            ''''''''''''''''''''''''
                            '   Always "ON/OFF"    '
                            ''''''''''''''''''''''''
                            'If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                            '    If (Area_P3(1) > old_value_y) And (Area_P3(1) < New_PositionY) Then
                            '        If RepeatPattern = True Then
                            '            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P3(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                            '            StringLine = StringLine & vbNewLine
                            '            ReadRepeatString = ReadRepeatString & StringLine
                            '        Else
                            '            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P3(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            '        End If
                            '    End If
                            'End If
                        
                            'If (New_PositionY > Area_P1(1)) And (New_PositionY < Area_P3(1)) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                            '    If RepeatPattern = True Then
                            '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                            '        StringLine = StringLine & vbNewLine
                            '        ReadRepeatString = ReadRepeatString & StringLine
                            '    Else
                            '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            '    End If
                            'Else
                            '    If RepeatPattern = True Then
                            '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            '        StringLine = StringLine & vbNewLine
                            '        ReadRepeatString = ReadRepeatString & StringLine
                            '    Else
                            '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            '    End If
                            'End If
                        
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                            'If (Area_P3(1) > New_PositionY) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(0) < New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(0) > New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    End If
                Next
            
                New_PositionY = New_PositionY - Pitch_Dis
                r = r + 1
            Loop
        End If
    '''''''''''''''''''''
    '   Drawing "Down"  '
    '''''''''''''''''''''
    Else
        If (y3 < (y2 + Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Xplus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
  
            r = 2
            Do While (r)
                New_PositionY = New_PositionY + Pitch_Dis
            
                If (New_PositionY > y3) Then
                    New_PositionY = New_PositionY - Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
            
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY - Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY - Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    'If (c = 1) Then
                    '    If RepeatPattern = True Then
                    '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                    '        StringLine = StringLine & vbNewLine
                    '        ReadRepeatString = ReadRepeatString & StringLine
                    '    Else
                    '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                    '    End If
                    'End If
                    ''''''''''''''''''''''''
                    '   Always "ON/OFF"    '
                    ''''''''''''''''''''''''
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
    
            Do While (r)
                If (New_PositionY > y3) Then
                    r = r - 1
                    New_PositionY = New_PositionY - Pitch_Dis
                
                    If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(0) > New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(0) < New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
        
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionY = New_PositionY - Pitch_Dis
                        old_value_y = New_PositionY
                    
                        If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(0) > New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(0) < New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        '''''''''''''''''''''
                        '   Always "OFf"    '
                        '''''''''''''''''''''
                        'If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                        '    'If (Area_P2(0) > New_PositionX) Then
                        '    If (Area_P1(1) > New_PositionY) And (Area_P1(1) < (New_PositionY + Pitch_Dis)) Then
                        '        If RepeatPattern = True Then
                        '            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                        '            StringLine = StringLine & vbNewLine
                        '            ReadRepeatString = ReadRepeatString & StringLine
                        '        Else
                        '            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        '        End If
                        '    End If
                        'End If
                        r = r + 1
                        New_PositionY = New_PositionY + Pitch_Dis
                    End If
                
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(0) > New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(0) > New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            '''''''''''''''''''''
                            '   Always "OFF"    '
                            '''''''''''''''''''''
                            'If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                            '    If (Area_P3(1) > old_value_y) And (Area_P3(1) < New_PositionY) Then
                            '        If RepeatPattern = True Then
                            '            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P3(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                            '            StringLine = StringLine & vbNewLine
                            '            ReadRepeatString = ReadRepeatString & StringLine
                            '        Else
                            '            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P3(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            '        End If
                            '    End If
                            'End If
                        
                            'If (New_PositionY > Area_P1(1)) And (New_PositionY < Area_P3(1)) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P2(0)) Then
                            '    If RepeatPattern = True Then
                            '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                            '        StringLine = StringLine & vbNewLine
                            '        ReadRepeatString = ReadRepeatString & StringLine
                            '    Else
                            '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            '    End If
                            'Else
                            '    If RepeatPattern = True Then
                            '        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            '        StringLine = StringLine & vbNewLine
                            '        ReadRepeatString = ReadRepeatString & StringLine
                            '    Else
                            '        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            '    End If
                            'End If
                            
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                            'If (Area_P3(1) > New_PositionY) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(0) < New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(0) > New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    End If
                Next
            
                New_PositionY = New_PositionY + Pitch_Dis
                r = r + 1
            Loop
        End If
    End If
    
    RepeatPattern = False
    Initialize_Fill_Area
    Spray_CalculationRectangle_Xplus = 0
End Function

Public Function Spray_CalculationRectangle_Xminus(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double, ByVal Pitch_Dis As Double, ByVal No_Fill_Area As Integer) As Integer
    Dim r As Integer, C As Integer, Next_Row As Integer
    Dim Delta_X As Double, New_PositionX As Double, New_PositionY As Double
    Dim StringLine As String

    If (No_Fill_Area = 0) Then
        New_PositionX = x2
        New_PositionY = y2
    Else
        New_PositionX = x1
        New_PositionY = y1
    End If
    Delta_X = x1 - x2
    
    '''''''''''''''''''''
    '   Drawing "Up"    '
    '''''''''''''''''''''
    If (y2 > y3) Then
        If (y3 > (y2 - Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Xminus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionY = New_PositionY - Pitch_Dis
                
                If (New_PositionY < y3) Then
                    New_PositionY = New_PositionY + Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
            
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY + Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY + Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionY < y3) Then
                    r = r - 1
                    New_PositionY = New_PositionY + Pitch_Dis
                
                    If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(0) < New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(0) > New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                    
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P2(0)) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionY = New_PositionY + Pitch_Dis
                    
                        If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(0) < New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(0) > New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P2(0)) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                
                        r = r + 1
                        New_PositionY = New_PositionY - Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(0) < New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(0) < New_PositionX) And (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                                A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionY <= Area_P1(1)) And (New_PositionY >= Area_P3(1)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(0) > New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(0) < New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    End If
                Next
                
                New_PositionY = New_PositionY - Pitch_Dis
                r = r + 1
            Loop
        End If
    '''''''''''''''''''''
    '   Drawing "Down"  '
    '''''''''''''''''''''
    Else
        If (y3 < (y2 + Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Xminus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionY = New_PositionY + Pitch_Dis
                
                If (New_PositionY > y3) Then
                    New_PositionY = New_PositionY - Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
            
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY - Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng((New_PositionY - Pitch_Dis) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionY > y3) Then
                    r = r - 1
                    New_PositionY = New_PositionY - Pitch_Dis
                
                    If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(0) < New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(0) > New_PositionX) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                    
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P2(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionY = New_PositionY - Pitch_Dis
                    
                        If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(0) < New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(0) > New_PositionX) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P2(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                
                        r = r + 1
                        New_PositionY = New_PositionY + Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(0) < New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(0) < New_PositionX) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P3(1)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(0) > New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P2(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(0) < New_PositionX) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(Area_P1(0) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionX = New_PositionX + Delta_X
                    ElseIf (C <> 2) Then
                        New_PositionX = New_PositionX - Delta_X
                    End If
                Next
                
                New_PositionY = New_PositionY + Pitch_Dis
                r = r + 1
            Loop
        End If
    End If
    
    RepeatPattern = False
    Initialize_Fill_Area
    Spray_CalculationRectangle_Xminus = 0
End Function

Public Function Spray_CalculationRectangle_Yplus(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double, ByVal Pitch_Dis As Double, ByVal No_Fill_Area As Integer) As Integer
    Dim r As Integer, C As Integer, Next_Row As Integer
    Dim Delta_Y As Double, New_PositionX As Double, New_PositionY As Double
    Dim StringLine As String
    
    If (No_Fill_Area = 0) Then
        New_PositionX = x2
        New_PositionY = y2
    Else
        New_PositionX = x1
        New_PositionY = y1
    End If
    Delta_Y = y2 - y1
    
    '''''''''''''''''''''
    '   Drawing "Left"  '
    '''''''''''''''''''''
    If (x2 > x3) Then
        If (x3 > (x2 - Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Yplus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionX = New_PositionX - Pitch_Dis
            
                If (New_PositionX < x3) Then
                    New_PositionX = New_PositionX + Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
                
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng((New_PositionX + Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng((New_PositionX + Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
                
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionX < x3) Then
                    r = r - 1
                    New_PositionX = New_PositionX + Pitch_Dis
                
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(1) > New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(1) < New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionX = New_PositionX + Pitch_Dis
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(1) > New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(1) < New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        r = r + 1
                        New_PositionX = New_PositionX - Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(1) > New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(1) > New_PositionY) And (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(1) < New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(1) > New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    End If
                Next
            
                New_PositionX = New_PositionX - Pitch_Dis
                r = r + 1
            Loop
        End If
    '''''''''''''''''''''''''
    '   Drawing "Right"     '
    '''''''''''''''''''''''''
    Else
        If (x3 < (x2 + Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Yplus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionX = New_PositionX + Pitch_Dis
            
                If (New_PositionX > x3) Then
                    New_PositionX = New_PositionX - Pitch_Dis
                
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
                
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng((New_PositionX - Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng((New_PositionX - Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
                
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionX > x3) Then
                    r = r - 1
                    New_PositionX = New_PositionX - Pitch_Dis
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(1) > New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(1) < New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionX = New_PositionX - Pitch_Dis
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(1) > New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(1) < New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        r = r + 1
                        New_PositionX = New_PositionX + Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        
                            If (Area_P1(1) > New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(1) > New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(1) < New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(1) > New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    End If
                Next
            
                New_PositionX = New_PositionX + Pitch_Dis
                r = r + 1
            Loop
        End If
    End If
    
    RepeatPattern = False
    Initialize_Fill_Area
    Spray_CalculationRectangle_Yplus = 0
End Function

Public Function Spray_CalculationRectangle_Yminus(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double, ByVal Pitch_Dis As Double, ByVal No_Fill_Area As Integer) As Integer
    Dim r As Integer, C As Integer, Next_Row As Integer
    Dim Delta_Y As Double, New_PositionX As Double, New_PositionY As Double
    Dim StringLine As String
    
    If (No_Fill_Area = 0) Then
        New_PositionX = x2
        New_PositionY = y2
    Else
        New_PositionX = x1
        New_PositionY = y1
    End If
    Delta_Y = y1 - y2
    
    '''''''''''''''''''''
    '   Drawing "Left"  '
    '''''''''''''''''''''
    If (x2 > x3) Then
        If (x3 > (x2 - Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Yminus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionX = New_PositionX - Pitch_Dis
            
                If (New_PositionX < x3) Then
                    New_PositionX = New_PositionX + Pitch_Dis
                    
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
                
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng((New_PositionX + Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng((New_PositionX + Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionX < x3) Then
                    r = r - 1
                    New_PositionX = New_PositionX + Pitch_Dis
                
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(1) < New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(1) > New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionX = New_PositionX + Pitch_Dis
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(1) < New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(1) > New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        r = r + 1
                        New_PositionX = New_PositionX - Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                            
                            If (Area_P1(1) < New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(1) < New_PositionY) And (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionX <= Area_P1(0)) And (New_PositionX >= Area_P3(0)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(1) > New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(1) < New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    End If
                Next
            
                New_PositionX = New_PositionX - Pitch_Dis
                r = r + 1
            Loop
        End If
    '''''''''''''''''''''''''
    '   Drawing "Right"     '
    '''''''''''''''''''''''''
    Else
        If (x3 < (x2 + Pitch_Dis)) Then
            MsgBox "Pitch_distance is out of range!"
            Spray_CalculationRectangle_Yminus = -1
            Exit Function
        End If
    
        If (No_Fill_Area = 0) Then
            If RepeatPattern = True Then
                StringLine = "Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")" & vbNewLine
                StringLine = StringLine & "line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")" & vbNewLine
                ReadRepeatString = ReadRepeatString & StringLine
            Else
                A.writeline ("Start(x=" & CLng(Format$(x1, "####0.000") * 1000) & ", y=" & CLng(Format$(y1, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & ArcDelay & ")")
                A.writeline ("line3D(x=" & CLng(Format$(x2, "####0.000") * 1000) & ", y=" & CLng(Format$(y2, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; sp=" & Speed & "; " & 1 & ")")
            End If
        
            r = 2
            Do While (r)
                New_PositionX = New_PositionX + Pitch_Dis
            
                If (New_PositionX > x3) Then
                    New_PositionX = New_PositionX - Pitch_Dis
                    
                    If RepeatPattern = True Then
                        'SCS4000M
                        'StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
                        'SCS4000N
                        StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")"
                        StringLine = StringLine & vbNewLine
                        ReadRepeatString = ReadRepeatString & StringLine
                    Else
                        'SCS4000M
                        'A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
                        'SCS4000N
                        A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(TravelSpeed, Len(TravelSpeed) - 1) & "1; " & Last & ")")
                    End If
                
                    Exit Do
                Else
                    If (r <> 2) Then
                        If RepeatPattern = True Then
                            StringLine = "line3D(x=" & CLng((New_PositionX - Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("line3D(x=" & CLng((New_PositionX - Pitch_Dis) * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                        End If
                    End If
                End If
            
                For C = 1 To 2
                    If (C = 1) Then
                        If (CInt(Middle_Pt_OnOff) = 1) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        End If
                    End If
                
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    End If
                Next
            
                r = r + 1
            Loop
        Else
            r = 1
            Do While (r)
                If (New_PositionX > x3) Then
                    r = r - 1
                    New_PositionX = New_PositionX - Pitch_Dis
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                        If (r Mod 2 = 0) Then
                            If (Area_P1(1) < New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        Else
                            If (Area_P2(1) > New_PositionY) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                        End If
                    End If
                
                    If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Left(Travel_Speed_Rec, Len(Travel_Speed_Rec) - 1) & " 0; " & Last_Rec & ")")
                        End If
                    Else
                        If RepeatPattern = True Then
                            StringLine = "End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                        Else
                            A.writeline ("End3D(x=" & CLng(Format$(New_PositionX, "####0.000") * 1000) & ", y=" & CLng(Format$(New_PositionY, "####0.000") * 1000) & ", z=" & CLng(Format$(z1, "####0.000") * 1000) & "; " & Travel_Speed_Rec & "; " & Last_Rec & ")")
                        End If
                    End If
                    
                    Initialize_Fill_Area
                    Exit Do
                Else
                    If (r <> 1) Then
                        r = r - 1
                        New_PositionX = New_PositionX - Pitch_Dis
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                            If (r Mod 2 = 0) Then
                                If (Area_P1(1) < New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            Else
                                If (Area_P2(1) > New_PositionY) Then
                                    If RepeatPattern = True Then
                                        StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                        StringLine = StringLine & vbNewLine
                                        ReadRepeatString = ReadRepeatString & StringLine
                                    Else
                                        A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                    End If
                                End If
                            End If
                        End If
                    
                        If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) And (New_PositionY >= Area_P1(1)) And (New_PositionY <= Area_P2(1)) Then
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                            End If
                        Else
                            If RepeatPattern = True Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            Else
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                        
                        r = r + 1
                        New_PositionX = New_PositionX + Pitch_Dis
                    End If
                End If
        
                For C = 1 To 2
                    If (C = 1) And (r = 1) Then
                        If RepeatPattern = True Then
                            StringLine = "Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
                            StringLine = StringLine & vbNewLine
                            ReadRepeatString = ReadRepeatString & StringLine
                            
                            If (Area_P1(1) < New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                StringLine = StringLine & vbNewLine
                                ReadRepeatString = ReadRepeatString & StringLine
                            End If
                        Else
                            A.writeline ("Start(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
                     
                            If (Area_P1(1) < New_PositionY) And (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                            End If
                        End If
                    Else
                        If (C = 1) Then
                            If (CInt(Middle_Pt_OnOff) = 1) Then
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 1)")
                                End If
                            Else
                                If RepeatPattern = True Then
                                    StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)"
                                    StringLine = StringLine & vbNewLine
                                    ReadRepeatString = ReadRepeatString & StringLine
                                Else
                                    A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(New_PositionY * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; 0)")
                                End If
                            End If
                            
                            If (New_PositionX >= Area_P1(0)) And (New_PositionX <= Area_P3(0)) Then
                                If (r Mod 2 = 0) Then
                                    If (Area_P2(1) > New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P2(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                Else
                                    If (Area_P1(1) < New_PositionY) Then
                                        If RepeatPattern = True Then
                                            StringLine = "line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")"
                                            StringLine = StringLine & vbNewLine
                                            ReadRepeatString = ReadRepeatString & StringLine
                                        Else
                                            A.writeline ("line3D(x=" & CLng(New_PositionX * 1000) & ", y=" & CLng(Area_P1(1) * 1000) & ", z=" & CLng(z1 * 1000) & "; sp=" & Speed & "; " & 1 & ")")
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                                    
                    If (r Mod 2 = 0) And (C <> 2) Then
                        New_PositionY = New_PositionY + Delta_Y
                    ElseIf (C <> 2) Then
                        New_PositionY = New_PositionY - Delta_Y
                    End If
                Next
            
                New_PositionX = New_PositionX + Pitch_Dis
                r = r + 1
            Loop
        End If
    End If
    
    RepeatPattern = False
    Initialize_Fill_Area
    Spray_CalculationRectangle_Yminus = 0
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Change new procedure for new spray system becuase of a big sound    '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CalculationRectangle_XW(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double) As Integer
    Dim r As Double
    Dim x4, y4, z4 As Double
    Dim Point4(0 To 2) As Double
    Dim StringLine As String
    
    Call FindPoint4(x1, y1, z1, x2, y2, z2, x3, y3, z3, Point4())
    x4 = Point4(0)
    y4 = Point4(1)
    z4 = Point4(2)
    
    If RepeatPattern = True Then
        StringLine = "Start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("Start(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & ArcDelay & ")")
    End If
    
    If RepeatPattern = True Then
        StringLine = "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
    End If
    
    If RepeatPattern = True Then
        StringLine = "line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("line3D(x=" & CLng(x3 * 1000) & ", y=" & CLng(y3 * 1000) & ", z=" & CLng(z3 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
    End If
    
    If RepeatPattern = True Then
        StringLine = "line3D(x=" & CLng(x4 * 1000) & ", y=" & CLng(y4 * 1000) & ", z=" & CLng(z4 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("line3D(x=" & CLng(x4 * 1000) & ", y=" & CLng(y4 * 1000) & ", z=" & CLng(z4 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
    End If
    
    If RepeatPattern = True Then
        StringLine = "End3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & "; " & Last & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("End3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & TravelSpeed & "; " & Last & ")")
    End If
    
    RepeatPattern = False
    CalculationRectangle_XW = 0
End Function

Public Function CalculationRectangle(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal Speed As Double) As Integer
    Dim r As Double
    Dim x4, y4, z4 As Double
    Dim Point4(0 To 2) As Double
    Dim NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2 As Double
    Dim NewValueX3, NewValueY3, NewValueZ3, NewValueX4, NewValueY4, NewValueZ4 As Double
    Dim NewValueX5, NewValueY5, NewValueZ5, NewValueX6, NewValueY6, NewValueZ6 As Double
    Dim NewValueX7, NewValueY7, NewValueZ7, NewValueX8, NewValueY8, NewValueZ8 As Double
    Dim NewCenterX1, NewCenterY1, NewCenterZ1, NewCenterX2, NewCenterY2, NewCenterZ2 As Double
    Dim NewCenterX3, NewCenterY3, NewCenterZ3, NewCenterX4, NewCenterY4, NewCenterZ4 As Double
    Dim StringLine As String
    
    Call FindPoint4(x1, y1, z1, x2, y2, z2, x3, y3, z3, Point4())
    x4 = Point4(0)
    y4 = Point4(1)
    z4 = Point4(2)
    
    r = 0
    If Speed >= 1 And Speed <= 50 Then
        r = 0.167
    ElseIf Speed > 50 And Speed <= 100 Then
        'r = 0.667
        r = 1.5
    ElseIf Speed > 100 And Speed <= 150 Then
        'r = 1.5
        r = 2.667
    ElseIf Speed > 150 And Speed <= 200 Then
        'r = 2.667
        r = 4.167
    ElseIf Speed > 200 And Speed <= 250 Then
        'r = 4.167
        r = 6
    ElseIf Speed > 250 And Speed <= 300 Then
        'r = 6
        r = 8.167
    ElseIf Speed > 300 And Speed <= 350 Then
        'r = 8.167
        r = 10.667
    ElseIf Speed > 350 And Speed <= 400 Then
        'r = 10.667
        r = 13.5
    ElseIf Speed > 400 And Speed <= 450 Then
        'r = 13.5
        r = 16.667
    ElseIf Speed > 450 And Speed <= 500 Then
        r = 20
    End If
    
    If NewPointOnRectangle(x1, y1, z1, x2, y2, z2, x3, y3, z3, r, NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1) < 0 Then
        CalculationRectangle = -1
        Exit Function
    End If
    If NewPointOnRectangle(x2, y2, z2, x3, y3, z3, x4, y4, z4, r, NewValueX3, NewValueY3, NewValueZ3, NewValueX4, NewValueY4, NewValueZ4, NewCenterX2, NewCenterY2, NewCenterZ2) < 0 Then
        CalculationRectangle = -1
        Exit Function
    End If
    If NewPointOnRectangle(x3, y3, z3, x4, y4, z4, x1, y1, z1, r, NewValueX5, NewValueY5, NewValueZ5, NewValueX6, NewValueY6, NewValueZ6, NewCenterX3, NewCenterY3, NewCenterZ3) < 0 Then
        CalculationRectangle = -1
        Exit Function
    End If
    If NewPointOnRectangle(x4, y4, z4, x1, y1, z1, x2, y2, z2, r, NewValueX7, NewValueY7, NewValueZ7, NewValueX8, NewValueY8, NewValueZ8, NewCenterX4, NewCenterY4, NewCenterZ4) < 0 Then
        CalculationRectangle = -1
        Exit Function
    End If
    
    If RepeatPattern = True Then
        StringLine = "Start(x=" & CLng(Format$(NewValueX8, "####0.000") * 1000) & ", y=" & CLng(Format$(NewValueY8, "####0.000") * 1000) & ", z=" & CLng(Format$(NewValueZ8, "####0.000") * 1000) & "; " & ArcDelay & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("Start(x=" & CLng(Format$(NewValueX8, "####0.000") * 1000) & ", y=" & CLng(Format$(NewValueY8, "####0.000") * 1000) & ", z=" & CLng(Format$(NewValueZ8, "####0.000") * 1000) & "; " & ArcDelay & ")")
    End If
    
    Call Rectangle2D3D(NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1, r, Speed)
    Call Rectangle2D3D(NewValueX3, NewValueY3, NewValueZ3, NewValueX4, NewValueY4, NewValueZ4, NewCenterX2, NewCenterY2, NewCenterZ2, r, Speed)
    Call Rectangle2D3D(NewValueX5, NewValueY5, NewValueZ5, NewValueX6, NewValueY6, NewValueZ6, NewCenterX3, NewCenterY3, NewCenterZ3, r, Speed)
    Last_Pair = True
    Call Rectangle2D3D(NewValueX7, NewValueY7, NewValueZ7, NewValueX8, NewValueY8, NewValueZ8, NewCenterX4, NewCenterY4, NewCenterZ4, r, Speed)
    Last_Pair = False
    
    If RepeatPattern = True Then
        StringLine = "End3D(x=" & CLng(Format$(NewValueX8, "####0.000") * 1000) & ", y=" & CLng(Format$(NewValueY8, "####0.000") * 1000) & ", z=" & CLng(Format$(NewValueZ8, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")"
        StringLine = StringLine & vbNewLine
        ReadRepeatString = ReadRepeatString & StringLine
    Else
        A.writeline ("End3D(x=" & CLng(Format$(NewValueX8, "####0.000") * 1000) & ", y=" & CLng(Format$(NewValueY8, "####0.000") * 1000) & ", z=" & CLng(Format$(NewValueZ8, "####0.000") * 1000) & "; " & TravelSpeed & "; " & Last & ")")
    End If
    
    RepeatPattern = False
    CalculationRectangle = 0
End Function

Public Function FindPoint4(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, Point4() As Double)
    Point4(0) = x1 + x3 - x2
    Point4(1) = y1 + y3 - y2
    Point4(2) = z1 + z3 - z2
End Function

Public Function NewPointOnRectangle(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal r As Double, ByRef NewX As Variant, NewY As Variant, NewZ As Variant, NewXX As Variant, NewYY As Variant, NewZZ As Variant, NewFilletCenterX As Variant, NewFilletCenterY As Variant, NewFilletCenterZ As Variant) As Integer
    Dim FilletCenterX, FilletCenterY, FilletCenterZ As Double
    Dim LineFillet As Integer
    Dim x4, y4, z4 As Double
    
    x4 = x3
    y4 = y3
    z4 = z3
    x3 = x2
    y3 = y2
    z3 = z2
    
    If (z1 = z2) And (z1 = z4) And (z2 = z4) Then
        LineFillet = LinesFillet2D(x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, r, FilletCenterX, FilletCenterY, FilletCenterZ)
        If LineFillet = -1 Then
            NewPointOnRectangle = -1
            Exit Function
        Else
            NewX = x2
            NewY = y2
            NewZ = z2
            NewXX = x3
            NewYY = y3
            NewZZ = z3
            NewFilletCenterX = FilletCenterX
            NewFilletCenterY = FilletCenterY
            NewFilletCenterZ = FilletCenterZ
        End If
    Else
        LineFillet = LinesFilletArc(x1, y1, z1, x2, y2, z2, x3, y3, z3, x4, y4, z4, r, FilletCenterX, FilletCenterY, FilletCenterZ)
        If LineFillet = -1 Then
            NewPointOnRectangle = -1
            Exit Function
        Else
            NewX = x2
            NewY = y2
            NewZ = z2
            NewXX = x3
            NewYY = y3
            NewZZ = z3
            NewFilletCenterX = FilletCenterX
            NewFilletCenterY = FilletCenterY
            NewFilletCenterZ = FilletCenterZ
        End If
    End If
    NewPointOnRectangle = 0
End Function

Public Function LinesFillet2D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, x2 As Variant, y2 As Variant, z2 As Variant, x3 As Variant, y3 As Variant, z3 As Variant, ByVal x4 As Double, ByVal y4 As Double, ByVal z4 As Double, ByVal r As Double, FilletCenterX As Variant, FilletCenterY As Variant, FilletCenterZ As Variant) As Integer
    
    Dim LineA1, LineB1, LineC1, LineA2, LineB2, LineC2, LineConstant1, LineConstant2 As Double
    Dim PointToLine1, PointToLine2, NewX1, NewY1, NewX2, NewY2, SlopeDistance, radius As Double
    Dim midPoint(0 To 1), centerPoint(0 To 1)
        
    Call LineCoefficient(x1, y1, x2, y2, LineA1, LineB1, LineC1)
    Call LineCoefficient(x3, y3, x4, y4, LineA2, LineB2, LineC2)
    
    ' Two lines are parallel to each other if their slopes are equal.
    If Abs((LineA1 * LineB2) - (LineA2 * LineB1)) < 0.0001 Then
        MsgBox ("The two lines are parallel or coincident!")   '  /* Parallel or coincident lines /*
        LinesFillet2D = -1
        Exit Function
    End If
    
    midPoint(0) = (x3 + x4) / 2
    midPoint(1) = (y3 + y4) / 2
    If LineToPoint(LineA1, LineB1, LineC1, midPoint, PointToLine1) < 0 Then 'Find distance p1p2 to p3 */
        LinesFillet2D = -1
        Exit Function
    End If
    
    midPoint(0) = (x1 + x2) / 2
    midPoint(1) = (y1 + y2) / 2
    If LineToPoint(LineA2, LineB2, LineC2, midPoint, PointToLine2) < 0 Then ' Find distance p3p4 to p2 */
        LinesFillet2D = -1
        Exit Function
    End If
    
    radius = r
    If (PointToLine1 <= 0) Then
        radius = -radius
    End If
    
    LineConstant1 = LineC1 - radius * Sqr((LineA1 * LineA1) + (LineB1 * LineB1)) ' Line parallel l1 at d */
    
    radius = r
    If (PointToLine2 <= 0) Then
        radius = -radius
    End If

    LineConstant2 = LineC2 - radius * Sqr((LineA2 * LineA2) + (LineB2 * LineB2)) 'Line parallel l2 at d */
    
    SlopeDistance = LineA1 * LineB2 - LineA2 * LineB1
    FilletCenterX = (LineConstant2 * LineB1 - LineConstant1 * LineB2) / SlopeDistance    ' Intersect constructed lines */
    FilletCenterY = (LineConstant1 * LineA2 - LineConstant2 * LineA1) / SlopeDistance    ' to find center of arc */
    centerPoint(0) = FilletCenterX
    centerPoint(1) = FilletCenterY
    
    Call Pointperp(LineA1, LineB1, LineC1, centerPoint, NewX1, NewY1)  ' Clip or extend lines as required */
    Call Pointperp(LineA2, LineB2, LineC2, centerPoint, NewX2, NewY2)
    
    x2 = NewX1
    y2 = NewY1
    x3 = NewX2
    y3 = NewY2
    FilletCenterZ = z2
   
    LinesFillet2D = 0
End Function

Public Function LinesFilletArc(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, x2 As Variant, y2 As Variant, z2 As Variant, x3 As Variant, y3 As Variant, z3 As Variant, ByVal x4 As Double, ByVal y4 As Double, ByVal z4 As Double, ByVal r As Double, _
                                 FilletCenterX As Variant, FilletCenterY As Variant, FilletCenterZ As Variant) As Integer
                                
    Dim LineA1, LineB1, LineC1, LineA2, LineB2, LineC2, LineConstant1, LineConstant2 As Double
    Dim PointToLine1, PointToLine2, NewX1, NewX2, NewY1, NewY2, SlopeDistance, radius As Double
    Dim midPoint(0 To 1), centerPoint(0 To 1)
    Dim PlaneCoefficientA, PlaneCoefficientB, PlaneCoefficientC, PlaneCoefficientD As Double
    Dim tmpx1, tmpx2, tmpx3, tmpx4, tmpy1, tmpy2, tmpy3, tmpy4 As Double
    Dim dx, dy, Distance As Double
    Dim tmpCenterX, tmpCenterY As Double
    Dim plane As Integer
   
    plane = 0
    dx = x1 - x4
    dy = y1 - y4
    Distance = Sqr(dx * dx + dy * dy)
    
    'If (distance < r + pr) Then
    If (Distance < r) Then
        MsgBox ("He He 1 notic!")
        LinesFilletArc = -1
        Exit Function
    End If
    
    If (((y1 = y2) And (x2 = x4)) And ((z1 = z2) Or (z2 = z4))) Or (((x1 = x2) And (y2 = y4)) And ((z1 = z2) Or (z2 = z4))) _
        Or (((z1 = z2) And (z2 <> z4)) And ((x1 = x2) Or (x2 = x4))) Or (((z1 <> z2) And (z2 = z4)) And ((x1 = x2) Or (x2 = x4))) Then 'X-Y plane
        plane = 3
        tmpx1 = x1
        tmpy1 = y1
        tmpx2 = x2
        tmpy2 = y2
        tmpx3 = x3
        tmpy3 = y3
        tmpx4 = x4
        tmpy4 = y4
    Else
        If (Abs(dx) > Abs(dy)) Then 'Z-X plane
            plane = 2
            tmpx1 = x1
            tmpy1 = z1
            tmpx2 = x2
            tmpy2 = z2
            tmpx3 = x3
            tmpy3 = z3
            tmpx4 = x4
            tmpy4 = z4
        Else  'Y-Z plane
            plane = 1
            tmpx1 = y1
            tmpy1 = z1
            tmpx2 = y2
            tmpy2 = z2
            tmpx3 = y3
            tmpy3 = z3
            tmpx4 = y4
            tmpy4 = z4
        End If
    End If
    
    Call LineCoefficient(tmpx1, tmpy1, tmpx2, tmpy2, LineA1, LineB1, LineC1)
    Call LineCoefficient(tmpx3, tmpy3, tmpx4, tmpy4, LineA2, LineB2, LineC2)
    
    ' Two lines are parallel to each other if their slopes are equal.
    If Abs((LineA1 * LineB2) - (LineA2 * LineB1)) < 0.0001 Then
        MsgBox ("The two lines are parallel or coincident!")   '  /* Parallel or coincident lines /*
        LinesFilletArc = -1
        Exit Function
    End If
    
    midPoint(0) = (tmpx3 + tmpx4) / 2
    midPoint(1) = (tmpy3 + tmpy4) / 2
    If LineToPoint(LineA1, LineB1, LineC1, midPoint, PointToLine1) < 0 Then 'Find distance p1p2 to p3 */
        LinesFilletArc = -1
        Exit Function
    End If
    
    midPoint(0) = (tmpx1 + tmpx2) / 2
    midPoint(1) = (tmpy1 + tmpy2) / 2
    If LineToPoint(LineA2, LineB2, LineC2, midPoint, PointToLine2) < 0 Then ' Find distance p3p4 to p2 */
        LinesFilletArc = -1
        Exit Function
    End If
    
    radius = r
    If (PointToLine1 <= 0) Then
        radius = -radius
    End If
    
    LineConstant1 = LineC1 - radius * Sqr((LineA1 * LineA1) + (LineB1 * LineB1)) ' Line parallel l1 at d */
    
    radius = r

    If (PointToLine2 <= 0) Then
        radius = -radius
    End If

    LineConstant2 = LineC2 - radius * Sqr((LineA2 * LineA2) + (LineB2 * LineB2)) 'Line parallel l2 at d */
    
    SlopeDistance = LineA1 * LineB2 - LineA2 * LineB1
    
    tmpCenterX = (LineConstant2 * LineB1 - LineConstant1 * LineB2) / SlopeDistance    ' Intersect constructed lines */
    tmpCenterY = (LineConstant1 * LineA2 - LineConstant2 * LineA1) / SlopeDistance    ' to find center of arc */
    centerPoint(0) = tmpCenterX
    centerPoint(1) = tmpCenterY
    
    Call Pointperp(LineA1, LineB1, LineC1, centerPoint, NewX1, NewY1)  ' Clip or extend lines as required */
    Call Pointperp(LineA2, LineB2, LineC2, centerPoint, NewX2, NewY2)
        
    dx = tmpx2 - NewX1
    dy = tmpy2 - NewY1
    PointToLine1 = dx * dx + dy * dy
    dx = tmpx2 - tmpx1
    dy = tmpy2 - tmpy1
    PointToLine2 = dx * dx + dy * dy
    If (PointToLine1 >= PointToLine2) Then
        LinesFilletArc = -1
        Exit Function
    End If
        
    dx = tmpx3 - NewX2
    dy = tmpy3 - NewY2
    PointToLine1 = dx * dx + dy * dy
    dx = tmpx3 - tmpx4
    dy = tmpy3 - tmpy4
    PointToLine2 = dx * dx + dy * dy
    If (PointToLine1 >= PointToLine2) Then
        LinesFilletArc = -1
        Exit Function
    End If
    
    Call PlaneCoefficient(x1, y1, z1, x2, y2, z2, x4, y4, z4, PlaneCoefficientA, PlaneCoefficientB, PlaneCoefficientC, PlaneCoefficientD)
   
    If (plane = 1) Then    'YZ
        x2 = (-PlaneCoefficientB * NewX1 - PlaneCoefficientC * NewY1 - PlaneCoefficientD) / PlaneCoefficientA
        y2 = NewX1
        z2 = NewY1
        x3 = (-PlaneCoefficientB * NewX2 - PlaneCoefficientC * NewY2 - PlaneCoefficientD) / PlaneCoefficientA
        y3 = NewX2
        z3 = NewY2
        FilletCenterY = tmpCenterX
        FilletCenterZ = tmpCenterY
        FilletCenterX = Center(y2, x2, z2, y3, x3, z3, tmpCenterX, tmpCenterY, r)
    ElseIf (plane = 2) Then 'XZ
        x2 = NewX1
        y2 = (-PlaneCoefficientA * NewX1 - PlaneCoefficientC * NewY1 - PlaneCoefficientD) / PlaneCoefficientB
        z2 = NewY1
        x3 = NewX2
        y3 = (-PlaneCoefficientA * NewX2 - PlaneCoefficientC * NewY2 - PlaneCoefficientD) / PlaneCoefficientB
        z3 = NewY2
        FilletCenterX = tmpCenterX
        FilletCenterZ = tmpCenterY
        FilletCenterY = Center(x2, y2, z2, x3, y3, z3, tmpCenterX, tmpCenterY, r)
    ElseIf (plane = 3) Then 'XY
        x2 = NewX1
        y2 = NewY1
        z2 = (-PlaneCoefficientA * NewX1 - PlaneCoefficientB * NewY1 - PlaneCoefficientD) / PlaneCoefficientC
        x3 = NewX2
        y3 = NewY2
        z3 = (-PlaneCoefficientA * NewX2 - PlaneCoefficientB * NewY2 - PlaneCoefficientD) / PlaneCoefficientC
        FilletCenterX = tmpCenterX
        FilletCenterY = tmpCenterY
        FilletCenterZ = Center(x2, z2, y2, x3, z3, y3, tmpCenterX, tmpCenterY, r)
    End If

    LinesFilletArc = 0
End Function

Public Sub LineCoefficient(ByVal Pointx1 As Variant, ByVal Pointy1 As Variant, ByVal Pointx2 As Variant, ByVal Pointy2 As Variant, CoefficientA As Variant, CoefficientB As Variant, ConstantC As Variant)

    ' Find the coefficient of the line equation
    ConstantC = (Pointx2 * Pointy1) - (Pointx1 * Pointy2)
    CoefficientA = Pointy2 - Pointy1
    CoefficientB = Pointx1 - Pointx2
End Sub

Public Function LineToPoint(ByVal LineA1 As Double, ByVal LineB1 As Double, ByVal LineC1 As Double, midPoint() As Variant, PointToLineDistance As Variant) As Integer
    
    'Calculate the distance from a point to line.
    Dim d1 As Double

    d1 = Sqr((LineA1 * LineA1) + (LineB1 * LineB1))
    
    If (Abs(d1) < 0.0001) Then
        LineToPoint = -1
        Exit Function
    Else
        PointToLineDistance = (LineA1 * midPoint(0) + LineB1 * midPoint(1) + LineC1) / d1
    End If
    
    LineToPoint = 0
End Function

Public Function Pointperp(ByVal LineA1 As Double, ByVal LineB1 As Double, ByVal LineC1 As Double, p() As Variant, NewX As Variant, NewY As Variant) As Integer
    'Compute x,y position when p(x,y) is perpendicular to Line.
    Dim D, cp As Double
    
    NewX = 0
    NewY = 0
    D = LineA1 * LineA1 + LineB1 * LineB1
    cp = LineA1 * p(1) - LineB1 * p(0)
    If (Abs(D) < 0.0001) Then
        Pointperp = -1
        Exit Function
    Else
        NewX = (-LineA1 * LineC1 - LineB1 * cp) / D
        NewY = (LineA1 * cp - LineB1 * LineC1) / D
    End If
    
    Pointperp = 0
End Function

Public Function Center(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal CenX As Double, ByVal CenZ As Double, ByVal r As Double) As Double
    Dim A, b, C, D As Double
    
    A = (x1 - CenX) * (x1 - CenX) + (z1 - CenZ) * (z1 - CenZ)
    b = (x2 - CenX) * (x2 - CenX) + (z2 - CenZ) * (z2 - CenZ)
    C = y1 * y1 - y2 * y2
    D = 2 * y1 - 2 * y2
    
    If Abs(D) = 0 Then
        Exit Function
    End If
    Center = (A - b + C) / D
End Function

Public Function FilletNormal3D(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal r As Double, ByRef normal() As Variant) As Integer
    'Find the resultant vector of the third point
    
    Dim a1, a2, a3, b1, b2, b3, c1, c2, c3, NormalD As Double
    
    a1 = x3 - x1
    a2 = y3 - y1
    a3 = z3 - z1

    b1 = x3 - x2
    b2 = y3 - y2
    b3 = z3 - z2

    c1 = a1 + b1
    c2 = a2 + b2
    c3 = a3 + b3
    
    If (Abs(c1) < 0.0001 And Abs(c2) < 0.0001 And Abs(c3) < 0.0001) Then
        FilletNormal3D = -1
        Exit Function
    End If
    
    NormalD = Sqr(c1 * c1 + c2 * c2 + c3 * c3)
    normal(0) = c1 / NormalD
    normal(1) = c2 / NormalD
    normal(2) = c3 / NormalD
    
    normal(0) = normal(0) + r
    normal(1) = normal(1) + r
    normal(2) = normal(2) + r
    
    FilletNormal3D = 0
    
End Function

Public Function Dot2(ByVal vectorx1 As Double, ByVal vectory1 As Double, ByVal vectorx2 As Double, ByVal vectory2 As Double, ByRef span As Double) As Integer
    Dim D, t As Double
        D = Sqr(((vectorx1 * vectorx1) + (vectory1 * vectory1)) * ((vectorx2 * vectorx2) + (vectory2 * vectory2)))
        If (Abs(D) < 0.0001) Then
            Dot2 = -1
        Else
            t = (vectorx1 * vectorx2 + vectory1 * vectory2) / D
            If Abs(t) > 1 Then
                Dot2 = -1
            End If
            span = (Atn(-t / Sqr(1 - t * t)) + Atn(1) * 2)
        End If

        Dot2 = 0
End Function

Public Function Cross2(ByVal vectorx1 As Double, ByVal vectory1 As Double, ByVal vectorx2 As Double, ByVal vectory2 As Double) As Double
    Cross2 = (vectorx1 * vectory2 - vectorx2 * vectory1)
End Function

Public Function Rectangle2D3D(ByVal NewX As Double, ByVal NewY As Double, ByVal NewZ As Double, ByVal NewXX As Double, ByVal NewYY As Double, ByVal NewZZ As Double, ByVal NewCenter1 As Double, ByVal NewCenter2 As Double, ByVal NewCenter3 As Double, ByVal r As Double, ByVal Speed As Double)
    Dim Center(0 To 2), normal(0 To 2), Filletnormal(0 To 2), NewPoint(0 To 2), angle(0 To 1), tmpt1(0 To 2), tmpt2(0 To 2), tmpt3(0 To 2)
    Dim vector1(1), vector2(1) As Double
    Dim NewpointX, NewpointY, NewpointZ As Double
    Dim gPROCISION, StartAngle, EndAngle, SpanAngle As Double
    Dim actual_angle_step As Double
    Dim Dir, step_no, i As Integer
    Dim AngleReturn, normalVector As Integer
    Dim direction, signal, same As Boolean
    Dim Angle_step, Result, CurrentAngle As Double
    Dim PreviousX, PreviousY, PreviousZ, ActualX, ActualY, ActualZ As Double
    Dim StringLine As String
        
    gPROCISION = 0.05   '0.02
    step_no = 0
    CurrentAngle = 0
    StartAngle = 0
    EndAngle = 0
    SpanAngle = 0
    
    If (NewZ = NewZZ) And (NewZ = NewCenter3) And (NewZZ = NewCenter3) Then
        vector1(0) = NewX - NewCenter1
        vector1(1) = NewY - NewCenter2
        vector2(0) = NewXX - NewCenter1
        vector2(1) = NewYY - NewCenter2
        
        StartAngle = DetalAngle(NewX - NewCenter1, NewY - NewCenter2)
        If StartAngle < 0 Then
            StartAngle = StartAngle + 2 * (Atn(1) * 4)
        End If
        
        If Dot2(vector1(0), vector1(1), vector2(0), vector2(1), SpanAngle) < 0 Then
            Exit Function
        End If
        
        If (Cross2(vector1(0), vector1(1), vector2(0), vector2(1)) < 0) Then
            SpanAngle = -SpanAngle
        End If
    
        ' SpanAngle < 0 -->CW  (or)  SpanAngle > 0 -->CCW
        If SpanAngle < 0 Then
            Dir = 1
            SpanAngle = SpanAngle * (-1)
        Else
            Dir = 0
        End If
        
        PreviousX = NewX
        PreviousY = NewY
        PreviousZ = NewZ
        CurrentAngle = StartAngle
    
        If (gPROCISION / r > 1) Then
            Exit Function
        End If
        
        Result = 1 - gPROCISION / r
        Angle_step = 2 * (Atn(-Result / Sqr(1 - Result * Result)) + Atn(1) * 2)
    
        step_no = CInt(SpanAngle / Angle_step) + 1
        
        If (step_no > 450) Then     'Just define
            step_no = 450
        End If
        
        If (step_no = 0) Then
            step_no = Abs(CInt(SpanAngle / 0.05))
            actual_angle_step = SpanAngle / step_no
        Else
            actual_angle_step = SpanAngle / step_no
        End If
        
        For i = 1 To step_no
            If Dir = 1 Then
                CurrentAngle = CurrentAngle - actual_angle_step
            Else
                CurrentAngle = CurrentAngle + actual_angle_step
            End If
        
            ActualX = Cos(CurrentAngle) * r
            ActualY = Sin(CurrentAngle) * r
        
            ActualX = ActualX + NewCenter1
            ActualY = ActualY + NewCenter2
        
            PreviousX = Format(PreviousX, "####0.000")
            PreviousY = Format(PreviousY, "####0.000")
            PreviousZ = Format(PreviousZ, "####0.000")
            
            ActualX = Format(ActualX, "####0.000")
            ActualY = Format(ActualY, "####0.000")
        
            If i >= 1 And i <= step_no Then
                If RepeatPattern = True Then
                    StringLine = "line3D(x=" & CLng(PreviousX * 1000) & ", y=" & CLng(PreviousY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)"
                    StringLine = StringLine & vbNewLine
                    ReadRepeatString = ReadRepeatString & StringLine
                Else
                    A.writeline ("line3D(x=" & CLng(PreviousX * 1000) & ", y=" & CLng(PreviousY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)")
                End If
            End If
        
            PreviousX = ActualX
            PreviousY = ActualY
        Next i
        
        If RepeatPattern = True Then
            StringLine = "line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)"
            StringLine = StringLine & vbNewLine
            ReadRepeatString = ReadRepeatString & StringLine
        Else
            If (Last_Pair = False) Then
                A.writeline ("line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)")
            End If
        End If
        
    Else
        
        direction = False
        signal = False
        same = False
        
        'May be we don't need to find the 3rd point
        'Call FilletNormal3D(NewX, NewY, NewZ, NewXX, NewYY, NewZZ, NewCenter1, NewCenter2, NewCenter3, r, Filletnormal())
        'NewpointX = Filletnormal(0) + NewCenter1
        'NewpointY = Filletnormal(1) + NewCenter2
        'NewpointZ = Filletnormal(2) + NewCenter3
        
        'Use "the center point" to find the normal vector
        Dim normalXX As Double, normalYY As Double, normalZZ As Double
        'normalVector = Normal3D(NewX, NewY, NewZ, NewpointX, NewpointY, NewpointZ, NewXX, NewYY, NewZZ, normal())
        'Test and see the result whether it is the same or not
        normalVector = Normal3D(NewX, NewY, NewZ, NewXX, NewYY, NewZZ, NewCenter1, NewCenter2, NewCenter3, normal())
        'normalVector = Normal3D(NewX, NewY, NewZ, NewCenter1, NewCenter2, NewCenter3, NewXX, NewYY, NewZZ, normal())
    
        If normalVector < 0 Then
            normalXX = 0
            normalYY = 0
            normalZZ = 0
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
            Else
                direction = False
            End If
            'Check it again
            If direction = True Then
                angle1 = -1             'Can put 0 or -1 (result is the same)
                angle2 = 0
                If (NewX >= NewXX) Then
                    signal = False
                    direction = False
                    same = True
                Else
                    signal = True
                End If
            Else
                angle1 = 1             'Can put 0 or 1
                angle2 = 0
            End If
        Else
            angle1 = angle(0)
            angle2 = angle(1)
        End If
        
        'Change the rotation's axis from Y_axis to X_axis
        Call Rotation3D(NewX, NewY, NewZ, angle1, 2, tmpt1())
        Dim tmptX1, tmptY1, tmptZ1 As Double
        tmptX1 = tmpt1(0)
        tmptY1 = tmpt1(1)
        tmptZ1 = tmpt1(2)
        Call Rotation3D(tmptX1, tmptY1, tmptZ1, angle2, 1, tmpt1())
        tmptX1 = tmpt1(0)
        tmptY1 = tmpt1(1)
        tmptZ1 = tmpt1(2)
        
        Call Rotation3D(NewpointX, NewpointY, NewpointZ, angle1, 2, tmpt2())
        Dim tmptX2, tmptY2, tmptZ2 As Double
        tmptX2 = tmpt2(0)
        tmptY2 = tmpt2(1)
        tmptZ2 = tmpt2(2)
        Call Rotation3D(tmptX2, tmptY2, tmptZ2, angle2, 1, tmpt2())
        tmptX2 = tmpt2(0)
        tmptY2 = tmpt2(1)
        tmptZ2 = tmpt2(2)
    
        Call Rotation3D(NewXX, NewYY, NewZZ, angle1, 2, tmpt3())
        Dim tmptX3, tmptY3, tmptZ3 As Double
        tmptX3 = tmpt3(0)
        tmptY3 = tmpt3(1)
        tmptZ3 = tmpt3(2)
        Call Rotation3D(tmptX3, tmptY3, tmptZ3, angle2, 1, tmpt3())
        tmptX3 = tmpt3(0)
        tmptY3 = tmpt3(1)
        tmptZ3 = tmpt3(2)
        
        Call Rotation3D(NewCenter1, NewCenter2, NewCenter3, angle1, 2, Center())
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
        
        StartAngle = DetalAngle(TranX1, TranY1)
        If StartAngle < 0 Then
            StartAngle = StartAngle + 2 * (Atn(1) * 4)
        End If
        
        EndAngle = DetalAngle(TranX3, TranY3)
        If EndAngle < 0 Then
            EndAngle = EndAngle + 2 * (Atn(1) * 4)
        End If
        
        If (signal = False) And (direction = False) Then
            If EndAngle <= StartAngle Then
                SpanAngle = EndAngle + 2 * (Atn(1) * 4) - StartAngle
            Else
                SpanAngle = EndAngle - StartAngle
            End If
            If same = True Then
                SpanAngle = (SpanAngle - 2 * (Atn(1) * 4))
                direction = True
                signal = True
            End If
        Else
            If EndAngle <= StartAngle Then
                SpanAngle = EndAngle - StartAngle
            Else
                SpanAngle = EndAngle + 2 * (Atn(1) * 4) - StartAngle
            End If
        End If
        
        If (gPROCISION / r > 1) Then
            Exit Function
        End If
    
        Result = 1 - gPROCISION / r
        Angle_step = 2 * (Atn(-Result / Sqr(1 - Result * Result)) + Atn(1) * 2)
    
        step_no = CInt(SpanAngle / Angle_step) + 1
        If (direction = True) And (signal = True) Then
            step_no = step_no * (-1)
        End If
        
        If (step_no > 450) Then     'Just define
            step_no = 450
        End If
    
        actual_angle_step = SpanAngle / step_no
        
        Dim X, y, Z As Double
        Z = TranZ1
    
        Dim prev_x, prev_y, prev_z As Double
        prev_x = TranX1
        prev_y = TranY1
        prev_z = TranZ1
    
        CurrentAngle = StartAngle
        
        For i = 1 To step_no
            CurrentAngle = CurrentAngle + actual_angle_step
            X = r * Cos(CurrentAngle)
            y = r * Sin(CurrentAngle)
        
            Dim prev_p(0 To 2), p(0 To 2)
            PreviousX = prev_x
            PreviousY = prev_y
            PreviousZ = prev_z
            ActualX = X
            ActualY = y
            ActualZ = Z
           
            Call Rotation3D(PreviousX, PreviousY, PreviousZ, -angle2, 1, prev_p())
            PreviousX = prev_p(0)
            PreviousY = prev_p(1)
            PreviousZ = prev_p(2)
            Call Rotation3D(PreviousX, PreviousY, PreviousZ, -angle1, 2, prev_p())
            PreviousX = prev_p(0)
            PreviousY = prev_p(1)
            PreviousZ = prev_p(2)
        
            Call Rotation3D(ActualX, ActualY, ActualZ, -angle2, 1, p())
            ActualX = p(0)
            ActualY = p(1)
            ActualZ = p(2)
            Call Rotation3D(ActualX, ActualY, ActualZ, -angle1, 2, p())
            ActualX = p(0)
            ActualY = p(1)
            ActualZ = p(2)
       
            Call Translation3D(PreviousX, PreviousY, PreviousZ, NewCenter1, NewCenter2, NewCenter3, prev_p())
            Call Translation3D(ActualX, ActualY, ActualZ, NewCenter1, NewCenter2, NewCenter3, p())
        
            PreviousX = Format(prev_p(0), "####0.000")
            PreviousY = Format(prev_p(1), "####0.000")
            PreviousZ = Format(prev_p(2), "####0.000")
        
            ActualX = Format(p(0), "####0.000")
            ActualY = Format(p(1), "####0.000")
            ActualZ = Format(p(2), "####0.000")
        
            If i >= 1 And i <= step_no Then
                If RepeatPattern = True Then
                    StringLine = "line3D(x=" & CLng(PreviousX * 1000) & ", y=" & CLng(PreviousY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)"
                    StringLine = StringLine & vbNewLine
                    ReadRepeatString = ReadRepeatString & StringLine
                Else
                    A.writeline ("line3D(x=" & CLng(PreviousX * 1000) & ", y=" & CLng(PreviousY * 1000) & ", z=" & CLng(PreviousZ * 1000) & "; sp=" & Speed & "; 1)")
                End If
            End If
        
            prev_x = X
            prev_y = y
            prev_z = Z
    
        Next i
        
        If RepeatPattern = True Then
            StringLine = "line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; sp=" & Speed & "; 1)"
            StringLine = StringLine & vbNewLine
            ReadRepeatString = ReadRepeatString & StringLine
        Else
            A.writeline ("line3D(x=" & CLng(ActualX * 1000) & ", y=" & CLng(ActualY * 1000) & ", z=" & CLng(ActualZ * 1000) & "; sp=" & Speed & "; 1)")
        End If
    End If
End Function

Public Function DispensingSpeed(ByVal line As String) As Double
    Dim words() As String
    Dim char, value As String
    Dim charlenght, i As Integer
    Dim flag As Boolean
    
    char = ""
    value = ""
    flag = False
    words() = Split(line, ";")
    charlenght = Len(words(1))
    
    For i = 1 To charlenght
        char = Mid(words(1), i, 1)
        If flag = True Then
            value = value & char
        End If
        If char = "=" Then
            flag = True
        End If
    Next i
    
    flag = False
    DispensingSpeed = Val(value)
    
End Function

Public Sub Initialize_Fill_Area()
    Area_P1(0) = 0
    Area_P1(1) = 0
    Area_P1(2) = 0
    Area_P2(0) = 0
    Area_P2(1) = 0
    Area_P2(2) = 0
    Area_P3(0) = 0
    Area_P3(1) = 0
    Area_P3(2) = 0
End Sub

Public Function Three_Points_Collinear(ByVal P1_x As Double, ByVal P1_y As Double, ByVal P1_z As Double, ByVal P2_x As Double, ByVal P2_y As Double, ByVal P2_z As Double, ByVal P3_x As Double, ByVal P3_y As Double, ByVal P3_z As Double) As Boolean
    Dim New_x1 As Double, New_y1 As Double, New_z1 As Double
    Dim New_x2 As Double, New_y2 As Double, New_z2 As Double
    Dim New_x3 As Double, New_y3 As Double, New_z3 As Double
    
    New_x1 = P2_x - P1_x
    New_y1 = P2_y - P1_y
    New_z1 = P2_z - P1_z
    
    New_x2 = P3_x - P1_x
    New_y2 = P3_y - P1_y
    New_z2 = P3_z - P1_z
    
    New_x3 = (New_y1 * New_z2) - (New_z1 * New_y2)
    New_y3 = (New_z1 * New_x2) - (New_x1 * New_z2)
    New_z3 = (New_x1 * New_y2) - (New_y1 * New_x2)
    
    If (New_x3 = 0) And (New_y3 = 0) And (New_z3 = 0) Then
        Three_Points_Collinear = True
    Else
        Three_Points_Collinear = False
    End If
End Function
Public Function LinkLine_Arc(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByRef Speed As Double) As Integer
    Dim r As Double, distance1 As Double, distance2 As Double, actual_dis As Double
    Dim dx As Double, dy As Double, dz As Double, dx2 As Double, dy2 As Double, dz2 As Double
    Dim NewValueX1 As Double, NewValueY1 As Double, NewValueZ1 As Double, NewValueX2 As Double, NewValueY2 As Double, NewValueZ2 As Double, NewCenterX1 As Double, NewCenterY1 As Double, NewCenterZ1 As Double
    Dim angle As Double, newDistance As Double
    Dim StringLine As String
    
    If (Speed < 50) Then
        If (RepeatPattern = False) Then
            A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
        Else
            StringLine = "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
            ReadRepeatString = ReadRepeatString & StringLine & vbNewLine
        End If
        LinkLine_Arc = 0
    Else
        dx = x1 - x2
        dy = y1 - y2
        dz = z1 - z2
    
        dx2 = x2 - x3
        dy2 = y2 - y3
        dz2 = z2 - z3
    
        distance1 = Sqr((dx * dx) + (dy * dy) + (dz * dz))
        distance2 = Sqr((dx2 * dx2) + (dy2 * dy2) + (dz2 * dz2))
        
        actual_dis = FindMaxDistance(distance1, distance2)
        'Old method
        'r = radius(Speed, z1, z2, z3, distance1, distance2)
        r = radius(Speed, actual_dis, distance1, distance2)
        Speed = Speed
        If (distance1 < r) And (distance2 < r) Then
            If (RepeatPattern = False) Then
                A.writeline ("line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
            Else
                StringLine = "line3D(x=" & CLng(x2 * 1000) & ", y=" & CLng(y2 * 1000) & ", z=" & CLng(z2 * 1000) & "; sp=" & Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                ReadRepeatString = ReadRepeatString & StringLine & vbNewLine
            End If
            LinkLine_Arc = 0
        Else
            'Find "Angle" between two lines
            angle = FindAngle(x1, y1, z1, x2, y2, z2, x3, y3, z3)
            
            If (angle >= 90) Then
                If NewPointOnRectangle(x1, y1, z1, x2, y2, z2, x3, y3, z3, r, NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1) < 0 Then
                    LinkLine_Arc = -1
                    Exit Function
                End If
    
                Call Rectangle2D3D(NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1, r, Speed)
            Else
                If NewPointOnRectangle(x1, y1, z1, x2, y2, z2, x3, y3, z3, r, NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1) < 0 Then
                    LinkLine_Arc = -1
                    Exit Function
                End If
                
                newDistance = Magnitude((x2 - NewValueX1), (y2 - NewValueY1), (z2 - NewValueZ1))
                
                If (newDistance <= actual_dis) Then
                    Call Rectangle2D3D(NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1, r, Speed)
                Else
                    Call FindNewPointOnRectangel(x1, y1, z1, x2, y2, z2, x3, y3, z3, actual_dis, NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1, r)
                   
                    Call Rectangle2D3D(NewValueX1, NewValueY1, NewValueZ1, NewValueX2, NewValueY2, NewValueZ2, NewCenterX1, NewCenterY1, NewCenterZ1, r, Speed)
                End If
            End If
            
            LinkLine_Arc = 0
        End If
    End If
End Function
Public Function FindNewPointOnRectangel(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double, ByVal actual_d As Double, ByRef NewX As Double, ByRef NewY As Double, ByRef NewZ As Double, ByRef NewXX As Double, ByRef NewYY As Double, ByRef NewZZ As Double, ByRef NewFilletCenterX As Double, ByRef NewFilletCenterY As Double, ByRef NewFilletCenterZ As Double, ByRef radious As Double)
    
    Dim temp1(0 To 2)  As Double, temp2(0 To 2) As Double           'Temp data for calculation
    Dim temp3(0 To 2) As Double                                     'Temp data for calculation
    Dim d1 As Double, d2 As Double                                  'magminute for two lines
    Dim req_distance As Double
    Dim newPoint1(0 To 2) As Double, newPoint2(0 To 2) As Double    'new coordinates
    
    temp1(0) = (x2 - x1)
    temp1(1) = (y2 - y1)
    temp1(2) = (z2 - z1)
        
    d1 = Sqr((temp1(0) * temp1(0)) + (temp1(1) * temp1(1)) + (temp1(2) * temp1(2)))
        
    temp2(0) = (x2 - x3)
    temp2(1) = (y2 - y3)
    temp2(2) = (z2 - z3)
        
    d2 = Sqr((temp2(0) * temp2(0)) + (temp2(1) * temp2(1)) + (temp2(2) * temp2(2)))
    
    temp1(0) = (temp1(0) / d1)
    temp1(1) = (temp1(1) / d1)
    temp1(2) = (temp1(2) / d1)
            
    temp2(0) = (temp2(0) / d2)
    temp2(1) = (temp2(1) / d2)
    temp2(2) = (temp2(2) / d2)
            
    req_distance = d1 - actual_d
    newPoint1(0) = x1 + (temp1(0) * req_distance)
    newPoint1(1) = y1 + (temp1(1) * req_distance)
    newPoint1(2) = z1 + (temp1(2) * req_distance)
    NewX = Format(newPoint1(0), "###0.000")
    NewY = Format(newPoint1(1), "###0.000")
    NewZ = Format(newPoint1(2), "###0.000")
            
    req_distance = d2 - actual_d
    newPoint2(0) = x3 + (temp2(0) * req_distance)
    newPoint2(1) = y3 + (temp2(1) * req_distance)
    newPoint2(2) = z3 + (temp2(2) * req_distance)
    
    'Befor overwriting, calculate first.
    req_distance = Sqr((NewXX - x2) * (NewXX - x2) + (NewYY - y2) * (NewYY - y2) + (NewZZ - z2) * (NewZZ - z2))
    
    NewXX = Format(newPoint2(0), "###0.000")
    NewYY = Format(newPoint2(1), "###0.000")
    NewZZ = Format(newPoint2(2), "###0.000")
    
    temp1(0) = (x2 - NewFilletCenterX)
    temp1(1) = (y2 - NewFilletCenterY)
    temp1(2) = (z2 - NewFilletCenterZ)
        
    d1 = Sqr((temp1(0) * temp1(0)) + (temp1(1) * temp1(1)) + (temp1(2) * temp1(2)))
        
    temp1(0) = (temp1(0) / d1)
    temp1(1) = (temp1(1) / d1)
    temp1(2) = (temp1(2) / d1)
        
    req_distance = (actual_d / req_distance) * d1
    
    'Calculage new radious
    radious = (req_distance / d1) * radious
    
    req_distance = d1 - req_distance
        
    NewFilletCenterX = NewFilletCenterX + (temp1(0) * req_distance)
    NewFilletCenterY = NewFilletCenterY + (temp1(1) * req_distance)
    NewFilletCenterZ = NewFilletCenterZ + (temp1(2) * req_distance)
End Function

Public Function FindMaxDistance(ByVal d1 As Double, ByVal d2 As Double) As Double
    '''''''''''''''''''''''''
    '   Set Max distance    '
    '''''''''''''''''''''''''
    Dim actual_d As Double              'Actual distance to draw a curve
    
    If (d1 / 2) <= (d2 / 2) Then
        actual_d = d1 / 2
    Else
        actual_d = d2 / 2
    End If
    
    If (actual_d > 35) Then
        FindMaxDistance = 35
    Else
        FindMaxDistance = actual_d
    End If
End Function

Public Function radius(ByRef Speed As Double, ByVal actual_d As Double, ByVal d1 As Double, ByVal d2 As Double) As Double
    '''''''''''''''''''''''''
    '   Used by formula     '
    '''''''''''''''''''''''''
    'Accelearation set 0.2G.
    Dim Temp_Radius As Double
    
    originSpeed = Speed
    
    'For Desktop/ILS/SCS4000
    Temp_Radius = (0.5 * Speed * Speed) / (AccelSpeed / 1000)
    
    If (Temp_Radius > actual_d) Then
        Temp_Radius = actual_d
        Speed = CInt(Sqr(Temp_Radius * (2 * (AccelSpeed / 1000))))
    End If
    
    'MsgBox "Speed and radius are " & Speed & " and " & Temp_Radius & "."
    
    If ((d1 < Temp_Radius) And (d2 < Temp_Radius)) Or ((d1 > Temp_Radius) And (d2 > Temp_Radius)) Then
        radius = Temp_Radius
    Else
        If (Temp_Radius > d1) Then
            If (Temp_Radius < d2) Then
                'Directly copare with the speed will be better.
                If (Speed >= 50) Then
                    'Define a new radius as half of "d1"
                    radius = d1 / 2
                Else
                    'This will make longer radius than d2. We will not put a small arc if the
                    'distance of "d1" is too short.
                    radius = Temp_Radius + d2
                End If
            End If
        Else
            If (Temp_Radius > d2) Then
                If (Speed >= 50) Then
                    'Redefine a new radius according to "d2" (Test the result)
                    'Temp_Radius = d2 * (1 / 3)
                    'Temp_Radius = d2 * (1 / 2)
                    'Temp_Radius = d2 * (2 / 3)
                    Temp_Radius = d2 * (3 / 4)
                Else
                    'Redefine a new radius as half of "d2) and calculate the suitable speed.
                    '(May not be good)
                    Temp_Radius = (d2 / 2)
                End If
                'If the user doesn't want to reduce the speed, just comment this line
                Speed = CInt(Sqr(Temp_Radius * (2 * (AccelSpeed / 1000))))
                'Speed = CInt(Sqr(Temp_Radius * (2 * (AccelSpeed / 250))))
                radius = Temp_Radius
            End If
        End If
    End If
    
    'MsgBox radius
End Function

Public Function FindAngle(ByVal x1 As Double, ByVal y1 As Double, ByVal z1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal z2 As Double, ByVal x3 As Double, ByVal y3 As Double, ByVal z3 As Double) As Double
    Dim vector1(0 To 2) As Double, vector2(0 To 2) As Double
    Dim magVector1 As Double, magVector2 As Double
    Dim cosQ As Double
    
    vector1(0) = (x1 - x2)
    vector1(1) = (y1 - y2)
    vector1(2) = (z1 - z2)
    
    vector2(0) = (x3 - x2)
    vector2(1) = (y3 - y2)
    vector2(2) = (z3 - z2)
    
    'Find magnitude
    magVector1 = Magnitude(vector1(0), vector1(1), vector1(2))
    magVector2 = Magnitude(vector2(0), vector2(1), vector2(2))
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Calculate angle by using                                           '
    '   cos(Q) = (vector1 . vector2)/ mag(vector1) . mag (vector2)      '
    '                                                                   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    cosQ = ((vector1(0) * vector2(0)) + (vector1(1) * vector2(1)) + (vector1(2) * vector2(2))) / (magVector1 * magVector2)
    
    'Find angle.
    cosQ = InvCos(cosQ)
    
    FindAngle = Format(ConvertToDegree(cosQ), "##0.00")
End Function
Public Function ConvertToDegree(ByVal angle As Double) As Double
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    '   Change angle value from degree to radian.       '
    '                                                   '
    '''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim degreeValue As Double
    
    degreeValue = (CDbl(angle) * 180) / 3.14159265358979
    
    ConvertToDegree = degreeValue
    
End Function
Public Function InvCos(ByVal value As Double) As Double
    If Abs(value) <> 1 Then
        InvCos = 1.5707963267949 - Atn(value / Sqr(1 - value * value))
    ElseIf value = -1 Then
        InvCos = 3.14159265358979
    End If
End Function
Public Function Magnitude(ByVal xValue As Double, ByVal yValue As Double, ByVal zValue As Double) As Double
    '''''''''''''''''''''''''''''
    '   Calculate magnitude     '
    '                           '
    '''''''''''''''''''''''''''''
    Magnitude = Sqr((xValue * xValue) + (yValue * yValue) + (zValue * zValue))

End Function


