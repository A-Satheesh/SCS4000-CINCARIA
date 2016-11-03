Attribute VB_Name = "translatorRoutines"
Dim CalculationResult As Integer
Dim words() As String, words2() As String, angle() As String, Pth() As String
Dim x1, x2, x3, y1, y2, y3, z1, z2, z3 As Double
Dim CheckX, CheckY, CheckZ As Double              'Check the same point (LinksArcEnd & LinksArcStart)
Dim CheckXX, CheckYY, CheckZZ As Double           'Check the same point (LinksLinePoint & LinksArcStart)
Dim StartLineX, StartLineY, StartLineZ As Double  'To compare the start point and other points
Dim EndLineX, EndLineY, EndLineZ As Double        'End position
Dim CompareString As String                       'To compare the string
Dim Speed1 As Double

'Save 3 points for rectangle and check whether the user create "No Fill Area"
Dim First_RetC1(0 To 2) As Double, First_RetC2(0 To 2) As Double, First_RetC3(0 To 2) As Double
Dim Pitch As Double, Speed_1 As Double            'For first rectangle
Dim First_Rect As Boolean                         'Flag for first rectangle
Dim No_Fillet_Area As Integer                     'Indicate whether the user choose "No_Fill_Area"
Dim Spray_Valve As Boolean, Arc_Start As Boolean
Dim Rotation_Angle As String
Dim Right_Spray_Valve As Boolean                    'Flag for "LinksLinePoint" to do the procedure as "LineStart" and "LineEnd"
Dim Rect_Spray As Boolean                           'Flag for "rectangle" not to conflit with "LinksLinePoint"
Dim Start_Draw_Arc As Boolean                       'Flag for starting of arc
Dim Previous_State_Rotation As String               'Save the previous rotation becuase we need to know whether the tilting is "ON" or "OFF"

Dim speedForCornerArc1() As String, speedForCornerArc() As String
'Do the purging or spraying for evey part of "Part Array"
Dim Purge_Every_Array As Boolean


Public Sub DoInitializeStateMachine()
    
    PreviousState = UnKnownState
    PrevPrevX = 0
    PrevPrevY = 0
    PrevPrevZ = 0
    PrevX = 0
    PrevY = 0
    PrevZ = 0
    error = False
    
End Sub

Private Sub ChangeState(ByVal State As Integer, ByVal X As Long, ByVal y As Long, ByVal Z As Long)

    PreviousState = State
    PrevPrevX = PrevX
    PrevPrevY = PrevY
    PrevPrevZ = PrevZ
    PrevX = X
    PrevY = y
    PrevZ = Z
    
End Sub

Public Function DoPreProcessParsing(ByVal dataline As String) As Boolean

    Dim errorString As String
    Dim Response As GPMessageConstants
    Dim Parser   As New GOLDParser
    'Dim Done, error As Boolean                                    'Controls when we leave the loop
    Dim Done As Boolean
   
    Dim ReductionNumber As Integer                         'Just for information
    Dim n As Integer, Text As String
            
    If Parser.LoadCompiledGrammar(txtCGTFilePath1) Then
        Parser.OpenTextString (dataline)
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
                Call doPreProcess(Parser.CurrentReduction)
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
    
    If error = True Then
        DoPreProcessParsing = True
        A.Close
    Else
        DoPreProcessParsing = False
    End If
End Function

Private Function Calculate_Spray_Rect() As Boolean
    'sp=10101 => No tilt, 11111 => flag for rectangle
    ReadRepeatString = ReadRepeatString & "dot(x=11111, y=11111, z=11111; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)" & vbNewLine
                        
    RepeatPattern = True
    If (First_RetC1(1) = First_RetC2(1)) And (First_RetC1(0) < First_RetC2(0)) Then
        If (First_RetC2(1) > First_RetC3(1)) Then
            If Not ((y1 = y2) And (x1 < x2) And (y2 > y3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((y1 = y2) And (x1 < x2) And (y2 < y3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
        
        If Spray_CalculationRectangle_Xplus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(1) = First_RetC2(1)) And (First_RetC1(0) > First_RetC2(0)) Then
        If (First_RetC2(1) > First_RetC3(1)) Then
            If Not ((y1 = y2) And (x1 > x2) And (y2 > y3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((y1 = y2) And (x1 > x2) And (y2 < y3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
                            
        If Spray_CalculationRectangle_Xminus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(0) = First_RetC2(0)) And (First_RetC1(1) < First_RetC2(1)) Then
        If (First_RetC2(0) > First_RetC3(0)) Then
            If Not ((x1 = x2) And (y1 < y2) And (x2 > x3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((x1 = x2) And (y1 < y2) And (x2 < x3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
                            
        If Spray_CalculationRectangle_Yplus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(0) = First_RetC2(0)) And (First_RetC1(1) > First_RetC2(1)) Then
        If (First_RetC2(0) > First_RetC3(0)) Then
            If Not ((x1 = x2) And (y1 > y2) And (x2 > x3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((x1 = x2) And (y1 > y2) And (x2 < x3)) Then
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
        
        If Spray_CalculationRectangle_Yminus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Calculate_Spray_Rect = True
            Exit Function
        End If
    Else
        MsgBox "The spray rectangle has to be perpandicular!"
    End If
                    
    Pitch = 0
    First_Rect = False
    No_Fillet_Area = 0
    Calculate_Spray_Rect = False
End Function

Private Sub doPreProcess(ByVal TheReduction As Reduction)

Dim n As Integer
Dim moveHeight, lines, linenum, X, y, Z, xDev, yDev, xcyclenum, ycyclenum, withdrawalSpeed, WithDrawalZ, Speed, dispense, OffSetX, OffSetY, OffSetZ, potDepth, potdepthspeed, potheight, potheightspeed As Long
Dim withdrawaldelay, potdepthdelay, delay As Double

Dim fileLocationTemp As String
'Dim fsRepeat, aRepeat, retstring, readstring, tmpstring
Dim fsRepeat, aRepeat, retstring, tmpstring
Dim tmpoffsetlist, offsetlist As Offsets

'''''''''''''''''''''
'   Line with arc   '
'''''''''''''''''''''
'Flag the track to put arc between lines.
Dim First_Node As Boolean, Second_Node As Boolean
'First, Second and Third Groups mean starting with "Line Start", "Links Line Point" and "Links Arc End".
Dim First_Group As Boolean, Second_Group As Boolean, Third_Group As Boolean
Dim Link_Line_Point1(0 To 2) As Double, Link_Line_Point2(0 To 2) As Double
'To draw arc between two lines, it has 2 type of speeds. First speed will be used in calculation.
Dim First_Speed As Double, Valve_OnOff() As String
    
OffSetX = offsetstk.Top.getOffsetX
OffSetY = offsetstk.Top.getOffsetY
OffSetZ = offsetstk.Top.getOffsetZ


For n = 0 To TheReduction.TokenCount - 1
    Select Case TheReduction.Tokens(n).Kind
        Case SymbolTypeNonterminal
            If (error = False) Then
                Call doPreProcess(TheReduction.Tokens(n).Data)
            Else
                Exit Sub
            End If
        Case Else
            Select Case LCase(TheReduction.Tokens(n).Data)
                Case "repeat"
                    'Patch for multiple repeats
                    
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - X + OffSetX
                        OffSetY = glbOffsetY - y + OffSetY
                        OffSetZ = glbOffsetZ - Z + OffSetZ
                        'glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                    End If
                    glbOffsetX = X
                    glbOffsetY = y
                    glbOffsetZ = Z
                    
                    glbOffsetChg = True
                    
                    fileLocationTemp = TheReduction.Tokens(6).Data.Tokens(0).Data
                    txtDataFilePath = Mid$(fileLocationTemp, 2, Len(fileLocationTemp) - 2)
                    
                    Set fsRepeat = CreateObject("Scripting.FileSystemObject")
                    
                    'Check the old file (XW)
                    If Not fs.FileExists(txtDataFilePath) Then
                        MsgBox "There is no pattern file to do 'Part Array'. Please check it again."
                        error = True
                        Exit Sub
                    End If
                    
                    Set aRepeat = fs.OpenTextFile(txtDataFilePath, 1, False)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    'XW
                    RepeatPattern = False
                    ContinuousLine = False
                    Middle_Pt_OnOff = "1"
                    
                    Do While aRepeat.AtEndOfStream <> True
                        ReadRepeatString = ""
                        Previous_State_Rotation = ""
                        Last_Rec = ""
                        Travel_Speed_Rec = ""
                        
                        'If (Purge_Every_Array = True) Then
                        '    'set the default value, 77777, to do the purging for every part of Part Array(testing XW)
                        '    ReadRepeatString = ReadRepeatString & "dot(x=77777, y=77777, z=77777; 77777, 77777; 77777, 77777; z=77777; sp=77777; 77777.000; 77777.000; z=77777)" & vbNewLine
                        '    Purge_Every_Array = False
                        'End If
                        
                        For lines = 1 To 100
                            'Test
                            'tmpstring = aRepeat.ReadLine & vbNewLine
                            tmpstring = aRepeat.ReadLine
                            
                            words() = Split(tmpstring, "(")
                            Valve_OnOff() = Split(tmpstring, ";")
                            
                            If (tmpstring <> "EndArray") And (tmpstring <> "*** Left-Needle ***") And (tmpstring <> "*** Right-Needle ***") Then
                                words2() = Split(words(1), ";")
                                angle() = Split(tmpstring, "=")
                                speedForCornerArc() = Split(tmpstring, ";")
                                speedForCornerArc1() = Split(speedForCornerArc(1), "=")
                            End If
                
                            Select Case (words(0))
                                Case "reference"
                                    tmpstring = "reference(" & words2(0) & ")"
                                    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                Case "fudicial"
                                    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                Case "dot", "   dot", "dotArray"
                                    tmpstring = "dot(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ";" & words2(7) & ")"
                                    Rotation_Angle = Left(angle(7), Len(angle(7)) - 1)
                
                                    RepeatPattern = True
                                    If (First_Rect = True) Then
                                        If (Calculate_Spray_Rect = True) Then
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        'Turnning_Angle (Rotation_Angle)
                                        'Previous_State_Rotation = Rotation_Angle
                                        Turnning_Angle ("None")
                                        Previous_State_Rotation = "None"
                                    End If
                                    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                
                                Case "lineStart"
                                    tmpstring = "lineStart(" & words2(0) & ";" & words2(1) & ")"
                                    Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                
                                    Call Data(tmpstring)
                                    StartLineX = CalculateX
                                    StartLineY = CalculateY
                                    StartLineZ = CalculateZ
                                    tmpstring = "start(" & words2(0) & ";" & words2(1) & ")"
                                    
                                    RepeatPattern = True
                                    If (First_Rect = True) Then
                                        If (Calculate_Spray_Rect = True) Then
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        Turnning_Angle (Rotation_Angle)
                                        Previous_State_Rotation = Rotation_Angle
                                    End If
                                    
                                    First_Node = True
                                    First_Group = True
                                    
                                    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                Case "arcStart"
                                    tmpstring = "arcStart(" & words2(0) & ";" & words2(1) & ")"
                                    'Rotation_Angle = Left(Angle(4), Len(Angle(4)) - 1)
                                    Arc_Start = True
                                    Data (tmpstring)
                                    x1 = CalculateX
                                    y1 = CalculateY
                                    z1 = CalculateZ
                                    
                                    If (First_Rect = True) Then
                                        If (Calculate_Spray_Rect = True) Then
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        Turnning_Angle ("None")
                                        Previous_State_Rotation = "None"
                                    End If
                                Case "       arcPoint"
                                    tmpstring = "       arcPoint(" & words2(0) & ")"
                
                                    Data (tmpstring)
                                    x2 = CalculateX
                                    y2 = CalculateY
                                    z2 = CalculateZ
                                Case "   linksLinePoint"
                                    tmpstring = "   linksLinePoint(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                                    Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                                    Data (tmpstring)
                                    CheckXX = CalculateX
                                    CheckYY = CalculateY
                                    CheckZZ = CalculateZ
                             
                                    RepeatPattern = True
                                    
                                    If (StartLineX = CheckXX) And (StartLineY = CheckYY) And (StartLineZ = CheckZZ) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
                                    'Else
                                        'StartLineX = 0
                                        'StartLineY = 0
                                        'StartLineZ = 0
                                    End If
                                    
                                    If (CheckX = CheckXX) And (CheckY = CheckYY) And (CheckZ = CheckZZ) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
                                    'Else
                                    '    CheckX = 0
                                    '    CheckY = 0
                                    '    CheckZ = 0
                                    '    tmpstring = "line3D(" & words(1)
                                    '    ReadRepeatString = ReadRepeatString & tmpstring
                                    
                                    End If
                                    
                                    'corner arc
                                    If (First_Node = True) And (Second_Node = True) Then
                                        'Do the calculation
                                        If (First_Group = True) Then
                                            If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf (Second_Group = True) Then
                                            If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                        ElseIf (Third_Group = True) Then
                                            If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                        End If
                                        'If (First_Group = True) Or (Third_Group = True) Then
                                        Link_Line_Point1(0) = Link_Line_Point2(0)
                                        Link_Line_Point1(1) = Link_Line_Point2(1)
                                        Link_Line_Point1(2) = Link_Line_Point2(2)
                    
                                        Link_Line_Point2(0) = CheckXX
                                        Link_Line_Point2(1) = CheckYY
                                        Link_Line_Point2(2) = CheckZZ
                                        First_Group = False
                                        Third_Group = False
                                        Second_Group = True
                                        'End If
                        
                                        CheckX = 0
                                        CheckY = 0
                                        CheckZ = 0
                                        StartLineX = 0
                                        StartLineY = 0
                                        StartLineZ = 0
                                        
                                    Else
                                        If (First_Node = True) Then
                                            Link_Line_Point2(0) = CheckXX
                                            Link_Line_Point2(1) = CheckYY
                                            Link_Line_Point2(2) = CheckZZ
                                            Second_Node = True

                                        End If
                                    End If
                                   
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        'Turnning_Line_Angle (Rotation_Angle)
                                        Previous_State_Rotation = Rotation_Angle
                                    End If
                                    
                                    'First_Speed = CLng(words2(1))
                                    First_Speed = CDbl(speedForCornerArc1(1))
                                    Middle_Pt_OnOff = words2(2)
                                Case "   linksArcStart"
                                    tmpstring = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                                    'Rotation_Angle = Left(Angle(5), Len(Angle(5)) - 1)
                
                                    NoChange = True
                                    Data (tmpstring)
                                    x1 = CalculateX
                                    y1 = CalculateY
                                    z1 = CalculateZ
                                
                                    RepeatPattern = True
                                    'Turnning_Angle (Rotation_Angle)
                                    
                                    If (StartLineX = x1) And (StartLineY = y1) And (StartLineZ = z1) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        StartLineX = 0
'                                        StartLineY = 0
'                                        StartLineZ = 0
                                    End If
                                    
                                    If (CheckX = x1) And (CheckY = y1) And (CheckZ = z1) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        CheckX = 0
'                                        CheckY = 0
'                                        CheckZ = 0
                                    End If
                    
                                    If (CheckXX = x1) And (CheckYY = y1) And (CheckZZ = z1) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        CheckXX = 0
'                                        CheckYY = 0
'                                        CheckZZ = 0
                                    End If
                                    
                                    If (First_Node = True) And (Second_Node = True) Then
                                        'Do the calculation
                                        If (First_Group = True) Then
                                            If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            First_Group = False
                                        ElseIf (Second_Group = True) Then
                                            If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            Second_Group = False
                                        ElseIf (Third_Group = True) Then
                                            If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            Third_Group = False
                                        End If
                                        
                                        Second_Node = False
                                        First_Node = False
                                        End If
                                        
                                    '''''''''''''''''''''''''
                                    '   May not be needed   '
                                    '''''''''''''''''''''''''
                                    ''Make start and stop procedure when drawing a acr
                                    ''If (Spray_Valve = True) Then
                                    '    '10101 means "start and stop"
                                    '    tmpstring = "end3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & words2(1) & "; " & words2(2) & "; 0.000; z=0; sp=10101; z=0)"
                                    '    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                    '    NoChange = False
                                    ''End If
                                    ''For rotation
                                    'If (Spray_Valve = True) Then
                                    '    RepeatPattern = True
                                    '    'Turnning_Angle (Rotation_Angle)
                                    '    'Previous_State_Rotation = Rotation_Angle
                                        
                                    '    Turnning_Angle ("None")
                                    '    Previous_State_Rotation = "None"
                                    'End If
                                Case "   linksArcRestart"
                                    tmpstring = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                                    'Rotation_Angle = Left(Angle(5), Len(Angle(5)) - 1)
                
                                    NoChange3 = True
                                    RepeatPattern = True
                                    Data (tmpstring)
                                    x3 = CalculateX
                                    y3 = CalculateY
                                    z3 = CalculateZ
                              
                                    'Turnning_Angle (Rotation_Angle)
                                  
                                    If (Spray_Valve = True) Then
                                        'No previous node, starting with "Arc"
                                        If (Previous_State_Rotation = "") Then
                                            Previous_State_Rotation = "None"
                                        End If
                                        
                                        If (Previous_State_Rotation <> "None") Then
                                            If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "   linksArcRestart", Previous_State_Rotation) = True Then
                                                A.Close
                                                Exit Sub
                                            End If
                                            NoChange3 = False
                                            Arc_Start = False
                                        Else
                                            CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                        End If
                                    Else
                                        CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                    End If
                                    
                                    If CalculationResult = -1 Then
                                        A.Close             'Close the .txt file first, then unload me
                                        Exit Sub
                                    End If
                                    If NoChange3 = False Then
                                        x1 = x3
                                        y1 = y3
                                        z1 = z3
                                        Change = True
                                    End If
                                Case "   linksArcEnd"
                                    tmpstring = "   linksArcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                                    Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                                    NoChange2 = True
                                    RepeatPattern = True
                                    Data (tmpstring)
                                    x3 = CalculateX
                                    y3 = CalculateY
                                    z3 = CalculateZ
                                    'To check the next position
                                    CheckX = x3
                                    CheckY = y3
                                    CheckZ = z3
                                   
                                    'Turnning_Angle (Rotation_Angle)
                                    
                                    If (Spray_Valve = True) Then
                                        'No previous node, starting with "Arc"
                                        If (Previous_State_Rotation = "") Then
                                            Previous_State_Rotation = "None"
                                        End If
                                        
                                        If (Previous_State_Rotation <> "None") Then
                                            If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "   linksArcEnd", Previous_State_Rotation) = True Then
                                                A.Close
                                                Exit Sub
                                            End If
                                            Arc_Start = False
                                        Else
                                            CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                        End If
                                    Else
                                        CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                    End If
                                    
                                    If CalculationResult = -1 Then
                                        A.Close
                                        Exit Sub
                                    End If
                                    
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        'Turnning_Line_Angle (Rotation_Angle)
                                        Previous_State_Rotation = Rotation_Angle
                                    End If
                                Case "lineEnd"
                                    tmpstring = "lineEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                                    Data (tmpstring)
                                    EndLineX = CalculateX
                                    EndLineY = CalculateY
                                    EndLineZ = CalculateZ
                                    
                                    If (StartLineX = EndLineX) And (StartLineY = EndLineY) And (StartLineZ = EndLineZ) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        StartLineX = 0
'                                        StartLineY = 0
'                                        StartLineZ = 0
                                    End If
                                    
                                    If (CheckXX = EndLineX) And (CheckYY = EndLineY) And (CheckZZ = EndLineZ) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        CheckXX = 0
'                                        CheckYY = 0
'                                        CheckZZ = 0
                                    End If
                                    
                                    If (CheckX = EndLineX) And (CheckY = EndLineY) And (CheckZ = EndLineZ) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
'                                    Else
'                                        CheckX = 0
'                                        CheckY = 0
'                                        CheckZ = 0
                                        
'                                        tmpstring = "end3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
'                                        ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                    End If
                                    If (First_Node = True) And (Second_Node = True) Then
                                        'Do the calculation
                                        If (First_Group = True) Then
                                            If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, CheckXX, CheckYY, CheckZZ, EndLineX, EndLineY, EndLineZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(CheckXX * 1000) & ", y=" & CLng(CheckYY * 1000) & ", z=" & CLng(CheckZZ * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, CheckXX, CheckYY, CheckZZ, EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            First_Group = False
                                        ElseIf (Second_Group = True) Then
                                            If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            Second_Group = False
                                        ElseIf (Third_Group = True) Then
                                            If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ) = True) Then
                                                tmpstring = "line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")"
                                                ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                            Else
                                                If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                                    A.Close
                                                    Exit Sub
                                                End If
                                            End If
                                            Third_Group = False
                                        End If
                                        
                                        'Doesn't change the travell speed
                                        'tmpstring = "end3D(" & words(1)
                                        'Change the travell speed because the second line is too short
                                        tmpstring = "end3d(x=" & CLng(EndLineX * 1000) & ", y=" & CLng(EndLineY * 1000) & ", z=" & CLng(EndLineZ * 1000) & "; sp=" & First_Speed & "; " & Valve_OnOff(2) & "; " & Valve_OnOff(3) & "; " & Valve_OnOff(4) & "; " & Valve_OnOff(5) & "; " & Valve_OnOff(6) & ")"
                                        ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                        Second_Node = False
                                        First_Node = False
                                        
                                    Else
                                        tmpstring = "end3d(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                                        ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                    End If
                
                                    CheckX = 0
                                    CheckY = 0
                                    CheckZ = 0
                                    CheckXX = 0
                                    CheckYY = 0
                                    CheckZZ = 0
                                    StartLineX = 0
                                    StartLineY = 0
                                    StartLineZ = 0
                                    
                                Case "arcEnd"
                                    tmpstring = "arcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                                    RepeatPattern = True
                                    Data (tmpstring)
                                    x3 = CalculateX
                                    y3 = CalculateY
                                    z3 = CalculateZ
                                    
                                    If (Spray_Valve = True) Then
                                        'No previous node, starting with "Arc"
                                        If (Previous_State_Rotation = "") Then
                                            Previous_State_Rotation = "None"
                                        End If
                                    
                                        If (Previous_State_Rotation <> "None") Then
                                            If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "arcEnd", Previous_State_Rotation) = True Then
                                                A.Close
                                                Exit Sub
                                            End If
                                            Arc_Start = False
                                        Else
                                            CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                        End If
                                    Else
                                        CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                                    End If
                                    
                                    If CalculationResult = -1 Then
                                        A.Close
                                        Exit Sub
                                    End If
                                Case "rectC1"
                                    tmpstring = "rectC1(" & words2(0) & ";" & words2(1) & ")"
                                    
                                    If (words2(3) = "1") Then
                                        No_Fillet_Area = 1
                                    Else
                                        If (First_Rect = True) Then
                                            RepeatPattern = True
                                            If (Calculate_Spray_Rect = True) Then
                                                Exit Sub
                                            End If
                                            First_Rect = False
                                        End If
                                        
                                        No_Fillet_Area = 0
                                    End If
                                    
                                    If (First_Rect = False) Then
                                        Pth() = Split(angle(4), ";")
                                        'Pitch = Left(Angle(4), Len(Angle(4)) - 5)
                                        Pitch = Pth(0)
                                        Pitch = Pitch / 1000
                                    End If
                                    
                                    Data (tmpstring)
                                    x1 = CalculateX
                                    y1 = CalculateY
                                    z1 = CalculateZ
                                    
                                    'For rotation
                                    If (Spray_Valve = True) Then
                                        RepeatPattern = True
                                        Turnning_Angle ("None")
                                        Previous_State_Rotation = "None"
                                    End If
                                Case "   rectC2"
                                    tmpstring = "   rectC2(" & words2(0) & ")"
                                    Data (tmpstring)
                                    x2 = CalculateX
                                    y2 = CalculateY
                                    z2 = CalculateZ
                                    
                                    If (x1 = x2) And (y1 = y2) And (z1 = z3) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
                                    End If
                                    
                                Case "rectC3"
                                    RepeatPattern = True
                                    tmpstring = "rectC3(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                                    Speed = DispensingSpeed(tmpstring)
                                    Data (tmpstring)
                                    x3 = CalculateX
                                    y3 = CalculateY
                                    z3 = CalculateZ
                                    
                                    If (x2 = x3) And (y2 = y3) And (z2 = z3) Then
                                        MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
                                        A.Close
                                        Exit Sub
                                    End If
                            
                                    If (Spray_Valve = True) Then
                                        If (First_Rect = False) Then
                                            First_RetC1(0) = x1
                                            First_RetC1(1) = y1
                                            First_RetC1(2) = z1
                                            First_RetC2(0) = x2
                                            First_RetC2(1) = y2
                                            First_RetC2(2) = z2
                                            First_RetC3(0) = x3
                                            First_RetC3(1) = y3
                                            First_RetC3(2) = z3
                                            Speed_1 = Speed
                                            Travel_Speed_Rec = TravelSpeed
                                            Last_Rec = Last
                                            First_Rect = True
                                            Middle_Pt_OnOff = words2(2)
                                        Else
                                            Area_P1(0) = x1
                                            Area_P1(1) = y1
                                            Area_P1(2) = z1
                                            Area_P2(0) = x2
                                            Area_P2(1) = y2
                                            Area_P2(2) = z2
                                            Area_P3(0) = x3
                                            Area_P3(1) = y3
                                            Area_P3(2) = z3
                        
                                            'Middle_Pt_OnOff = words2(2)
                                            RepeatPattern = True
                                            If (Calculate_Spray_Rect = True) Then
                                                Exit Sub
                                            End If
                        
                                        End If
                                    Else
                                        Middle_Pt_OnOff = words2(2)
                                        If CalculationRectangle(x1, y1, z1, x2, y2, z2, x3, y3, z3, Speed) < 0 Then
                                        'If CalculationRectangle_XW(x1, y1, z1, x2, y2, z2, x3, y3, z3, Speed) < 0 Then
                                            A.Close
                                            Exit Sub
                                        End If
                                    End If
                                Case "repeat"
                                    '??????
                                    If (First_Rect = True) Then
                                        RepeatPattern = True
                                        'If CalculationRectangle(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1) < 0 Then
                                        If CalculationRectangle_XW(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1) < 0 Then
                                            A.Close
                                            Exit Sub
                                        End If
                                        First_Rect = False
                                    End If
                                    
                                    ReadRepeatString = ReadRepeatString & tmpstring & vbNewLine
                                Case "*** Left-Needle ***"
                                    If (First_Rect = True) Then
                                        RepeatPattern = True
                                        If (Calculate_Spray_Rect = True) Then
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'set the default value, 66666, for leftNeedle becuase of parser (testing XW)
                                    ReadRepeatString = ReadRepeatString & "dot(x=66666, y=66666, z=66666; 66666, 66666; 66666, 66666; z=66666; sp=66666; 66666.000; 66666.000; z=66666)" & vbNewLine
                                    'Spray_Valve = True
                                    Spray_Valve = False
                                    leftside = True
                                Case "*** Right-Needle ***"
                                    If (First_Rect = True) Then
                                        RepeatPattern = True
                                        If (Calculate_Spray_Rect = True) Then
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    'set the default value, 55555, for rightNeedle becuase of parser (testing XW)
                                    ReadRepeatString = ReadRepeatString & "dot(x=55555, y=55555, z=55555; 55555, 55555; 55555, 55555; z=55555; sp=55555; 55555.000; 55555.000; z=55555)" & vbNewLine
                                    'Spray_Valve = False
                                    Spray_Valve = True
                                    rightside = True
                            End Select
            
                            If aRepeat.AtEndOfStream Then
                                'Only one rectangle for Spray Valve
                                If (First_Rect = True) Then
                                    RepeatPattern = True
                                    If (Calculate_Spray_Rect = True) Then
                                        Exit Sub
                                    End If
                                End If
                                Exit For
                            End If
                        Next
                        
                        DoPreProcessParsing (ReadRepeatString)
                    Loop
                    
                    
                    'Do While aRepeat.AtEndOfStream <> True
                    'readstring = ""
                    '    For lines = 1 To 100
                    '        tmpstring = aRepeat.ReadLine & vbNewLine
                    '        readstring = readstring & tmpstring
                    '        If aRepeat.AtEndOfStream Then
                    '            Exit For
                    '        End If
                    '    Next
                    'DoPreProcessParsing (readstring)
                    'Loop
                    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                    Set tmpoffsetlist = offsetstk.Top
                    offsetstk.Pop
                    Set tmpoffsetlist = Nothing
                    
                    Purge_Every_Array = True
                Case "linksarcrestart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(4).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(6).Data.Tokens(0).Data
                    A.writeline ("linksArcRestart(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & "; sp=" & Speed & "; " & dispense & ")")
                Case "arc"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    A.writeline ("arc(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ")")
                Case "start"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - X + OffSetX
                        OffSetY = glbOffsetY - y + OffSetY
                        OffSetZ = glbOffsetZ - Z + OffSetZ
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                    End If
                    
                    A.writeline ("start(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; " & Format(delay, "####0.000") & ")")
                Case "arcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - X + OffSetX
                        OffSetY = glbOffsetY - y + OffSetY
                        OffSetZ = glbOffsetZ - Z + OffSetZ
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                    End If
                    
                    A.writeline ("ArcStart(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; " & Format(delay, "####0.000") & ")")
                Case "line3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    'Do check (XW)
                    If (X = -55555) And (y = -55555) And (Z = -55555) Then
                        A.writeline ("line3d(x=-55555, y=55555, z=55555; sp=" & Speed & "; " & dispense & ")")
                    ElseIf (X = -11111) And (y = -11111) And (Z = -11111) Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '   Flag not to do "Start" and "Stop" procedure for "L-Needle"  '
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Just for "New Spray System"
                        A.writeline ("line3d(x=-11111, y=11111, z=11111; sp=" & Speed & "; " & dispense & ")")
                    Else
                        A.writeline ("line3d(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; sp=" & Speed & "; " & dispense & ")")
                    End If
                Case "end3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    moveHeight = TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data
                    A.writeline ("end3d(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; sp=" & Speed & "; " & dispense & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ * (-1) & "; sp=" & withdrawalSpeed & "; z=" & moveHeight * (-1) & ")")
                Case "arcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    moveHeight = TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data
                    A.writeline ("ArcEnd(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; sp=" & Speed & "; " & dispense & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ * (-1) & "; sp=" & withdrawalSpeed & "; z=" & moveHeight * (-1) & ")")
                Case "linksarcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    A.writeline ("linksArcStart(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; sp=" & Speed & "; " & dispense & ")")
                Case "linksarcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    'delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    A.writeline ("linksArcStart(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; sp=" & Speed & "; " & dispense & ")")
                Case "dot"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    withdrawaldelay = TheReduction.Tokens(20).Data.Tokens(0).Data
                    moveHeight = TheReduction.Tokens(22).Data.Tokens(2).Data.Tokens(0).Data
                    If (X = 55555) And (y = 55555) And (Z = 55555) And (xDev = 55555) And (xcyclenum = 55555) _
                       And (yDev = 55555) And (ycyclenum = 55555) Then
                       A.writeline ("dot(x=55555, y=55555, z=55555; 55555, 55555; 55555, 55555; z=55555; sp=55555; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 66666) And (y = 66666) And (Z = 66666) And (xDev = 66666) And (xcyclenum = 66666) _
                       And (yDev = 66666) And (ycyclenum = 66666) Then
                       A.writeline ("dot(x=66666, y=66666, z=66666; 66666, 66666; 66666, 66666; z=66666; sp=66666; 66666.000; 66666.000; z=66666)")
                    ElseIf (xDev = 10) And (xcyclenum = 1) And (yDev = 10) And (ycyclenum = 1) And (Speed = 77777) Then
                        'For part array,
                        A.writeline ("dot(x=0, y=0, z=0; 10, 1; 10, 1; z=0; sp=77777; 0.000; 0.000; z=0)")
                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 36363) Then
                        A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=36363; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 99999) Then
                        A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=99999; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 18181) Then
                        A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=18181; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 27272) Then
                        A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=27272; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 10101) Then
                        'No tilt
                        A.writeline ("dot(x=00000, y=00000, z=00000; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)")
                    ElseIf (X = 11111) And (y = 11111) And (Z = 11111) And (Speed = 11111) Then
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        '   Flag not to do "Start" and "Stop" procedure for "L-Needle"  '
                        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                        'Just for "New Spray System"
                        A.writeline ("dot(x=11111, y=11111, z=11111; 11111, 11111; 11111, 11111; z=11111; sp=11111; 11111.000; 11111.000; z=11111)")
                    ElseIf (X = 11111) And (y = 11111) And (Z = 11111) And (Speed = 10101) Then
                        'sp=10101 => No tilt, 11111 => flag for rectangle
                        A.writeline ("dot(x=11111, y=-11111, z=-11111; 55555, 55555; 55555, 55555; z=-55555; sp=10101; 55555.000; 55555.000; z=-55555)")
                    Else
                        If (glbOffsetChg = True) Then
                            OffSetX = glbOffsetX - X + OffSetX
                            OffSetY = glbOffsetY - y + OffSetY
                            OffSetZ = glbOffsetZ - Z + OffSetZ
                            glbOffsetChg = False
                            Set offsetlist = New Offsets
                            Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                            offsetstk.Push offsetlist
                        End If
                        A.writeline ("dot(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & WithDrawalZ * (-1) & "; sp=" & Speed & "; " & Format(delay, "####0.000") & "; " & Format(withdrawaldelay, "####0.000") & "; z=" & moveHeight * (-1) & ")")
                    End If
                Case "pottype1"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    withdrawaldelay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    moveHeight = TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - X + OffSetX
                        OffSetY = glbOffsetY - y + OffSetY
                        OffSetZ = glbOffsetZ - Z + OffSetZ
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                    End If
                    A.writeline ("pottype1(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & potDepth & "; sp=" & potdepthspeed & "; " & Format(potdepthdelay, "####0.000") & "; z=" & WithDrawalZ * (-1) & "; sp=" & Speed & "; " & Format(withdrawaldelay, "####0.000") & "; z=" & moveHeight * (-1) & ")")
                Case "pottype2"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    potheight = TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data
                    potheightspeed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(28).Data.Tokens(2).Data
                    moveHeight = TheReduction.Tokens(30).Data.Tokens(2).Data.Tokens(0).Data
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - X + OffSetX
                        OffSetY = glbOffsetY - y + OffSetY
                        OffSetZ = glbOffsetZ - Z + OffSetZ
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                    End If
                    A.writeline ("pottype2(x=" & X + OffSetX & ", y=" & (y + OffSetY) * (-1) & ", z=" & (Z + OffSetZ) * (-1) & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & potDepth & "; sp=" & potdepthspeed & "; " & Format(potdepthdelay, "####0.000") & "; z=" & potheight & "; sp=" & potheightspeed & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ * (-1) & "; sp=" & Speed & "; z=" & moveHeight * (-1) & ")")
                Case "fudicial"
                    If (glbOffsetChg = True) Then
                        OffSetX = glbOffsetX - TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                        OffSetY = glbOffsetY - TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                        OffSetZ = 0
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(OffSetX, OffSetY, OffSetZ)
                        offsetstk.Push offsetlist
                        A.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data + OffSetX & ", y=" & (TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data + OffSetY) * (-1) & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data + OffSetX & ", y=" & (TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data + OffSetY) * (-1) & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
                    Else
                        A.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data * (-1) & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data * (-1) & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
                    End If
                    'a.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data + offsetX & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data + offsetY & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data + offsetX & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data + offsetY & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
                    'a.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data & ";" & TheReduction.Tokens(6).Data & ")")
            End Select
    End Select
Next
   
End Sub

Public Function DoTranslateParse(ByVal dataline As String) As Boolean

    Dim errorString As String
    Dim Response As GPMessageConstants
    Dim Parser   As New GOLDParser
    'Dim Done, error As Boolean                                    'Controls when we leave the loop
    Dim Done As Boolean
    
    Dim ReductionNumber As Integer                         'Just for information
    Dim n As Integer, Text As String
            
    If Parser.LoadCompiledGrammar(txtCGTFilePath1) Then
        Parser.OpenTextString (dataline)
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
                Call doTranslate(Parser.CurrentReduction)
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
    
    If error = True Then
        DoTranslateParse = True
        A.Close
    Else
        DoTranslateParse = False
    End If
    
    
    Parser.CloseFile
    Set Parser = Nothing
    
End Function

Private Sub doTranslate(ByVal TheReduction As Reduction)

Dim n As Integer
Dim X, y, Z, xDev, xcyclenum, yDev, ycyclenum, withdrawalSpeed, potDepth, potdepthspeed, potheight, potheightspeed, WithDrawalZ, Speed, dispense As Long
Dim withdrawaldelay, potdepthdelay, delay As Double

For n = 0 To TheReduction.TokenCount - 1
    Select Case TheReduction.Tokens(n).Kind
        Case SymbolTypeNonterminal
            If (error = False) Then
                doTranslate TheReduction.Tokens(n).Data
            Else
                Exit Sub
            End If
        Case Else
            Select Case LCase(TheReduction.Tokens(n).Data)
                Case "fudicial"
                    A.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
                Case "linksarcrestart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(4).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(6).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArc(X, y, Speed, dispense)
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Call ChangeState(linksArcRestartState, X, y, PrevZ)
                Case "arc"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    'ccw = TheReduction.Tokens(4).Data.Tokens(2).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoNothing
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoStopArc
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoStopArc
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoStopArc
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Call ChangeState(ArcState, X, y, PrevZ)
                Case "start"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    'No star top for "Left Needle"
                    'If (Right_Spray_Valve = False) Then
                        No_Start_Stop = True
                    'End If
                    
                    Remove_DispOff = False
                    
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                Call DoApproach(X, y, Z, delay)
                            Else
                                Call DoApproach_XW(X, y, Z, delay)
                            End If
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                Call DoApproach(X, y, Z, delay)
                            Else
                                Call DoApproach_XW(X, y, Z, delay)
                            End If
                        Case PotState
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                Call DoApproach(X, y, Z, delay)
                            Else
                                Call DoApproach_XW(X, y, Z, delay)
                            End If
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                Call DoApproach(X, y, Z, delay)
                            Else
                                Call DoApproach_XW(X, y, Z, delay)
                            End If
                        Case UnKnownState
                            'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
                            
                            'If (systemHomeX + needleOffsetX = CLng(X)) And (systemHomeY + (needleOffsetY * (-1)) = CLng(Y)) And (systemHomeZ <= ActualSystemMoveHeight) Then
                                'To flag the starting node (XW)
                            '    FlagStartingPoint = True
                            'End If
                            
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                Call DoApproach(X, y, Z, delay)
                            Else
                                Call DoApproach_XW(X, y, Z, delay)
                            End If
                    End Select
                    Call ChangeState(StartState, X, y, Z)
                Case "arcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoArcStart(X, y, Z, delay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoArcStart(X, y, Z, delay)
                        Case PotState
                            Call DoArcStart(X, y, Z, delay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoArcStart(X, y, Z, delay)
                        Case UnKnownState
                            Call DoArcStart(X, y, Z, delay)
                    End Select
                    Call ChangeState(ArcStartState, X, y, Z)
                Case "linksarcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    'delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArcStop(X, y, Speed, dispense, 0)
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Call ChangeState(linksArcEndState, X, y, Z)
                Case "line3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    
                    'Do check (XW)
                    If (X = -11111) And (y = 11111) And (Z = 11111) And (Speed = 11111) Then
                        A.writeline ("ptp (x=11111, y=11111, z=11111; sp =" & Speed & ")" & vbNewLine)
                        A.writeline ("contibuffer" & vbNewLine)
      
                        ContiBufferLines = ""
                        SegmentPropertyLines = ""
                        SegmentSeqNum = 1
                        'Flag not to do "Start" and "Stop" procedure for "L-Needle"
                        No_Start_Stop = True
                        Call ChangeState(Line3DState, PrevX, PrevY, PrevZ)
                    ElseIf (Right_Spray_Valve = True) And (Rect_Spray = False) And (Start_Draw_Arc = False) And (No_Start_Stop = False) And (X <> -55555) And (y <> 55555) And (Z <> 55555) Then
                        Call DoStartEnd3D_Spray_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                        Call ChangeState(Line3DState, X, y, Z)
                    ElseIf (Right_Spray_Valve = True) And (X = -55555) And (y = 55555) And (Z = 55555) Then
                            Call DoRotate_Spray_XW(PrevX, PrevY, PrevZ, Speed)
                            Call ChangeState(Line3DState, PrevX, PrevY, PrevZ)
                    Else
                        Select Case PreviousState
                            Case StartState
                                If (dispense = 0) Then
                                    First_3D_Line = True
                                End If
                                'Origin (NYP)
                                If (No_Start_Stop = True) Then
                                    Call DoLine3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                Else
                                    Call DoLine3D_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                End If
                            Case linksArcRestartState
                                Call DoError
                            Case ArcState
                                Call DoError
                            Case ArcStopState
                                Call DoError
                            Case linksArcEndState
                                If (CLng(X) = PrevX And CLng(y) = PrevY And CLng(Z) = PrevZ) Then
                                    Call DoError
                                Else
                                    Call DoStopLine3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                End If
                            Case Line3DState
                                If (CLng(X) = PrevX And CLng(y) = PrevY And CLng(Z) = PrevZ) Then
                                    Call DoError
                                Else
                                    'Origin (NYP)
                                    If (No_Start_Stop = True) Then
                                        
                                        '@$K
                                        If (dispense = 20 Or dispense = 21) Then
                                            Call DoLineEnd3D(X - PrevX, y - PrevY, (Z) - Previous_U, Speed, dispense, 0, 0, 0)
                                            Call DoLine3D(X - PrevX, y - PrevY, ((Z) - Previous_U) * 10, Speed, dispense)
                                            Previous_U = (Z)
                                            X = X + 1
                                            y = y + 1
                                            Z = PrevZ
                                        ElseIf (dispense = 30 Or dispense = 31) Then
                                            Call DoLine3D(X - PrevX, y - PrevY, ((Z) - Previous_U) * 10, Speed, dispense)
                                            Previous_U = (Z)
                                            Z = PrevZ
                                        ElseIf (dispense = 10 Or dispense = 11) Then
                                            Call DoLineEnd3D(X - PrevX, y - PrevY, (Z) - Previous_U, Speed, dispense, 0, 0, 0)
                                            Call DoLine3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                        Else
                                            Call DoLine3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                        End If
                                    Else
                                        Call DoLine3D_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                                    End If
                                End If
                            Case Stop3DState
                                Call DoStopLine3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense)
                            Case End3DState
                                Call DoError
                            Case linksArcStartState
                                Call DoError
                            Case DotState
                                Call DoError
                            Case PotState
                                Call DoError
                            Case ArcStartState
                                Call DoError
                            Case ArcEndState
                                Call DoError
                            Case UnKnownState
                                Call DoError
                        End Select
                        Call ChangeState(Line3DState, X, y, Z)
                        ContinuousLine = True                   'XW
                    End If
                Case "end3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = Z + CLng(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data)
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    SystemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            If (CLng(X) = PrevX And CLng(y) = PrevY And CLng(Z) = PrevZ) Then
                                Call DoError
                            Else
                                'Call DoStartEnd3D(x - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                                Call DoStartEnd3D_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                            End If
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoStopEnd3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case Line3DState
                            'Origin (NYP)
                            If (No_Start_Stop = True) Then
                                If (withdrawalSpeed = 10101) Then
                                    Call DoLineEnd3D_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                                Else
                                    Call DoLineEnd3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                                End If
                                No_Start_Stop = False
                            Else
                                Call DoLineEnd3D_XW(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                            End If
                        Case Stop3DState
                            Call DoStopEnd3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Rect_Spray = False
                    Call ChangeState(End3DState, X, y, Z)
                Case "arcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = Z + CLng(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data)
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    SystemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArcEnd(X, y, Z, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Call ChangeState(ArcEndState, X, y, Z)
                Case "linksarcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoStartStop3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, 0)
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            'Modified to suit comment on 300505
                            'Call DoError
                            Call DoStopStop3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, 0)
                        Case Line3DState
                            Call DoLineStop3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, 0)
                        Case Stop3DState
                            Call DoStopStop3D(X - PrevX, y - PrevY, Z - PrevZ, Speed, dispense, 0)
                        Case End3DState
                            Call DoError
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoError
                        Case PotState
                            Call DoError
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoError
                        Case UnKnownState
                            Call DoError
                    End Select
                    Call ChangeState(linksArcStartState, X, y, Z)
                Case "dot"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    WithDrawalZ = Z + CLng(TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    withdrawaldelay = TheReduction.Tokens(20).Data.Tokens(0).Data
                    systemMoveHeightDotPot = WithDrawalZ + CLng(TheReduction.Tokens(22).Data.Tokens(2).Data.Tokens(0).Data)
                    
                    'Testing (two valves) XW
                    If (X = 55555) And (y = 55555) And (Z = 55555) And (xDev = 55555) And (xcyclenum = 55555) _
                       And (yDev = 55555) And (ycyclenum = 55555) Then
                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        
                        If Right_Spray_Valve = False Then
                            Right_Spray_Valve = True
                            'Right_Spray_Valve = False
                            SystemMoveHeight = 0
                        End If
                    ElseIf (X = 66666) And (y = 66666) And (Z = 66666) And (xDev = 66666) And (xcyclenum = 66666) _
                       And (yDev = 66666) And (ycyclenum = 66666) Then
                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        
                        If Right_Spray_Valve = True Then
                            Right_Spray_Valve = False
                            'Right_Spray_Valve = True
                            SystemMoveHeight = 0
                        End If
                    ElseIf (xDev = 10) And (xcyclenum = 1) And (yDev = 10) And (ycyclenum = 1) And (Speed = 77777) Then
                        'For part array,
                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
                        '@SK
'                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 36363) Then
'                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
'                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
'                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 99999) Then
'                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
'                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
'                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 18181) Then
'                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
'                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
'                    ElseIf (X = 0) And (y = 0) And (Z = 0) And (Speed = 27272) Then
'                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
'                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                    ElseIf (Speed = 10101) Then
                        If (X = 0) And (y = 0) And (Z = 0) Then
                            Rect_Spray = False
                            A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        Else
                            X = 0
                            y = 0
                            Z = 0
                            Rect_Spray = True
                            A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        End If
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        '@$K
                        SystemMoveHeight = convertToPulses(GetStringSetting("EpoxyDispenser", "Setup", "SystemMoveHeight", "0"), Z_axis)
                    ElseIf (X = 0) And (y = 0) And (Z = 0) Then
                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed * (-1) & ")" & vbNewLine)
                        Previous_U = CLng(Speed)
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        
                    ElseIf (X = 11111) And (y = 11111) And (Z = 11111) And (Speed = 11111) Then
                        A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & Speed & ")" & vbNewLine)
                        Call ChangeState(DotState, PrevX, PrevY, PrevZ)
                        'Flag not to do "Start" and "Stop" procedure for "L-Needle"
                        No_Start_Stop = True
                    Else
                        Select Case PreviousState
                            Case StartState
                                Call DoError
                            Case linksArcRestartState
                                Call DoError
                            Case ArcState
                                Call DoError
                            Case ArcStopState
                                Call DoError
                            Case linksArcEndState
                                Call DoError
                            Case Line3DState
                                Call DoError
                            Case Stop3DState
                                Call DoError
                            Case End3DState
                                Call DoDottingXY(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                            Case linksArcStartState
                                Call DoError
                            Case DotState
                                Call DoDottingXY(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                            Case PotState
                                Call DoDottingXY(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                            Case ArcStartState
                                Call DoError
                            Case ArcEndState
                                Call DoDottingXY(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                            Case UnKnownState
                                'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
                            
                                'If (systemHomeX + needleOffsetX = CLng(X)) And (systemHomeY + (needleOffsetY * (-1)) = CLng(Y)) And (systemHomeZ <= ActualSystemMoveHeight) Then
                                'To flag the starting node (XW)
                                '    FlagStartingPoint = True
                                'End If
                                Call DoDottingXY(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                        End Select
                        Call ChangeState(DotState, X + (xDev * xcyclenum), y + (yDev * ycyclenum), Z)
                    End If
                Case "pottype1"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    WithDrawalZ = Z + CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    withdrawaldelay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    systemMoveHeightDotPot = WithDrawalZ + CLng(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoPottingXYType1(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoPottingXYType1(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case PotState
                            Call DoPottingXYType1(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoPottingXYType1(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case UnKnownState
                            'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
                            
                            'If (systemHomeX + needleOffsetX = CLng(X)) And (systemHomeY + (needleOffsetY * (-1)) = CLng(Y)) And (systemHomeZ <= ActualSystemMoveHeight) Then
                                'To flag the starting node (XW)
                            '    FlagStartingPoint = True
                            'End If
                            
                            Call DoPottingXYType1(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                    End Select
                    Call ChangeState(PotState, X + (xDev * xcyclenum), y + (yDev * ycyclenum), Z)
                Case "pottype2"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    'potheight = z + CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    potheight = CLng(Z) - CLng(potDepth) + CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    potheightspeed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    WithDrawalZ = Z + CLng(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(28).Data.Tokens(2).Data
                    systemMoveHeightDotPot = WithDrawalZ + CLng(TheReduction.Tokens(30).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoError
                        Case Stop3DState
                            Call DoError
                        Case End3DState
                            Call DoPottingXYType2(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoPottingXYType2(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case PotState
                            Call DoPottingXYType2(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoPottingXYType2(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case UnKnownState
                            'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
                            
                            'If (systemHomeX + needleOffsetX = CLng(X)) And (systemHomeY + (needleOffsetY * (-1)) = CLng(Y)) And (systemHomeZ <= ActualSystemMoveHeight) Then
                                'To flag the starting node (XW)
                            '    FlagStartingPoint = True
                            'End If
                            
                            Call DoPottingXYType2(X, y, Z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                    End Select
                    Call ChangeState(PotState, X + (xDev * xcyclenum), y + (yDev * ycyclenum), Z)
            End Select
    End Select
Next
   
End Sub

Private Sub DoArc(ByVal X As Long, ByVal y As Long, ByVal Speed As Long, ByVal dispense As Long)

    If (DoArcCalCen(X, y) = True) Then
        ccw = detCCW(X, y)
        ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
        SegmentSeqNum = SegmentSeqNum + 1
    Else
        MsgBox ("Error in Arc co-ords!")
        error = True
        Exit Sub
    End If

End Sub

Private Sub DoArcEnd(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    If (DoArcCalCen(X, y) = True) Then
        ccw = detCCW(X, y)
        ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
        SegmentSeqNum = SegmentSeqNum + 1
    
        ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=1; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
        SegmentSeqNum = 1
        A.writeline (ContiBufferLines)
        A.writeline (SegmentPropertyLines)
        A.writeline ("contiend")
        A.writeline ("dispenseoff")
        A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        A.writeline ("ptp(x=" & X & ", y=" & y & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    
        A.writeline ("ptp(x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")")
    
        ContiBufferLines = ""
        SegmentPropertyLines = ""
    Else
        MsgBox ("Error in Arc co-ords!")
        error = True
        Exit Sub
    End If
    
End Sub

Private Sub DoArcStop(ByVal X As Long, ByVal y As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    If (DoArcCalCen(X, y) = True) Then
        ccw = detCCW(X, y)
        ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
        SegmentSeqNum = SegmentSeqNum + 1
    
        ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
        SegmentSeqNum = 1
        A.writeline (ContiBufferLines)
        A.writeline (SegmentPropertyLines)
        A.writeline ("contiend")
    
        If (delay <> 0) Then
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        End If
    
        ContiBufferLines = ""
        SegmentPropertyLines = ""
    Else
        MsgBox ("Error in Arc co-ords!")
        error = True
        Exit Sub
    End If

End Sub

Private Function DoDottingY(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To ycyclenum
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
        
        If (delay <> 0) Then
            A.writeline ("dispenseon")
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
            A.writeline ("dispenseoff")
        End If
        
        If (withdrawaldelay <> 0) Then
            A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
        End If
        
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
        
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
        
        yTemp = yTemp + yDev
    Next counter

    DoDottingY = yTemp

End Function

Private Function DoDottingX(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To xcyclenum
        'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
        'To reduce the cycle time because the robot always goes up before starting the first dispensing point. (XW)
        'If (FlagStartingPoint = False) Then
        '    A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        'Else
        '    FlagStartingPoint = False
        'End If
    
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
    
        'Testing (XW)
        If (delay <> 0) Then
            A.writeline ("dispenseon")
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
            A.writeline ("dispenseoff")
        End If
    
        If (withdrawaldelay <> 0) Then
            A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
        End If
    
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
    
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
    
        xTemp = xTemp + xDev
    Next counter

    DoDottingX = xTemp

End Function

Private Function DoSingleDot(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double)

    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
    SystemMoveHeight = systemMoveHeightDotPot
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & zSystemTravelSpeed & ")")
    A.writeline ("dispenseon")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    A.writeline ("dispenseoff")
    A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")

End Function

Private Sub DoDottingXY(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double)

    Dim xTemp, yTemp, zTemp, xcounter As Long
    Dim ycounter As Long                      'To match all array patterns (XW)

    xTemp = X
    yTemp = y
    zTemp = Z

    If (xDev = 0 And yDev = 0) Then
        yTemp = DoSingleDot(X, y, Z, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
    ElseIf (yDev = 0) Then
        xTemp = DoDottingX(X, y, Z, xDev, xcyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
    ElseIf (xDev = 0) Then
        yTemp = DoDottingY(X, y, Z, yDev, ycyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
    Else
        'For xcounter = 1 To xcyclenum
        '    yTemp = DoDottingY(xTemp, yTemp, zTemp, yDev, ycyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
        '    yDev = -yDev
        '    yTemp = yTemp + yDev
        '    xTemp = xTemp + xDev
        'Next xcounter
    
        For ycounter = 1 To ycyclenum
            xTemp = DoDottingX(xTemp, yTemp, zTemp, xDev, xcyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
            xDev = -xDev
            yTemp = yTemp - yDev
            xTemp = xTemp + xDev
        Next ycounter
    End If
End Sub

Private Function DoPottingYType1(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To ycyclenum
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp - potDepth & "; sp =" & potdepthspeed & ")")
        
        If (potdepthdelay <> 0) Then
            A.writeline ("dispenseon")
            A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
            A.writeline ("dispenseoff")
        End If
        
        If (withdrawaldelay <> 0) Then
            A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
        End If
        
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
        
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
        
        yTemp = yTemp + yDev
    Next counter

    DoPottingYType1 = yTemp

End Function

Private Function DoPottingXType1(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To xcyclenum
        'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
        'To reduce the cycle time because the robot always goes up before starting the first dispensing point. (XW)
        'If (FlagStartingPoint = False) Then
        '    A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        'Else
        '    FlagStartingPoint = False
        'End If
    
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp - potDepth & "; sp =" & potdepthspeed & ")")
    
        If (potdepthdelay <> 0) Then
            A.writeline ("dispenseon")
            A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
            A.writeline ("dispenseoff")
        End If
    
        If (withdrawaldelay <> 0) Then
            A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
        End If
    
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
    
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
    
        xTemp = xTemp + xDev
    Next counter

    DoPottingXType1 = xTemp

End Function

Private Function DoSinglePotType1(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double)

    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
    SystemMoveHeight = systemMoveHeightDotPot
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp =" & zSystemTravelSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z - potDepth & "; sp =" & potdepthspeed & ")")
    A.writeline ("dispenseon")
    A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    A.writeline ("dispenseoff")
    A.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")

End Function

Private Sub DoPottingXYType1(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal withdrawaldelay As Double)

    Dim xTemp, yTemp, zTemp, xcounter As Long
    Dim ycounter As Long                        'To match all array patterns (XW)

    xTemp = X
    yTemp = y
    zTemp = Z

    If (xDev = 0 And yDev = 0) Then
        yTemp = DoSinglePotType1(X, y, Z, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
    ElseIf (yDev = 0) Then
        xTemp = DoPottingXType1(X, y, Z, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
    ElseIf (xDev = 0) Then
        yTemp = DoPottingYType1(X, y, Z, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
    Else
        'For xcounter = 1 To xcyclenum
        '    yTemp = DoPottingYType1(xTemp, yTemp, zTemp, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
        '    yDev = -yDev
        '    yTemp = yTemp + yDev
        '    xTemp = xTemp + xDev
        'Next xcounter
    
        For ycounter = 1 To ycyclenum
            xTemp = DoPottingXType1(xTemp, yTemp, zTemp, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
            xDev = -xDev
            '"- yDev" is that we want the robot to move our side when we key in +ve value
            yTemp = yTemp - yDev
            xTemp = xTemp + xDev
        Next ycounter
    End If
End Sub

Private Function DoPottingYType2(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To ycyclenum
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp - potDepth & "; sp =" & potdepthspeed & ")")
        
        If (potheight <> 0) Then
            A.writeline ("dispenseon")
        End If
        
        If (potdepthdelay <> 0) Then
            A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
        End If
        
        If (potheight <> 0) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & potheight & "; sp =" & potheightspeed & ")")
            A.writeline ("dispenseoff")
        End If
        
        If (delay <> 0) Then
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        End If
        
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
        
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
        
        yTemp = yTemp + yDev
    Next counter

    DoPottingYType2 = yTemp

End Function

Private Function DoPottingXType2(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

    Dim xTemp, yTemp, zTemp, counter As Long

    xTemp = X
    yTemp = y
    zTemp = Z

    For counter = 1 To xcyclenum
        'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
        'To reduce the cycle time because the robot always goes up before starting the first dispensing point. (XW)
        'If (FlagStartingPoint = False) Then
        '    A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        'Else
        '    FlagStartingPoint = False
        'End If
    
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        SystemMoveHeight = systemMoveHeightDotPot
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp =" & zSystemTravelSpeed & ")")
        A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp - potDepth & "; sp =" & potdepthspeed & ")")
    
        If (potheight <> 0) Then
            A.writeline ("dispenseon")
        End If
    
        If (potdepthdelay <> 0) Then
            A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
        End If
    
        If (potheight <> 0) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & potheight & "; sp =" & potheightspeed & ")")
            A.writeline ("dispenseoff")
        End If
    
        If (delay <> 0) Then
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        End If
    
        If (WithDrawalZ <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
        End If
    
        If (SystemMoveHeight <> zTemp) Then
            A.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
        End If
    
        xTemp = xTemp + xDev
    Next counter

    DoPottingXType2 = xTemp

End Function

Private Function DoSinglePotType2(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z & "; sp = " & zSystemTravelSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & Z - potDepth & "; sp =" & potdepthspeed & ")")
    A.writeline ("dispenseon")
    A.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & potheight & "; sp =" & potheightspeed & ")")
    A.writeline ("dispenseoff")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    A.writeline ("ptp (x=" & X & ", y=" & y & ", z=" & SystemMoveHeight & "; sp =" & xySystemTravelSpeed & ")")

End Function

Private Sub DoPottingXYType2(ByVal X As Long, ByVal y As Long, ByVal Z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    Dim xTemp, yTemp, zTemp, xcounter As Long
    Dim ycounter As Long                        'To match all array patterns (XW)

    xTemp = X
    yTemp = y
    zTemp = Z

    If (xDev = 0 And yDev = 0) Then
        yTemp = DoSinglePotType2(X, y, Z, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
    ElseIf (yDev = 0) Then
        xTemp = DoPottingXType2(X, y, Z, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
    ElseIf (xDev = 0) Then
        yTemp = DoPottingYType2(X, y, Z, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
    Else
        'For xcounter = 1 To xcyclenum
        '    yTemp = DoPottingYType2(xTemp, yTemp, zTemp, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
        '    yDev = -yDev
        '    yTemp = yTemp + yDev
        '    xTemp = xTemp + xDev
        'Next xcounter
    
        For ycounter = 1 To ycyclenum
            xTemp = DoPottingXType2(xTemp, yTemp, zTemp, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
            xDev = -xDev
            yTemp = yTemp - yDev
            xTemp = xTemp + xDev
        Next ycounter
    End If
End Sub

Private Sub DoStopArc()

    A.writeline ("contibuffer")
    SegmentSeqNum = 1

End Sub

Private Sub DoApproach(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal delay As Double)

    'a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & systemMoveHeight & "; sp=400)")
    'a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & systemMoveHeight & "; sp=" & zSystemTravelSpeed & ")")
    
    'If we set "systemMoveHeight" same as "systemHomeZ", this procedure will not be needed. (XW)
    'To reduce the cycle time because the robot always goes up before starting the first dispensing point. (XW)
    'If (FlagStartingPoint = False) Then
    '    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & xySystemTravelSpeed & ")")
    'Else
    '    FlagStartingPoint = False
    'End If
    
    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & xySystemTravelSpeed & ")")
    'A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    
    If (XYU_Interpolate = False) Then
        A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    Else
        A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & -1 * (Z_High_Arc) & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    End If
    
    
    'Put the following procedure in DoStartEnd3D procedure because we change the dispensing procedure of 3D line (XW)
    'If (delay <> 0) Then
    '    A.writeline ("dispenseon")
    '    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    'End If
   
    StartLineDelay = delay
    
    'Remove it because we will not need it if we use "Line" command (XW)
    A.writeline ("contibuffer" & vbNewLine)
      
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    SegmentSeqNum = 1

End Sub

Private Sub DoArcStart(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal delay As Double)

    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & xySystemTravelSpeed & ")")
    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & zSystemTravelSpeed & ")")
    
    If (delay <> 0) Then
        A.writeline ("dispenseon")
        A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    End If
        
End Sub

Private Sub DoLine(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub

Private Sub DoLine3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    'Origin
    'If xCoOrd = PrevX And yCoOrd = PrevY And PrevZ = PrevZ Then
    If xCoOrd = 0 And yCoOrd = 0 And zCoOrd = 0 Then
        MsgBox ("Error in Links Coords. Identical nodes!")
        error = True
    Else
        'For rotation (line) XW.
        'If (xCoOrd = -55555) And (yCoOrd = 55555) And (zCoOrd = 55555) Then
        '    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & Speed & ")" & vbNewLine
        'Else
        
        '@$K
        If (dispense <> 0 And dispense <> 10 And dispense <> 20 And dispense <> 30 And dispense <> 40) Then
            If (Remove_DispOff = False) Then
                If (dispense = 1 Or dispense = 11 Or dispense = 21 Or dispense = 31 Or dispense = 41) Then
                    A.writeline ("dispenseon" & vbNewLine)
                End If
                Remove_DispOff = True
            End If
        End If
        
        If (StartLineDelay <> 0) Then
            A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")")
            StartLineDelay = 0
        End If
        
        
        ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ",z=" & zCoOrd & ")" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
        SegmentSeqNum = SegmentSeqNum + 1

    End If
    
End Sub

Private Sub DoLineStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoLineStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ",z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoLineEnd(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    A.writeline ("dispenseoff")
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
End Sub

Private Sub DoLineEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)
    '@$K
    If (dispense = 10 Or dispense = 11 Or dispense = 20 Or dispense = 21) Then
        SegmentSeqNum = 1
        A.writeline (ContiBufferLines)
        A.writeline (SegmentPropertyLines)
        A.writeline ("contiend")
        A.writeline ("contibuffer" & vbNewLine)
        
        ContiBufferLines = ""
        SegmentPropertyLines = ""
        SegmentSeqNum = 1
        ContinuousLine = False
    Else
        ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
        
        'Testing (XW)
        If (Rect_Spray = False) And (Start_Draw_Arc = False) And (Right_Spray_Valve = True) And (No_Start_Stop = False) Then
            ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=10)" & vbNewLine
            SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=1; 0; 2)" & vbNewLine
        End If
        
        SegmentSeqNum = 1
        
        A.writeline (ContiBufferLines)
        
        'Testing by XW (change 3D line procedure)
        If (ContinuousLine = True) Then
            If (First_3D_Line = True) Then
                First_3D_Line = False
            Else
                If (dispense <> 0) Then
                    A.writeline ("dispenseon")
                End If
            End If
            
            If (StartLineDelay <> 0) Then
                A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")")
                StartLineDelay = 0
            End If
            ContinuousLine = False
        End If
        
        A.writeline (SegmentPropertyLines)
        A.writeline ("contiend")
        
        'Not to close vlave for "LinksArcEnd"
        If (withdrawalSpeed <> 11111) Then
            If (dispense <> 0) Then
                A.writeline ("dispenseoff")
            End If
        
            'To reduce the cycle time
            If (delay <> 0) Then
                A.writeline ("delay(" & Format(delay, "####0.000") & ")")
            End If
        
            'If (WithDrawalZ <> 0) Then
                A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
            'End If
        
            'If (SystemMoveHeight <> 0) Then
                A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
            'End If
        End If
        
        ContiBufferLines = ""
        SegmentPropertyLines = ""
    End If
End Sub

Private Sub DoStartStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoStartStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoStartEnd(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    A.writeline ("dispenseoff")
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
End Sub

Private Sub DoRotate_Spray(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long)
    A.writeline ("ptp (x=00000, y=00000, z=00000; sp=" & Speed & ")" & vbNewLine)
    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & xySystemTravelSpeed & ")")
    A.writeline ("contibuffer")
End Sub

Private Sub DoStartEnd3D_Spray(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)
    If xCoOrd = 0 And yCoOrd = 0 And zCoOrd = 0 Then
        MsgBox ("Error in Links Coords. Identical nodes!")
        error = True
    Else
        If (Start_Draw_Arc = False) Then
            ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
            SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
        End If
        'ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
        'Testing (XW)
        If (Rect_Spray = False) And (Start_Draw_Arc = False) Then
            ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=10)" & vbNewLine
            SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=1; 0; 2)" & vbNewLine
        End If
        
        SegmentSeqNum = 1
        A.writeline (ContiBufferLines)
    
        'Testing by XW (change 3D line procedure)
        If (ContinuousLine = False) Then
            If (dispense <> 0) Then
                A.writeline ("dispenseon")
            End If
        
            If (StartLineDelay <> 0) Then
                A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")")
                StartLineDelay = 0
            End If
        End If
    
        A.writeline (SegmentPropertyLines)
        A.writeline ("contiend")
    
        If (dispense <> 0) Then
            A.writeline ("dispenseoff" & vbNewLine)
        End If

        ContiBufferLines = ""
        SegmentPropertyLines = ""
    End If
End Sub


Private Sub DoStartEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    'ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    'Testing (XW)
    If (Rect_Spray = False) And (Start_Draw_Arc = False) Then
        ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=4)" & vbNewLine
        SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=1; 0; 2)" & vbNewLine
    End If
    
    SegmentSeqNum = 1
    A.writeline (ContiBufferLines)
    
    'Testing by XW (change 3D line procedure)
    If (ContinuousLine = False) Then
        If (dispense <> 0) Then
            A.writeline ("dispenseon")
        End If
        
        If (StartLineDelay <> 0) Then
            A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")")
            StartLineDelay = 0
        End If
    End If
    
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    If (dispense <> 0) Then
        A.writeline ("dispenseoff")
    End If
    
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
        
End Sub

Private Sub DoStopStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    A.writeline ("contibuffer")

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0;" & dispense & "; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
End Sub

Private Sub DoStopStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    A.writeline ("contibuffer")

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3d(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3d(sp=0;" & dispense & "; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
End Sub

Private Sub DoStopLine(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    A.writeline ("contibuffer")
    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub

Private Sub DoStopLine3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    A.writeline ("contibuffer")
    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub

Private Sub DoStopEnd(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline ("contibuffer")
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    A.writeline ("dispenseoff")
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=50)")
        
End Sub

Private Sub DoStopEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    A.writeline ("contibuffer")
    A.writeline (ContiBufferLines)
    A.writeline (SegmentPropertyLines)
    A.writeline ("contiend")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    A.writeline ("dispenseoff")
    
    A.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    
    A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")")
    
End Sub

Private Sub DoError()
    MsgBox ("Error in translation! Invalid token expected.")
    error = True
    Exit Sub
End Sub

Private Sub DoNothing()

End Sub


'''''''''''''''''''''''''''''''''''''''''
'   New procedure for "LinksLinePoint"  '
'''''''''''''''''''''''''''''''''''''''''

Private Sub DoApproach_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal delay As Double)
    
    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & xySystemTravelSpeed & ")")
    A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
        
    StartLineDelay = delay
   
    SegmentSeqNum = 1
    
End Sub

Private Sub DoLine3D_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    If xCoOrd = 0 And yCoOrd = 0 And zCoOrd = 0 Then
        MsgBox ("Error in Links Coords. Identical nodes!")
        error = True
    Else
        If (dispense <> 0) Then
            If (Remove_DispOff = False) Then
                A.writeline ("dispenseon" & vbNewLine)
                Remove_DispOff = True
            End If
        End If
            
        If (ContinuousLine = False) Then
            If (StartLineDelay <> 0) Then
                A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")" & vbNewLine)
                StartLineDelay = 0
            End If
        End If
    
        A.writeline ("line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")")
        A.writeline ("segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine)
    
        'To reduce the cycle time, we will not close the valve
        'If (dispense <> 0) Then
        '    A.writeline ("dispenseoff")
        'End If
        
        'A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & PrevZ + zCoOrd & "; sp=" & xySystemTravelSpeed & ")" & vbNewLine)
        'If (Speed < 180) Then
        '    A.writeline ("delay(0.200)")
        'End If
        SegmentSeqNum = 1
    End If
    
End Sub

Private Sub DoLineEnd3D_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)
    
    'To reduce the cycle time, we will not open the valve again
    'If (dispense <> 0) Then
    '    A.writeline ("dispenseon" & vbNewLine)
    'End If
    
    A.writeline ("line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")")
    A.writeline ("segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine)
    
    'This checking is for "Links Arc Start"
    If (withdrawalSpeed <> 10101) Then
        If (dispense <> 0) Then
            A.writeline ("dispenseoff")
        End If
    
        If (delay <> 0) Then
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        End If
    
        A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
        A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    End If
    
    Remove_DispOff = False
    ContinuousLine = False
End Sub

Private Sub DoStartEnd3D_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)
    
    If (ContinuousLine = False) Then
        If (dispense <> 0) Then
            A.writeline ("dispenseon" & vbNewLine)
        End If
        
        If (StartLineDelay <> 0) Then
            A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")" & vbNewLine)
            StartLineDelay = 0
        End If
    End If
        
    A.writeline ("line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")")
    A.writeline ("line3D(x=0, y=0, z=4)")
    A.writeline ("segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine)
    A.writeline ("segmentproperty3D(sp=1; 0; 2)")
    
    'This checking is for "Links Arc Start"
    If (withdrawalSpeed <> 10101) Then
        If (dispense <> 0) Then
            A.writeline ("dispenseoff")
        End If
    
        If (delay <> 0) Then
            A.writeline ("delay(" & Format(delay, "####0.000") & ")")
        End If
    
        A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
        A.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & SystemMoveHeight & "; sp=" & zSystemTravelSpeed & ")" & vbNewLine)
    End If
    
    Remove_DispOff = False
End Sub

Private Sub DoRotate_Spray_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long)
    
    If (Speed <> 10101) And (Remove_DispOff = True) Then
        A.writeline ("dispenseoff")
    End If
    
    'For the spray system, "10101" may not be used. (Test XW)
    If (Speed <> 10101) Then
        A.writeline ("ptp (x=00000, y=00000, z=00000; sp=" & Speed & ")")
    End If
    'A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & xySystemTravelSpeed & ")" & vbNewLine)
    
    If (Speed <> 10101) And (Remove_DispOff = True) Then
        A.writeline ("dispenseon" & vbNewLine)
    End If
    
    'Remove for start/Stop (09-09-13)
    If (No_Start_Stop = False) Then
        SegmentSeqNum = 1
    End If
End Sub

Private Sub DoStartEnd3D_Spray_XW(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)
    If xCoOrd = 0 And yCoOrd = 0 And zCoOrd = 0 Then
        MsgBox ("Error in Links Coords. Identical nodes!")
        error = True
    Else
        
        If (dispense <> 0) Then
            If (Remove_DispOff = False) Then
                A.writeline ("dispenseon" & vbNewLine)
                Remove_DispOff = True
            End If
        End If
            
        If (ContinuousLine = False) Then
            If (StartLineDelay <> 0) Then
                A.writeline ("delay(" & Format(StartLineDelay, "####0.000") & ")" & vbNewLine)
                StartLineDelay = 0
            End If
        End If
    
        If (Start_Draw_Arc = False) Then
            A.writeline ("line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")")
            A.writeline ("segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine)
        End If
    
        'To reduce the cycle time, we will not close the valve
        'If (dispense <> 0) Then
        '    A.writeline ("dispenseoff")
        'End If
    
        'A.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=" & xySystemTravelSpeed & ")" & vbNewLine)
        'If (Speed < 180) And (Start_Draw_Arc = False) Then
        '    A.writeline ("delay(0.200)")
        'End If
        
        SegmentSeqNum = 1
    End If
End Sub


