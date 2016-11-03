Attribute VB_Name = "translatorRoutines"
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

Private Sub ChangeState(ByVal State As Integer, ByVal X As Long, ByVal Y As Long, ByVal z As Long)

    PreviousState = State
    PrevPrevX = PrevX
    PrevPrevY = PrevY
    PrevPrevZ = PrevZ
    PrevX = X
    PrevY = Y
    PrevZ = z
    
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
        a.Close
    Else
        DoPreProcessParsing = False
    End If
End Function

Private Sub doPreProcess(ByVal TheReduction As Reduction)

Dim n As Integer
Dim moveHeight, lines, linenum, X, Y, z, xDev, yDev, xcyclenum, ycyclenum, withdrawalSpeed, WithDrawalZ, Speed, dispense, offsetX, offsetY, offsetZ, potDepth, potdepthspeed, potheight, potheightspeed As Long
Dim withdrawaldelay, potdepthdelay, delay As Double



Dim fileLocationTemp As String
Dim fsRepeat, aRepeat, retstring, readstring, tmpstring
Dim tmpoffsetlist, offsetlist As Offsets


offsetX = offsetstk.Top.getOffsetX
offsetY = offsetstk.Top.getOffsetY
offsetZ = offsetstk.Top.getOffsetZ


   
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
                    glbOffsetX = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    glbOffsetY = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    glbOffsetZ = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    glbOffsetChg = True
                    fileLocationTemp = TheReduction.Tokens(6).Data.Tokens(0).Data
                    txtDataFilePath = Mid$(fileLocationTemp, 2, Len(fileLocationTemp) - 2)
                    Set fsRepeat = CreateObject("Scripting.FileSystemObject")
                    Set aRepeat = fs.OpenTextFile(txtDataFilePath, 1, False)


                    Do While aRepeat.AtEndOfStream <> True
                    readstring = ""
                        For lines = 1 To 100
                            tmpstring = aRepeat.ReadLine & vbNewLine
                            readstring = readstring & tmpstring
                            If aRepeat.AtEndOfStream Then
                                Exit For
                            End If
                        Next
                    DoPreProcessParsing (readstring)
                    Loop
                    Set tmpoffsetlist = offsetstk.Top
                    offsetstk.Pop
                    Set tmpoffsetlist = Nothing
                Case "linksarcrestart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(4).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(6).Data.Tokens(0).Data
                    a.writeline ("linksArcRestart(x=" & X + offsetX & ", y=" & Y + offsetY & "; sp=" & Speed & "; " & dispense & ")")
                Case "arc"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    a.writeline ("arc(x=" & X + offsetX & ", y=" & Y + offsetY & ")")
                Case "start"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    
                    If (glbOffsetChg = True) Then
                        offsetX = glbOffsetX - X
                        offsetY = glbOffsetY - Y
                        offsetZ = glbOffsetZ - z
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    
                    a.writeline ("start(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; " & Format(delay, "####0.000") & ")")
                Case "arcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(6).Data.Tokens(0).Data
                    
                    If (glbOffsetChg = True) Then
                        offsetX = glbOffsetX - X
                        offsetY = glbOffsetY - Y
                        offsetZ = glbOffsetZ - z
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    
                    a.writeline ("ArcStart(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; " & Format(delay, "####0.000") & ")")
                Case "line3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    a.writeline ("line3d(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; sp=" & Speed & "; " & dispense & ")")
                Case "end3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    moveHeight = TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data
                    a.writeline ("end3d(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; sp=" & Speed & "; " & dispense & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & "; z=" & moveHeight & ")")
                Case "arcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    moveHeight = TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data
                    a.writeline ("ArcEnd(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; sp=" & Speed & "; " & dispense & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & "; z=" & moveHeight & ")")
                Case "linksarcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    a.writeline ("linksArcStart(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; sp=" & Speed & "; " & dispense & ")")
                Case "linksarcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    'delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    a.writeline ("linksArcEnd(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; sp=" & Speed & "; " & dispense & ")")
                Case "dot"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    WithDrawalZ = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    withdrawaldelay = TheReduction.Tokens(20).Data.Tokens(0).Data
                    moveHeight = TheReduction.Tokens(22).Data.Tokens(2).Data.Tokens(0).Data
                    If (glbOffsetChg = True) Then
                        offsetX = glbOffsetX - X
                        offsetY = glbOffsetY - Y
                        offsetZ = glbOffsetZ - z
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    a.writeline ("dot(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & WithDrawalZ & "; sp=" & Speed & "; " & Format(delay, "####0.000") & "; " & Format(withdrawaldelay, "####0.000") & "; z=" & moveHeight & ")")
                Case "pottype1"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
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
                        offsetX = glbOffsetX - X
                        offsetY = glbOffsetY - Y
                        offsetZ = glbOffsetZ - z
                        glbOffsetChg = False
                        offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    a.writeline ("pottype1(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & potDepth & "; sp=" & potdepthspeed & "; " & Format(potdepthdelay, "####0.000") & "; z=" & WithDrawalZ & "; sp=" & Speed & "; " & Format(potdepthdelay, "####0.000") & "; z=" & moveHeight & ")")
                Case "pottype2"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
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
                        offsetX = glbOffsetX - X
                        offsetY = glbOffsetY - Y
                        offsetZ = glbOffsetZ - z
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    a.writeline ("pottype2(x=" & X + offsetX & ", y=" & Y + offsetY & ", z=" & z + offsetZ & "; " & xDev & ", " & xcyclenum & "; " & yDev & ", " & ycyclenum & "; z=" & potDepth & "; sp=" & potdepthspeed & "; " & Format(potdepthdelay, "####0.000") & "; z=" & potheight & "; sp=" & potheightspeed & "; " & Format(delay, "####0.000") & "; z=" & WithDrawalZ & "; sp=" & Speed & "; z=" & moveHeight & ")")
                Case "fudicial"
                    If (glbOffsetChg = True) Then
                        offsetX = glbOffsetX - TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                        offsetY = glbOffsetY - TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                        offsetZ = 0
                        glbOffsetChg = False
                        Set offsetlist = New Offsets
                        Call offsetlist.setOffsets(offsetX, offsetY, offsetZ)
                        offsetstk.Push offsetlist
                    End If
                    
                    a.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data + offsetX & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data + offsetY & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data + offsetX & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data + offsetY & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
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
        a.Close
    Else
        DoTranslateParse = False
    End If
    
    
    Parser.CloseFile
    Set Parser = Nothing
    
End Function


Private Sub doTranslate(ByVal TheReduction As Reduction)

Dim n As Integer
Dim X, Y, z, xDev, xcyclenum, yDev, ycyclenum, withdrawalSpeed, potDepth, potdepthspeed, potheight, potheightspeed, WithDrawalZ, Speed, dispense As Long
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
                    a.writeline ("fudicial(x=" & TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data & "; x=" & TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data & ", y=" & TheReduction.Tokens(4).Data.Tokens(6).Data.Tokens(0).Data & ";" & TheReduction.Tokens(6).Data & "; " & TheReduction.Tokens(8).Data.Tokens(0).Data & ")")
                Case "linksarcrestart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(4).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(6).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArc(X, Y, Speed, dispense)
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
                    Call ChangeState(linksArcRestartState, X, Y, PrevZ)
                Case "arc"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
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
                    Call ChangeState(ArcState, X, Y, PrevZ)
                Case "start"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
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
                            Call DoApproach(X, Y, z, delay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoApproach(X, Y, z, delay)
                        Case PotState
                            Call DoApproach(X, Y, z, delay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoApproach(X, Y, z, delay)
                        Case UnKnownState
                            Call DoApproach(X, Y, z, delay)
                    End Select
                    Call ChangeState(StartState, X, Y, z)
                Case "arcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
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
                            Call DoArcStart(X, Y, z, delay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoArcStart(X, Y, z, delay)
                        Case PotState
                            Call DoArcStart(X, Y, z, delay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoArcStart(X, Y, z, delay)
                        Case UnKnownState
                            Call DoArcStart(X, Y, z, delay)
                    End Select
                    Call ChangeState(ArcStartState, X, Y, z)
                Case "linksarcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    'delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArcStop(X, Y, Speed, dispense, 0)
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
                    Call ChangeState(linksArcEndState, X, Y, z)
                Case "line3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoLine3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense)
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoStopLine3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense)
                        Case Line3DState
                            Call DoLine3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense)
                        Case Stop3DState
                            Call DoStopLine3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense)
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
                    Call ChangeState(Line3DState, X, Y, z)
                Case "end3d"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = z - CLng(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data)
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    systemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            Call DoStartEnd3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoStopEnd3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case Line3DState
                            Call DoLineEnd3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
                        Case Stop3DState
                            Call DoStopEnd3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
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
                    Call ChangeState(End3DState, X, Y, z)
                Case "arcend"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    delay = TheReduction.Tokens(10).Data.Tokens(0).Data
                    WithDrawalZ = z - CLng(TheReduction.Tokens(12).Data.Tokens(2).Data.Tokens(0).Data)
                    withdrawalSpeed = TheReduction.Tokens(14).Data.Tokens(2).Data
                    systemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(16).Data.Tokens(2).Data.Tokens(0).Data)
                    Select Case PreviousState
                        Case StartState
                            Call DoError
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoArcEnd(X, Y, z, Speed, dispense, delay, WithDrawalZ, withdrawalSpeed)
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
                    Call ChangeState(ArcEndState, X, Y, z)
                Case "linksarcstart"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    Speed = TheReduction.Tokens(6).Data.Tokens(2).Data
                    dispense = TheReduction.Tokens(8).Data.Tokens(0).Data
                    Select Case PreviousState
                        Case StartState
                            Call DoStartStop3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, 0)
                        Case linksArcRestartState
                            Call DoError
                        Case ArcState
                            Call DoError
                        Case ArcStopState
                            Call DoError
                        Case linksArcEndState
                            Call DoError
                        Case Line3DState
                            Call DoLineStop3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, 0)
                        Case Stop3DState
                            Call DoStopStop3D(X - PrevX, Y - PrevY, z - PrevZ, Speed, dispense, 0)
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
                    Call ChangeState(linksArcStartState, X, Y, z)
                Case "dot"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    WithDrawalZ = z - CLng(TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    withdrawaldelay = TheReduction.Tokens(20).Data.Tokens(0).Data
                    systemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(22).Data.Tokens(2).Data.Tokens(0).Data)
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
                            Call DoDottingXY(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoDottingXY(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                        Case PotState
                            Call DoDottingXY(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoDottingXY(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                        Case UnKnownState
                            Call DoDottingXY(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, WithDrawalZ, Speed, delay, withdrawaldelay)
                    End Select
                    Call ChangeState(DotState, X + (xDev * xcyclenum), Y + (yDev * ycyclenum), z)
                Case "pottype1"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    WithDrawalZ = z + CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    withdrawaldelay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    systemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data)
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
                            Call DoPottingXYType1(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoPottingXYType1(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case PotState
                            Call DoPottingXYType1(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoPottingXYType1(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                        Case UnKnownState
                            Call DoPottingXYType1(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, Speed, withdrawaldelay)
                    End Select
                    Call ChangeState(PotState, X + (xDev * xcyclenum), Y + (yDev * ycyclenum), z)
                Case "pottype2"
                    X = TheReduction.Tokens(2).Data.Tokens(2).Data.Tokens(0).Data
                    Y = TheReduction.Tokens(2).Data.Tokens(6).Data.Tokens(0).Data
                    z = TheReduction.Tokens(4).Data.Tokens(2).Data.Tokens(0).Data
                    xDev = TheReduction.Tokens(6).Data.Tokens(0).Data
                    xcyclenum = TheReduction.Tokens(8).Data.Tokens(0).Data
                    yDev = TheReduction.Tokens(10).Data.Tokens(0).Data
                    ycyclenum = TheReduction.Tokens(12).Data.Tokens(0).Data
                    potDepth = TheReduction.Tokens(14).Data.Tokens(2).Data.Tokens(0).Data
                    potdepthspeed = TheReduction.Tokens(16).Data.Tokens(2).Data
                    potdepthdelay = TheReduction.Tokens(18).Data.Tokens(0).Data
                    'potheight = z + CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    potheight = CLng(z) + CLng(potDepth) - CLng(TheReduction.Tokens(20).Data.Tokens(2).Data.Tokens(0).Data)
                    potheightspeed = TheReduction.Tokens(22).Data.Tokens(2).Data
                    delay = TheReduction.Tokens(24).Data.Tokens(0).Data
                    WithDrawalZ = z + CLng(TheReduction.Tokens(26).Data.Tokens(2).Data.Tokens(0).Data)
                    Speed = TheReduction.Tokens(28).Data.Tokens(2).Data
                    systemMoveHeight = WithDrawalZ + CLng(TheReduction.Tokens(30).Data.Tokens(2).Data.Tokens(0).Data)
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
                            Call DoPottingXYType2(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case linksArcStartState
                            Call DoError
                        Case DotState
                            Call DoPottingXYType2(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case PotState
                            Call DoPottingXYType2(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case ArcStartState
                            Call DoError
                        Case ArcEndState
                            Call DoPottingXYType2(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                        Case UnKnownState
                            Call DoPottingXYType2(X, Y, z, xDev, xcyclenum, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, Speed)
                    End Select
                    Call ChangeState(PotState, X + (xDev * xcyclenum), Y + (yDev * ycyclenum), z)
            End Select
    End Select
Next
   
End Sub

Private Sub DoArc(ByVal X As Long, ByVal Y As Long, ByVal Speed As Long, ByVal dispense As Long)


If (DoArcCalCen(X, Y) = True) Then
    ccw = detCCW(X, Y)
    ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (Y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    SegmentSeqNum = SegmentSeqNum + 1
Else
    MsgBox ("Error in Arc co-ords!")
    error = True
    Exit Sub
End If

End Sub
Private Sub DoArcEnd(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

If (DoArcCalCen(X, Y) = True) Then
    ccw = detCCW(X, Y)
    ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (Y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    SegmentSeqNum = SegmentSeqNum + 1
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=1; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    
    SegmentSeqNum = 1
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    a.writeline ("ptp(x=" & X & ", y=" & Y & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    

    
    a.writeline ("ptp(x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp=50)")
    
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
      
Else
    MsgBox ("Error in Arc co-ords!")
    error = True
    Exit Sub
End If



End Sub


Private Sub DoArcStop(ByVal X As Long, ByVal Y As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)


If (DoArcCalCen(X, Y) = True) Then
    ccw = detCCW(X, Y)
    ContiBufferLines = ContiBufferLines & "arc(x=" & (X - PrevPrevX) & ", y=" & (Y - PrevPrevY) & "; x=" & (xCen - PrevPrevX) & ", y=" & (yCen - PrevPrevY) & "; ccw=" & ccw & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    SegmentSeqNum = SegmentSeqNum + 1
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    
    SegmentSeqNum = 1
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    If (delay <> 0) Then
        a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    End If
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
      
Else
    MsgBox ("Error in Arc co-ords!")
    error = True
    Exit Sub
End If

End Sub

Private Function DoDottingY(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To ycyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    yTemp = yTemp + yDev
Next counter

DoDottingY = yTemp

End Function
Private Function DoDottingX(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To xcyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    xTemp = xTemp + xDev
Next counter

DoDottingX = xTemp

End Function
Private Function DoSingleDot(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double)

a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 100)")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & z & "; sp = 50)")
a.writeline ("dispenseon")
a.writeline ("delay(" & Format(delay, "####0.000") & ")")
a.writeline ("dispenseoff")
a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 50)")

End Function
Private Sub DoDottingXY(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal delay As Double, ByVal withdrawaldelay As Double)

Dim xTemp, yTemp, zTemp, xcounter As Long

xTemp = X
yTemp = Y
zTemp = z


If (xDev = 0 And yDev = 0) Then
    yTemp = DoSingleDot(X, Y, z, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
ElseIf (yDev = 0) Then
    xTemp = DoDottingX(X, Y, z, xDev, xcyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
ElseIf (xDev = 0) Then
    yTemp = DoDottingY(X, Y, z, yDev, ycyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
Else
    For xcounter = 1 To xcyclenum
        yTemp = DoDottingY(xTemp, yTemp, zTemp, yDev, ycyclenum, WithDrawalZ, withdrawalSpeed, delay, withdrawaldelay)
        yDev = -yDev
        yTemp = yTemp + yDev
        xTemp = xTemp + xDev
    Next xcounter
End If
End Sub
Private Function DoPottingYType1(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To ycyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp + potDepth & "; sp =" & potdepthspeed & ")")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    yTemp = yTemp + yDev
Next counter

DoPottingYType1 = yTemp

End Function
Private Function DoPottingXType1(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To xcyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp + potDepth & "; sp =" & potdepthspeed & ")")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    xTemp = xTemp + xDev
Next counter

DoPottingXType1 = xTemp

End Function
Private Function DoSinglePotType1(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed, ByVal withdrawaldelay As Double)

a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 100)")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & z & "; sp = 50)")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & z + potDepth & "; sp =" & potdepthspeed & ")")
a.writeline ("dispenseon")
a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
a.writeline ("dispenseoff")
a.writeline ("delay(" & Format(withdrawaldelay, "####0.000") & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 50)")

End Function
Private Sub DoPottingXYType1(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long, ByVal withdrawaldelay As Double)

Dim xTemp, yTemp, zTemp, xcounter As Long

xTemp = X
yTemp = Y
zTemp = z


If (xDev = 0 And yDev = 0) Then
    yTemp = DoSinglePotType1(X, Y, z, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
ElseIf (yDev = 0) Then
    xTemp = DoPottingXType1(X, Y, z, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
ElseIf (xDev = 0) Then
    yTemp = DoPottingYType1(X, Y, z, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
Else
    For xcounter = 1 To xcyclenum
        yTemp = DoPottingYType1(xTemp, yTemp, zTemp, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, WithDrawalZ, withdrawalSpeed, withdrawaldelay)
        yDev = -yDev
        yTemp = yTemp + yDev
        xTemp = xTemp + xDev
    Next xcounter
End If
End Sub
Private Function DoPottingYType2(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal yDev As Long, ByVal ycyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To ycyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp + potDepth & "; sp =" & potdepthspeed & ")")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & potheight & "; sp =" & potheightspeed & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    yTemp = yTemp + yDev
Next counter

DoPottingYType2 = yTemp

End Function
Private Function DoPottingXType2(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

Dim xTemp, yTemp, zTemp, counter As Long

xTemp = X
yTemp = Y
zTemp = z

For counter = 1 To xcyclenum
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 100)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp & "; sp = 50)")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & zTemp + potDepth & "; sp =" & potdepthspeed & ")")
    a.writeline ("dispenseon")
    a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & potheight & "; sp =" & potheightspeed & ")")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
    a.writeline ("ptp (x=" & xTemp & ", y=" & yTemp & ", z=" & systemMoveHeight & "; sp = 50)")
    xTemp = xTemp + xDev
Next counter

DoPottingXType2 = xTemp

End Function
Private Function DoSinglePotType2(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long) As Long

a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 100)")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & z & "; sp = 50)")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & z + potDepth & "; sp =" & potdepthspeed & ")")
a.writeline ("dispenseon")
a.writeline ("delay(" & Format(potdepthdelay, "####0.000") & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & potheight & "; sp =" & potheightspeed & ")")
a.writeline ("dispenseoff")
a.writeline ("delay(" & Format(delay, "####0.000") & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & WithDrawalZ & "; sp =" & withdrawalSpeed & ")")
a.writeline ("ptp (x=" & X & ", y=" & Y & ", z=" & systemMoveHeight & "; sp = 50)")

End Function
Private Sub DoPottingXYType2(ByVal X As Long, ByVal Y As Long, ByVal z As Long, ByVal xDev As Long, ByVal xcyclenum As Long, ByVal yDev, ByVal ycyclenum, ByVal potDepth As Long, ByVal potdepthspeed As Long, ByVal potdepthdelay As Double, ByVal potheight As Long, ByVal potheightspeed As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

Dim xTemp, yTemp, zTemp, xcounter As Long

xTemp = X
yTemp = Y
zTemp = z


If (xDev = 0 And yDev = 0) Then
    yTemp = DoSinglePotType2(X, Y, z, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
ElseIf (yDev = 0) Then
    xTemp = DoPottingXType2(X, Y, z, xDev, xcyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
ElseIf (xDev = 0) Then
    yTemp = DoPottingYType2(X, Y, z, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
Else
    For xcounter = 1 To xcyclenum
        yTemp = DoPottingYType2(xTemp, yTemp, zTemp, yDev, ycyclenum, potDepth, potdepthspeed, potdepthdelay, potheight, potheightspeed, delay, WithDrawalZ, withdrawalSpeed)
        yDev = -yDev
        yTemp = yTemp + yDev
        xTemp = xTemp + xDev
    Next xcounter
End If
End Sub


Private Sub DoStopArc()

    a.writeline ("contibuffer")
    SegmentSeqNum = 1

End Sub

Private Sub DoApproach(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal delay As Double)

    a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & systemMoveHeight & "; sp=400)")
    a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=50)")
    
    If (delay <> 0) Then
        a.writeline ("dispenseon")
        a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    End If
        
    
    a.writeline ("contibuffer")
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    SegmentSeqNum = 1

End Sub
Private Sub DoArcStart(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal delay As Double)

    a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & systemMoveHeight & "; sp=400)")
    a.writeline ("ptp(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & "; sp=50)")
    
    If (delay <> 0) Then
        a.writeline ("dispenseon")
        a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    End If
        
End Sub
Private Sub DoLine(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub
Private Sub DoLine3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ",z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub

Private Sub DoLineStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoLineStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ",z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoLineEnd(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("dispenseoff")
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoLineEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("dispenseoff")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    
    
    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStartStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStartStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStartEnd(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    
    a.writeline ("dispenseoff")
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStartEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=1; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    
    a.writeline ("dispenseoff")
    
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    
    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub

Private Sub DoStopStop(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)


    a.writeline ("contibuffer")

    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line(x=0, y=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=0;" & dispense & "; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStopStop3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double)


    a.writeline ("contibuffer")

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3d(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3d(sp=0;" & dispense & "; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    
End Sub
Private Sub DoStopLine(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    a.writeline ("contibuffer")
    ContiBufferLines = ContiBufferLines & "line(x=" & xCoOrd & ", y=" & yCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    SegmentSeqNum = SegmentSeqNum + 1
    
End Sub
Private Sub DoStopLine3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long)

    a.writeline ("contibuffer")
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
    
    a.writeline ("contibuffer")
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    a.writeline ("dispenseoff")
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    
End Sub
Private Sub DoStopEnd3D(ByVal xCoOrd As Long, ByVal yCoOrd As Long, ByVal zCoOrd As Long, ByVal Speed As Long, ByVal dispense As Long, ByVal delay As Double, ByVal WithDrawalZ As Long, ByVal withdrawalSpeed As Long)

    ContiBufferLines = ContiBufferLines & "line3D(x=" & xCoOrd & ", y=" & yCoOrd & ", z=" & zCoOrd & ")" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=" & Speed & "; " & dispense & "; " & SegmentSeqNum & ")" & vbNewLine
    
    ContiBufferLines = ContiBufferLines & "line3D(x=0, y=0, z=0)" & vbNewLine
    SegmentPropertyLines = SegmentPropertyLines & "segmentproperty3D(sp=0; 0; 2)" & vbNewLine
    
    SegmentSeqNum = 1
    
    a.writeline ("contibuffer")
    a.writeline (ContiBufferLines)
    a.writeline (SegmentPropertyLines)
    a.writeline ("contiend")
    
    ContiBufferLines = ""
    SegmentPropertyLines = ""
    
    a.writeline ("dispenseoff")
    
    a.writeline ("delay(" & Format(delay, "####0.000") & ")")
    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & WithDrawalZ & "; sp=" & withdrawalSpeed & ")")
    

    
    a.writeline ("ptp(x=" & PrevX + xCoOrd & ", y=" & PrevY + yCoOrd & ", z=" & systemMoveHeight & "; sp=50)")
    
    
End Sub
Private Sub DoError()

    MsgBox ("Error in translation! Invalid token expected.")
    error = True
    Exit Sub
End Sub

Private Sub DoNothing()

End Sub
