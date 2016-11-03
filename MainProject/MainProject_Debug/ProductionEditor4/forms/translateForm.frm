VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form translateForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Translator Form"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "translateForm.frx":0000
      Top             =   7440
   End
   Begin VB.CommandButton closeButton 
      Caption         =   "Close"
      Height          =   495
      Left            =   10200
      TabIndex        =   1
      Top             =   7440
      Width           =   1575
   End
   Begin VB.TextBox txtMessage 
      Height          =   6735
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   11535
   End
End
Attribute VB_Name = "translateForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CalculationResult As Integer
Dim words() As String, words2() As String, angle() As String, Pth() As String
Dim words3() As String, words4() As String, words5() As String
Dim speedForCornerArc() As String
Dim speedForCornerArc1() As String

Dim x1, x2, x3, y1, y2, y3, z1, z2, z3 As Double
Dim CheckX, CheckY, CheckZ As Double              'Check the same point (LinksArcEnd & LinksArcStart)
Dim CheckXX, CheckYY, CheckZZ As Double           'Check the same point (LinksLinePoint & LinksArcStart)
Dim StartLineX, StartLineY, StartLineZ As Double  'To compare the start point and other points
Dim EndLineX, EndLineY, EndLineZ As Double        'End position
Dim CompareString As String                       'To compare the string
Dim Speed As Double
Dim First_Time_PtArray As Boolean                 'Flage for the first time of part array
Dim Rotation_Done_Check As Boolean

'Save 3 points for rectangle and check whether the user create "No Fill Area"
Dim First_RetC1(0 To 2) As Double, First_RetC2(0 To 2) As Double, First_RetC3(0 To 2) As Double
Dim Pitch As Long, Speed_1 As Double              'For first rectangle
Dim First_Rect As Boolean                         'Flag for first rectangle
Dim No_Fillet_Area As Integer                     'Indicate whether the user choose "No_Fill_Area"
Dim Spray_Valve As Boolean, Arc_Start As Boolean
Dim Rotation_Angle As String
Dim Previous_State_Rotation As String             'Save the previous rotation becuase we need to know whether the tilting is "ON" or "OFF"

'''''''''''''''''''''
'   Line with arc   '
'''''''''''''''''''''
'Flag the track to put arc between lines.
Dim First_Node As Boolean, Second_Node As Boolean
'First, Second and Third Groups mean starting with "Line Start", "Links Line Point" and "Links Arc End".
Dim First_Group As Boolean, Second_Group As Boolean, Third_Group As Boolean
Dim Link_Line_Point1(0 To 2) As Double, Link_Line_Point2(0 To 2) As Double
'To draw arc between two lines, it has 2 type of speeds. First speed will be used in calculation.
Dim First_Speed As Double
Dim Valve_OnOff() As String
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Delay time, and speed will be wrong from "Data function"  '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub startTranslate()
    Dim proceedToNextStep As Boolean
    Dim lines As Integer
    Dim tmpstring As String
    Dim fsPreProcess, aPreProcess, retstring, readstring
    Dim line As String
    
    txtDataFilePath = editorForm.Caption
    txtPreProcessedDataFilePath = App.Path & "\preprocessedpattern.txt"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set A = fs.createtextfile(App.Path & "\TempoaryPattern.txt", True)
    
    Set fsPreProcess = CreateObject("Scripting.FileSystemObject")
    Set aPreProcess = fsPreProcess.OpenTextFile(txtDataFilePath, 1, False)
    
    proceedToNextStep = True
    NoChange = False
    ContinuousLine = False
    line = ""
    CompareString = ""
    Middle_Pt_OnOff = "1"
    Previous_State_Rotation = ""
    Last_Rec = ""
    Travel_Speed_Rec = ""
    First_Time_PtArray = False
    
    Ini_Val_Rec
    RepeatPattern = False
    
    Do While aPreProcess.AtEndOfStream <> True
        For lines = 1 To TotalLine
            line = aPreProcess.ReadLine
            
            words() = Split(line, "(")
            Valve_OnOff() = Split(line, ";")
            
            If (line <> "EndArray") And (line <> "*** Left-Needle ***") And (line <> "*** Right-Needle ***") Then
                words2() = Split(words(1), ";")
                angle() = Split(line, "=")
                speedForCornerArc() = Split(line, ";")
                speedForCornerArc1() = Split(speedForCornerArc(1), "=")
            End If
            
            If (TotalLine = 1) Then
                ' For One Line Checking
                If (words(0) <> "repeat") And (words(0) <> "reference") Then
                    If (words(0) = "lineStart") Or (words(0) = "lineEnd") Or (words(0) = "arcStart") _
                        Or (words(0) = "       arcPoint") Or (words(0) = "arcEnd") Or (words(0) = "   linksLinePoint") _
                        Or (words(0) = "   linksArcStart") Or (words(0) = "   linksArcRestart") Or (words(0) = "   linksArcEnd") _
                        Or (words(0) = "rectC1") Or (words(0) = "   rectC2") Or (words(0) = "rectC3") Then
                        Call CloseForm2(lines)
                        Exit Sub
                    End If
                End If
            Else
                If (words(0) = "lineEnd") Then
                    If ((CompareString <> "lineStart") And (CompareString <> "   linksLinePoint") And (CompareString <> "   linksArcEnd")) Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf (words(0) = "lineStart") Then
                    If ((CompareString <> "reference") And (CompareString <> "dot") And (CompareString <> "   dot") _
                        And (CompareString <> "dotArray") And (CompareString <> "dotPotting") And (CompareString <> "   dotPotting") _
                        And (CompareString <> "dotPottingArray") And (CompareString <> "linePotting") And (CompareString <> "   linePotting") _
                        And (CompareString <> "linePottingArray") And (words(0) = "repeat") And (CompareString <> "lineEnd") And (CompareString <> "arcEnd") And (CompareString <> "rectC3")) Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf ((words(0) = "   linksLinePoint") Or (words(0) = "   linksArcStart")) Then
                    If ((CompareString <> "lineStart") And (CompareString <> "   linksLinePoint") And (CompareString <> "   linksArcEnd")) Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf (words(0) = "       arcPoint") Then
                    If ((CompareString <> "arcStart") And (CompareString <> "   linksArcStart") And (CompareString <> "   linksArcRestart")) Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf ((words(0) = "   linksArcRestart") Or (words(0) = "   linksArcEnd")) Then
                    If (CompareString <> "       arcPoint") Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf (words(0) = "arcEnd") Then
                    If (CompareString <> "       arcPoint") Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf (words(0) = "   rectC2") Then
                    If (CompareString <> "rectC1") Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf (words(0) = "rectC3") Then
                    If (CompareString <> "   rectC2") Then
                        Call CloseForm(lines, words(0))
                        Exit Sub
                    End If
                ElseIf ((words(0) = "reference") Or (words(0) = "dot") Or (words(0) = "   dot") _
                        Or (words(0) = "dotArray") Or (words(0) = "dotPotting") Or (words(0) = "   dotPotting") _
                        Or (words(0) = "dotPottingArray") Or (words(0) = "linePotting") Or (words(0) = "   linePotting") _
                        Or (words(0) = "linePottingArray") Or (words(0) = "repeat")) Then
                        
                        If ((CompareString = "lineStart") Or (CompareString = "arcStart") Or (CompareString = "rectC1")) Then
                            Call CloseForm(lines, words(0))
                            Exit Sub
                        End If
                End If
            End If
            
            Select Case (words(0))
            Case "reference"
                line = "reference(" & words2(0) & ")"
                A.writeline (line)
            Case "fudicial"
                A.writeline (line)
            Case "dot", "   dot", "dotArray"
                line = "dot(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ";" & words2(7) & ")"
                Rotation_Angle = Left(angle(7), Len(angle(7)) - 1)
                
                If (First_Rect = True) Then
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                
                'For rotation
                If (Spray_Valve = True) Then
                    Check_Rotation_Angle ("None")
                    Previous_State_Rotation = "None"
                End If
                
                XYU_Interpolate = False   '@$K
                A.writeline (line)

            Case "lineStart"
                line = "lineStart(" & words2(0) & ";" & words2(1) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                
                Call Data(line)
                StartLineX = CalculateX
                StartLineY = CalculateY
                StartLineZ = CalculateZ
                
                If (First_Rect = True) Then
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                
                'For rotation
                If (Spray_Valve = True) Then
                        Check_Rotation_Angle (Rotation_Angle)
                    Previous_State_Rotation = Rotation_Angle
                End If
                
                '@$K
                If (Rotation_Angle <> "None") Then
                    Previous_U = CLng(Rotation_Angle)
                    Rotation_Done_Check = True
                End If
                
                First_Node = True
                First_Group = True
                
                XYU_Interpolate = False                'only for ARC   '@$K
                line = "start(" & words2(0) & ";" & words2(1) & ")"
                A.writeline (line)
            Case "arcStart"
                'line = "arcStart(" & words2(0) & ";" & words2(1) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                '@$K
                If (Rotation_Angle <> "None") Then     'z value will be replaced with U value
                    words3() = Split(words2(0), ",")
                    line = "arcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & ")"
                    XYU_Interpolate = True
                    words4() = Split(words3(2), "=")
                    Z_High_Arc = CLng(words4(1))       'height of Z_axis when rotating
                    'PreArcDelay = ArcDelay
                Else
                    line = "arcStart(" & words2(0) & ";" & words2(1) & ")"
                    XYU_Interpolate = False
                    'PreArcDelay = ArcDelay
                End If
                
                'Arc_Start = True
                Data (line)
                x1 = CalculateX
                y1 = CalculateY
                z1 = CalculateZ
                
                If (First_Rect = True) Then
                    'PreArcDelay = ArcDelay
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                
                'For rotation
                If (Spray_Valve = True) Then
                    Check_Rotation_Angle (Rotation_Angle)
                    Previous_State_Rotation = Rotation_Angle
                    'Check_Rotation_Angle ("None")
                    'Previous_State_Rotation = "None"
                End If
            Case "       arcPoint"
                'line = "       arcPoint(" & words2(0) & ")"
                Rotation_Angle = Left(angle(4), Len(angle(4)) - 1)
                Previous_State_Rotation = Rotation_Angle
                If (Rotation_Angle <> "None") Then
                    words3() = Split(words2(0), ",")
                    line = "       arcPoint(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ")"
                    Arc_Start = True
                    'PreArcDelay = ArcDelay
                Else
                    line = "       arcPoint(" & words2(0) & ")"
                    Arc_Start = False
                    'PreArcDelay = ArcDelay
                End If
                
                Data (line)
                x2 = CalculateX
                y2 = CalculateY
                z2 = CalculateZ
            Case "   linksLinePoint"
                'line = "   linksLinePoint(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"'removed for corner arc
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                Data (line)
                CheckXX = CalculateX
                CheckYY = CalculateY
                CheckZZ = CalculateZ
                
                If (StartLineX = CheckXX) And (StartLineY = CheckYY) And (StartLineZ = CheckZZ) Then
                    CloseTranslateForm (lines)
                    Exit Sub
'                Else
'                    StartLineX_Zero'removed for corner arc
                End If
                
                If (CheckX = CheckXX) And (CheckY = CheckYY) And (CheckZ = CheckZZ) Then
                    CloseTranslateForm (lines)
                    Exit Sub
                Else
'                    CheckX_Zero'removed for corner arc
                    '@$K_linkline
                    '--To Rotate and Move X and Y at the same time--'
                    If Spray_Valve = True Then
                        If (Rotation_Angle = Previous_State_Rotation) Then 'if the rotation angle is same, no need to do anything
                            words3() = Split(words2(0), ",")
                            If (Rotation_Done_Check = False) Then 'if no rotation previously, normal X,Y,Z interpolation
                                'line = "line3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                                'corner arc
                                If (First_Node = True) And (Second_Node = True) Then
                                    'Do the calculation
                                    If (First_Group = True) Then
                                        If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                            A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                        Else
                                            If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                A.Close
                                                Unload Me
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf (Second_Group = True) Then
                                        If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                            A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                        Else
                                            If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                A.Close
                                                Unload Me
                                                Exit Sub
                                            End If
                                        End If
                                    ElseIf (Third_Group = True) Then
                                        If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                            A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                        Else
                                            If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                                A.Close
                                                Unload Me
                                                Exit Sub
                                            End If
                                        End If
                                    End If
                                
                                    Link_Line_Point1(0) = Link_Line_Point2(0)
                                    Link_Line_Point1(1) = Link_Line_Point2(1)
                                    Link_Line_Point1(2) = Link_Line_Point2(2)
                                    
                                    Link_Line_Point2(0) = CheckXX
                                    Link_Line_Point2(1) = CheckYY
                                    Link_Line_Point2(2) = CheckZZ
                                    First_Group = False
                                    Third_Group = False
                                    Second_Group = True
                                  
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
                            Else                                  'if rotation previously, change from X,Y,U to X,Y Z first
                                If (words2(2) = " 0") Then        ' check dispense ON/OFF, will make use of this bit to determine X,Y,Z interpolation or X,Y,U interpolation
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & words3(2) & ";" & words2(1) & "; 10" & ")"
                                    A.writeline (line)
                                    words5() = Split(words3(0), "=")
                                    words3(0) = "x=" & words5(1) - 1
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & words3(2) & ";" & words2(1) & "; 0" & ")"
                                ElseIf (words2(2) = " 1") Then
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & words3(2) & ";" & words2(1) & "; 11" & ")"
                                    A.writeline (line)
                                    words5() = Split(words3(0), "=")
                                    words3(0) = "x=" & words5(1) - 1
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & words3(2) & ";" & words2(1) & "; 1" & ")"
                                End If
                                A.writeline (line)
                                Rotation_Done_Check = False
                            End If
                        Else                                'if the angle is different from previous  point
                            words3() = Split(words2(0), ",")
                            If (Rotation_Done_Check = False) Then
                                If (words2(2) = " 0") Then  'check dispense ON/OFF, will make use of this bit to determine X,Y,Z interpolation or X,Y,U interpolation
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 20" & ")" ' 20 is to replace 0
                                    A.writeline (line)
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 30" & ")" ' 30 is to write another segment, link line need to have at least 2 segment
                                ElseIf (words2(2) = " 1") Then
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 21" & ")" ' 21 is to replace 1
                                    A.writeline (line)
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 31" & ")" ' 31 is to write another segment, link line need to have at least 2 segment
                                End If
                                Rotation_Done_Check = True
                            Else
                                If (words2(2) = " 0") Then  'check dispense , will make use of this bit to determine X,Y,Z interpolation or X,Y,U interpolation
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 30" & ")"
                                ElseIf (words2(2) = " 1") Then
                                    line = "line3D(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) & ";" & words2(1) & "; 31" & ")"
                                End If
                            End If
                            A.writeline (line)
                        End If
                    Else
                    'corner arc
                        If (First_Node = True) And (Second_Node = True) Then
                            'Do the calculation
                            If (First_Group = True) Then
                                If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                    A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                Else
                                    If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                        A.Close
                                        Unload Me
                                        Exit Sub
                                    End If
                                End If
                            ElseIf (Second_Group = True) Then
                                If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                    A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                Else
                                    If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                        A.Close
                                        Unload Me
                                        Exit Sub
                                    End If
                                End If
                            ElseIf (Third_Group = True) Then
                                If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ) = True) Then
                                    A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                                Else
                                    If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), CheckXX, CheckYY, CheckZZ, First_Speed) < 0) Then
                                        A.Close
                                        Unload Me
                                        Exit Sub
                                    End If
                                End If
                            End If
                        
                            Link_Line_Point1(0) = Link_Line_Point2(0)
                            Link_Line_Point1(1) = Link_Line_Point2(1)
                            Link_Line_Point1(2) = Link_Line_Point2(2)
                            
                            Link_Line_Point2(0) = CheckXX
                            Link_Line_Point2(1) = CheckYY
                            Link_Line_Point2(2) = CheckZZ
                            First_Group = False
                            Third_Group = False
                            Second_Group = True
                          
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
                    End If
                    
                    First_Speed = CDbl(speedForCornerArc1(1))
                    Middle_Pt_OnOff = words2(2)
                        
                    'For rotation
                    If (Spray_Valve = True) Then
                        'Turnning_Line_Angle (Rotation_Angle)
                        Previous_State_Rotation = Rotation_Angle
                    End If
                    'A.writeline (line)
                End If
            Case "   linksArcStart"
                line = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                Previous_State_Rotation = Rotation_Angle
                
                If (Rotation_Angle <> "None") Then ' z value will be replaced with U value
                    words3() = Split(words2(0), ",")
                    
                    If (words2(2) = " 0") Then
                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 40" & ")"
                    ElseIf (words2(2) = " 1") Then
'                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 41" & ")"
                    Else
                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 40" & ")"
                    End If
                    
                    XYU_Interpolate = True
                    words4() = Split(words3(2), "=")
                    'PreArcDelay = ArcDelay
                Else
                    If (First_Node = True) And (Second_Node = True) Then
                        'Do the calculation
                        If (First_Group = True) Then
                            If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                            Else
                                If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                    A.Close
                                    Unload Me
                                    Exit Sub
                                End If
                            End If
                            First_Group = False
                        ElseIf (Second_Group = True) Then
                            If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                            Else
                                If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                    A.Close
                                    Unload Me
                                    Exit Sub
                                End If
                            End If
                            Second_Group = False
                        ElseIf (Third_Group = True) Then
                            If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1) = True) Then
                                A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                            Else
                                If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), x1, y1, z1, First_Speed) < 0) Then
                                    A.Close
                                    Unload Me
                                    Exit Sub
                                End If
                            End If
                            Third_Group = False
                        End If
            
                        Second_Node = False
                        First_Node = False
                    End If
                
                    'line = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                    XYU_Interpolate = False
                    'PreArcDelay = ArcDelay
                End If
                
                NoChange = True
                Data (line)
                x1 = CalculateX
                y1 = CalculateY
                z1 = CalculateZ
                
                'Check_Rotation_Angle (Rotation_Angle)
                
                If (StartLineX = x1) And (StartLineY = y1) And (StartLineZ = z1) Then
                    CloseTranslateForm (lines)
                    Exit Sub
                Else
                    StartLineX_Zero
                End If
                
                If (CheckX = x1) And (CheckY = y1) And (CheckZ = z1) Then
                    CloseTranslateForm (lines)
                    Exit Sub
                Else
                    CheckX_Zero
                End If
                
                If (CheckXX = x1) And (CheckYY = y1) And (CheckZZ = z1) Then
                    CloseTranslateForm (lines)
                    Exit Sub
                Else
                    CheckXX_Zero
                End If
                
                '''''''''''''''''''''''''
                '   May not be needed   '
                '''''''''''''''''''''''''
                ''Make start and stop procedure when drawing a acr
                ''If (Spray_Valve = True) Then
                '    '10101 means "start and stop"
                '    A.writeline ("end3D(x=" & CLng(x1 * 1000) & ", y=" & CLng(y1 * 1000) & ", z=" & CLng(z1 * 1000) & "; " & words2(1) & "; " & words2(2) & "; 0.000; z=0; sp=10101; z=0)")
                '    NoChange = False
                ''End If
                ''For rotation
                'If (Spray_Valve = True) Then
                '    'Check_Rotation_Angle (Rotation_Angle)
                '    'Previous_State_Rotation = Rotation_Angle
                '    Check_Rotation_Angle ("None")
                '    Previous_State_Rotation = "None"
                'End If
                
            Case "   linksArcRestart"
                'line = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                If (Rotation_Angle <> "None") Then ' z value will be replaced with U value
                    words3() = Split(words2(0), ",")
                    
                    If (words2(2) = " 0") Then
                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 40" & ")"
                    ElseIf (words2(2) = " 1") Then
                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 41" & ")"
                    Else
                        line = "   linksArcStart(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 42" & ")"
                    End If
                    
                    XYU_Interpolate = True
                    words4() = Split(words3(2), "=")
                    'PreArcDelay = ArcDelay
                Else
                    line = "   linksArcStart(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                    XYU_Interpolate = False
                    'PreArcDelay = ArcDelay
                End If
                
                NoChange3 = True
                Data (line)
                x3 = CalculateX
                y3 = CalculateY
                z3 = CalculateZ
                
                'Check_Rotation_Angle (Rotation_Angle)
                            
                If (Spray_Valve = True) Then
                    'No previous node, starting with "Arc"
                    If (Previous_State_Rotation = "") Then
                        Previous_State_Rotation = "None"
                    End If
                    
                    If (Rotation_Angle <> "None") Then     '@$K
                        CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                    
                    ElseIf (Previous_State_Rotation <> "None") Then
                        If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "   linksArcRestart", Previous_State_Rotation) = True Then
                            A.Close
                            Unload Me
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
                    Unload Me
                    Exit Sub
                End If
                
                If NoChange3 = False Then
                    x1 = x3
                    y1 = y3
                    z1 = z3
                    Change = True
                End If
                
            Case "   linksArcEnd"
                line = "   linksArcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ")"
                Rotation_Angle = Left(angle(5), Len(angle(5)) - 1)
                
                NoChange2 = True
                Data (line)
                x3 = CalculateX
                y3 = CalculateY
                z3 = CalculateZ
                'To check the next position
                CheckX = x3
                CheckY = y3
                CheckZ = z3
                
                'Check_Rotation_Angle (Rotation_Angle)
                If (Spray_Valve = True) Then
                    'No previous node, starting with "Arc"
                    If (Previous_State_Rotation = "") Then
                        Previous_State_Rotation = "None"
                    End If
                    
                    If (Previous_State_Rotation <> "None") Then
                        If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "   linksArcEnd", Previous_State_Rotation) = True Then
                            A.Close
                            Unload Me
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
                    Unload Me
                    Exit Sub
                End If
                
                'For rotation
                If (Spray_Valve = True) And (Rotation_Angle <> "None") Then
                    Turnning_Line_Angle (Rotation_Angle)
                    Previous_State_Rotation = Rotation_Angle
                End If
            Case "lineEnd"
                line = "lineEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Data (line)
                EndLineX = CalculateX
                EndLineY = CalculateY
                EndLineZ = CalculateZ
                If (StartLineX = EndLineX) And (StartLineY = EndLineY) And (StartLineZ = EndLineZ) Then
                    CloseTranslateForm (lines)
                    Exit Sub
'                Else
'                    StartLineX_Zero
                End If
                
                If (CheckXX = EndLineX) And (CheckYY = EndLineY) And (CheckZZ = EndLineZ) Then
                    CloseTranslateForm (lines)
                    Exit Sub
'                Else
'                    CheckXX_Zero
                End If
                If (CheckX = EndLineX) And (CheckY = EndLineY) And (CheckZ = EndLineZ) Then
                    CloseTranslateForm (lines)
                    Exit Sub
'                Else
'                    CheckX_Zero
                End If
                
                If (First_Node = True) And (Second_Node = True) Then
                    'Do the calculation
                    If (First_Group = True) Then
                        If (Three_Points_Collinear(StartLineX, StartLineY, StartLineZ, CheckXX, CheckYY, CheckZZ, EndLineX, EndLineY, EndLineZ) = True) Then
                            A.writeline ("line3D(x=" & CLng(CheckXX * 1000) & ", y=" & CLng(CheckYY * 1000) & ", z=" & CLng(CheckZZ * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                        Else
                            If (LinkLine_Arc(StartLineX, StartLineY, StartLineZ, CheckXX, CheckYY, CheckZZ, EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                A.Close
                                Unload Me
                                Exit Sub
                            End If
                        End If
                        First_Group = False
                    ElseIf (Second_Group = True) Then
                        If (Three_Points_Collinear(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ) = True) Then
                            A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                        Else
                            If (LinkLine_Arc(Link_Line_Point1(0), Link_Line_Point1(1), Link_Line_Point1(2), Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                A.Close
                                Unload Me
                                Exit Sub
                            End If
                        End If
                        Second_Group = False
                    ElseIf (Third_Group = True) Then
                        If (Three_Points_Collinear(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ) = True) Then
                            A.writeline ("line3D(x=" & CLng(Link_Line_Point2(0) * 1000) & ", y=" & CLng(Link_Line_Point2(1) * 1000) & ", z=" & CLng(Link_Line_Point2(2) * 1000) & "; sp=" & First_Speed & "; " & Trim(Middle_Pt_OnOff) & ")")
                        Else
                            If (LinkLine_Arc(CheckX, CheckY, CheckZ, Link_Line_Point2(0), Link_Line_Point2(1), Link_Line_Point2(2), EndLineX, EndLineY, EndLineZ, First_Speed) < 0) Then
                                A.Close
                                Unload Me
                                Exit Sub
                            End If
                        End If
                        Third_Group = False
                    End If
                    
                    'Doesn't change the travell speed
                    'line = "end3D(" & words(1)
                    'Change the travell speed because the second line is too short
                    line = "end3d(x=" & CLng(EndLineX * 1000) & ", y=" & CLng(EndLineY * 1000) & ", z=" & CLng(EndLineZ * 1000) & "; sp=" & First_Speed & ";" & Valve_OnOff(2) & ";" & Valve_OnOff(3) & ";" & Valve_OnOff(4) & ";" & Valve_OnOff(5) & ";" & Valve_OnOff(6) & ")"
                    A.writeline (line)
                    Second_Node = False
                    First_Node = False
                    
                    'line = "end3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    XYU_Interpolate = False    '@$K
                Else
                    line = "end3D(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    XYU_Interpolate = False
                    A.writeline (line)
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
                Rotation_Angle = Left(angle(8), Len(angle(8)) - 1)
                If (Rotation_Angle <> "None") Then
                    words3() = Split(words2(0), ",")
                    If (words2(2) = " 0") Then
                        line = "arcEnd(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 40" & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    ElseIf (words2(2) = " 1") Then
                        line = "arcEnd(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 41" & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    Else
                        line = "arcEnd(" & words3(0) & "," & words3(1) & "," & " z=" & CLng(Rotation_Angle) * 10 & ";" & words2(1) & "; 42" & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    End If
                    
                    Arc_Start = True
                    'PreArcDelay = ArcDelay
                Else
                    line = "arcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                    Arc_Start = False
                    'PreArcDelay = ArcDelay
                End If
                
                'line = "arcEnd(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Data (line)
                x3 = CalculateX
                y3 = CalculateY
                z3 = CalculateZ
                If (Spray_Valve = True) Then
                    'No previous node, starting with "Arc"
                    If (Previous_State_Rotation = "") Then
                        Previous_State_Rotation = "None"
                    End If
                    If RotationAngle <> "None" Then    '@$K
                        CalculationResult = Calculation3D(x1, y1, z1, x2, y2, z2, x3, y3, z3)
                    ElseIf (Previous_State_Rotation <> "None") Then
                        If Turnning_Arc_Angle(x1, x2, x3, y1, y2, y3, z1, z2, z3, Arc_Start, "arcEnd", Previous_State_Rotation) = True Then
                            A.Close
                            Unload Me
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
                    Unload Me
                    Exit Sub
                End If
            Case "rectC1"
                line = "rectC1(" & words2(0) & ";" & words2(1) & ")"
    
                'If (First_Rect = True) Then
                '    If (Calculate_Spray_Rect = True) Then
                '        Exit Sub
                '    End If
                'End If
                
                If (words2(3) = "1") Then
                    No_Fillet_Area = 1
                Else
                    If (First_Rect = True) Then
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
                
                Data (line)
                x1 = CalculateX
                y1 = CalculateY
                z1 = CalculateZ
                
                'For rotation
                If (Spray_Valve = True) Then
                    'Check_Rotation_Angle (Rotation_Angle)
                    'Previous_State_Rotation = Rotation_Angle
                    Check_Rotation_Angle ("None")
                    Previous_State_Rotation = "None"
                End If
                XYU_Interpolate = False    '@$K
            Case "   rectC2"
                line = "   rectC2(" & words2(0) & ")"
                Data (line)
                x2 = CalculateX
                y2 = CalculateY
                z2 = CalculateZ
                If (x1 = x2) And (y1 = y2) And (z1 = z3) Then
                    CloseTranslateForm (lines)
                    Exit Sub
                End If
                XYU_Interpolate = False    '@$K
            Case "rectC3"
                line = "rectC3(" & words2(0) & ";" & words2(1) & ";" & words2(2) & ";" & words2(3) & ";" & words2(4) & ";" & words2(5) & ";" & words2(6) & ")"
                Speed = DispensingSpeed(line)
                Data (line)
                x3 = CalculateX
                y3 = CalculateY
                z3 = CalculateZ
                
                If (x2 = x3) And (y2 = y3) And (z2 = z3) Then
                    CloseTranslateForm (lines)
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
                        
                        'If Spray_CalculationRectangle_Xplus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
                        '    A.Close
                        '    Unload Me
                        '    Exit Sub
                        'End If
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
                        
                        If (Calculate_Spray_Rect = True) Then
                            Exit Sub
                        End If
                        
                    End If
                Else
                    Middle_Pt_OnOff = words2(2)
                    If CalculationRectangle(x1, y1, z1, x2, y2, z2, x3, y3, z3, Speed) < 0 Then
                    'If CalculationRectangle_XW(x1, y1, z1, x2, y2, z2, x3, y3, z3, Speed) < 0 Then
                        A.Close
                        Unload Me
                        Exit Sub
                    End If
                End If
                XYU_Interpolate = False    '@$K
            Case "repeat"
                If (First_Rect = True) Then
                    'If CalculationRectangle(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1) < 0 Then
                    If CalculationRectangle_XW(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1) < 0 Then
                        A.Close
                        Unload Me
                        Exit Sub
                    End If
                    First_Rect = False
                End If
                
                If (First_Time_PtArray = True) Then
                    'For systemMoveHeight after finishing one part array
                    A.writeline ("dot(x=0, y=0, z=0; 10, 1; 10, 1; z=0; sp=77777; 0.000; 0.000; z=0)")
                Else
                    First_Time_PtArray = True
                End If
                
                A.writeline (line)
            Case "*** Left-Needle ***"
                If (First_Rect = True) Then
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                
                'set the default value, 66666, for leftNeedle becuase of parser (testing XW)
                A.writeline ("dot(x=66666, y=66666, z=66666; 66666, 66666; 66666, 66666; z=66666; sp=66666; 66666.000; 66666.000; z=66666)")
                'Spray_Valve = True
                Spray_Valve = False       '@$K
                leftside = True
                
            Case "*** Right-Needle ***"
                If (First_Rect = True) Then
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                
                'set the default value, 55555, for rightNeedle becuase of parser (testing XW)
                A.writeline ("dot(x=55555, y=55555, z=55555; 55555, 55555; 55555, 55555; z=55555; sp=55555; 55555.000; 55555.000; z=55555)")
                'Spray_Valve = False
                Spray_Valve = True    '@$K
                rightside = True
            End Select
            
            CompareString = words(0)
            
            If aPreProcess.AtEndOfStream Then
                'NNO (check whether incomplete program or not)
                If ((CompareString = "reference") Or (CompareString = "fudicial") Or (CompareString = "*** Left-Needle ***") Or (CompareString = "*** Right-Needle ***") Or _
                    (CompareString = "lineStart") Or (CompareString = "rectC1") Or (CompareString = "   rectC2") Or _
                    (CompareString = "   linksArcStart") Or (CompareString = "   linksArcRestart") Or (CompareString = "arcStart") Or (CompareString = "       arcPoint")) Then
                    MsgBox ("Please Check the program again whether the procedure is wrong or not!")
                    A.Close
                    Unload Me
                    Exit Sub
                End If
                'Only one rectangle for Spray Valve
                If (First_Rect = True) Then
                    If (Calculate_Spray_Rect = True) Then
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next
    Loop
    
    aPreProcess.Close
    
    If proceedToNextStep = True Then
        A.Close
        Set A = fs.createtextfile(App.Path & "\preProcessedPattern.txt", True)
        
        Set aPreProcess = fsPreProcess.OpenTextFile(App.Path & "\TempoaryPattern.txt", 1, False)
    
        Dim offsetlist As New Offsets
    
        Call offsetlist.setOffsets(0, 0, 0)
    
        Dim Stack As New OffsetStack
    
        Stack.Push offsetlist

        Set offsetstk = Stack
         
        PrintParseTree ("Preprocessing......")

        Do While aPreProcess.AtEndOfStream <> True
            readstring = ""
            For lines = 1 To 100
                tmpstring = aPreProcess.ReadLine & vbNewLine
                readstring = readstring & tmpstring
                If aPreProcess.AtEndOfStream Then
                    Exit For
                End If
            Next
            If DoPreProcessParsing(readstring) = True Then
                proceedToNextStep = False
                Exit Do
            End If
        Loop
        
        'Check whether arc and rectangle are only two points or not
        If readstring = "" Then
            MsgBox ("Please Check the program again whether the procedure is wrong or not!")
            A.Close
            Unload Me
            Exit Sub
        End If
        
        aPreProcess.Close
    
        PrintParseTree ("Preprocessing ended")
    End If
    
    If proceedToNextStep = True Then
        A.Close
        
        Set A = fs.createtextfile(App.Path & "\translatedPattern.txt", True)
    
        Set aPreProcess = fsPreProcess.OpenTextFile(App.Path & "\preprocessedPattern.txt", 1, False)
        DoInitializeStateMachine
        PrintParseTree ("Translating......")
    

        Do While aPreProcess.AtEndOfStream <> True
            readstring = ""
            For lines = 1 To 100
                tmpstring = aPreProcess.ReadLine & vbNewLine
                readstring = readstring & tmpstring
                If aPreProcess.AtEndOfStream Then
                    Exit For
                End If
            Next
            If DoTranslateParse(readstring) = True Then
                proceedToNextStep = False
                Exit Do
            End If
        Loop
    
        aPreProcess.Close
        PrintParseTree ("Translation ended")
    End If
    
    If proceedToNextStep = True Then
        Visible = False
        executionForm.Show (vbModal)
        Visible = True
    End If
    
    Unload Me
            
End Sub

Private Sub Ini_Val_Rec()
    Dim i As Integer
    
    For i = 0 To 2
        First_RetC1(i) = 0
        First_RetC2(i) = 0
        First_RetC3(i) = 0
    Next
    
    Pitch = 0
    Speed_1 = 0
    No_Fillet_Area = 0
    
    Rotation_Angle = ""
    Previous_State_Rotation = ""
    
    First_Rect = False
    Spray_Valve = False
    Arc_Start = False
    
End Sub

Private Function Calculate_Spray_Rect() As Boolean
    'sp=10101 => No tilt, 11111 => flag for rectangle
    A.writeline ("dot(x=11111, y=11111, z=11111; 55555, 55555; 55555, 55555; z=55555; sp=10101; 55555.000; 55555.000; z=55555)")
                      
    If (First_RetC1(1) = First_RetC2(1)) And (First_RetC1(0) < First_RetC2(0)) Then
        If (First_RetC2(1) > First_RetC3(1)) Then
            If Not ((y1 = y2) And (x1 < x2) And (y2 > y3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((y1 = y2) And (x1 < x2) And (y2 < y3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
        
        If Spray_CalculationRectangle_Xplus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Unload Me
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(1) = First_RetC2(1)) And (First_RetC1(0) > First_RetC2(0)) Then
        If (First_RetC2(1) > First_RetC3(1)) Then
            If Not ((y1 = y2) And (x1 > x2) And (y2 > y3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((y1 = y2) And (x1 > x2) And (y2 < y3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
                            
        If Spray_CalculationRectangle_Xminus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Unload Me
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(0) = First_RetC2(0)) And (First_RetC1(1) < First_RetC2(1)) Then
        If (First_RetC2(0) > First_RetC3(0)) Then
            If Not ((x1 = x2) And (y1 < y2) And (x2 > x3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((x1 = x2) And (y1 < y2) And (x2 < x3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
                            
        If Spray_CalculationRectangle_Yplus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Unload Me
            Calculate_Spray_Rect = True
            Exit Function
        End If
    ElseIf (First_RetC1(0) = First_RetC2(0)) And (First_RetC1(1) > First_RetC2(1)) Then
        If (First_RetC2(0) > First_RetC3(0)) Then
            If Not ((x1 = x2) And (y1 > y2) And (x2 > x3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        Else
            If Not ((x1 = x2) And (y1 > y2) And (x2 < x3)) Then
                WrongProcedure
                Calculate_Spray_Rect = True
                Exit Function
            End If
        End If
        
        If Spray_CalculationRectangle_Yminus(First_RetC1(0), First_RetC1(1), First_RetC1(2), First_RetC2(0), First_RetC2(1), First_RetC2(2), First_RetC3(0), First_RetC3(1), First_RetC3(2), Speed_1, Pitch, No_Fillet_Area) < 0 Then
            A.Close
            Unload Me
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

Private Function Check_Rotation_Angle(ByVal Rot As String)
    Turnning_Angle (Rot)
End Function

Private Sub StartLineX_Zero()
    StartLineX = 0
    StartLineY = 0
    StartLineZ = 0
End Sub

Private Sub CheckX_Zero()
    CheckX = 0
    CheckY = 0
    CheckZ = 0
End Sub

Private Sub CheckXX_Zero()
    CheckXX = 0
    CheckYY = 0
    CheckZZ = 0
End Sub

Private Sub PrintParseTree(Text As String)
   'This sub just appends the Text to the end of the txtParseTree textbox.
    txtMessage.Text = txtMessage.Text & Text & vbNewLine
End Sub

Private Sub closeButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Skin1.LoadSkin (".\skin\epoxySkin.skn")
    Skin1.LoadSkin ("C:\MainProject\ProductionEditor4\skin\epoxySkin.skn") 'for login (NNO)
    Skin1.ApplySkin Me.hWnd
End Sub

Private Function CloseTranslateForm(ByVal lines As Integer)
    MsgBox "Line " & lines - 1 & " and Line " & lines & " can't be same positions!"
    A.Close             'Close the .txt file first, then unload me
    Unload Me
End Function

Private Function CloseForm(ByVal lines As Integer, ByVal CompString As String)
    MsgBox "Line " & lines - 1 & " and Line " & lines & " are wrong procedure!"
    A.Close             'Close the .txt file first, then unload me
    Unload Me
End Function

Private Function CloseForm2(ByVal lines As Integer)
    MsgBox "Line 1 is wrong procedure!"
    A.Close
    Unload Me
End Function

Private Sub WrongProcedure()
    MsgBox "The procedures for no fill area are wrong!"
    A.Close
    Unload Me
End Sub







