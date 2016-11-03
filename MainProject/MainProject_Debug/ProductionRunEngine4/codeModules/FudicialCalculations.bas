Attribute VB_Name = "FudicialCalculations"
Public Sub detXYfromFudicial(ByRef xFidOut As Long, ByRef yFidOut As Long, ByVal X As Long, ByVal Y As Long, ByVal xOrg As Long, ByVal yOrg As Long, ByVal dx As Long, ByVal dy As Long, ByVal dev As Double)


    If doneFudicial = True Then

        Dim x1, y1, hyp As Double
        Dim theta As Double
        Dim xdbl, ydbl As Double

        x1 = X - xOrg
        y1 = Y - yOrg
    
        If (X = xOrg) And (Y = yOrg) Then
            theta = 0
            hyp = 0
        Else
            If (X <> xOrg) Then
                theta = Atn(y1 / x1)
                If (x1 < 0) Then
                    theta = 3.14159265358979 + theta
                End If
            Else
                If (y1 > 0) Then
                    theta = 90 * 3.14159265358979 / 180
                Else
                    theta = -90 * 3.14159265358979 / 180
                End If
            End If
            hyp = Sqr(x1 * x1 + y1 * y1)
        End If
    
    'xdbl = (hyp * Cos((dev * 3.14159265358979 / 180) + theta)) + xOrg + dx
    'ydbl = (hyp * Sin((dev * 3.14159265358979 / 180) + theta)) + yOrg + dy
    
        xFidOut = CLng(hyp * Cos((dev * 3.14159265358979 / 180) + theta)) + xOrg + dx
        yFidOut = CLng(hyp * Sin((dev * 3.14159265358979 / 180) + theta)) + yOrg + dy
    
    Else
        xFidOut = X
        yFidOut = Y
    End If


End Sub
