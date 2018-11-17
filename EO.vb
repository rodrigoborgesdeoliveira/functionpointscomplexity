' Return the External Output's (EO) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getEOComplexity(ByVal ftr As Integer, ByVal det As Integer) As String
    If ftr = 1 Then
        If det > 0 And det < 6 Then
            getEOComplexity = "Low"
        ElseIf det > 5 And det < 20 Then
            getEOComplexity = "Low"
        ElseIf det > 19 Then
            getEOComplexity = "Average"
        Else
            getEOComplexity = "Invalid DET"    
        End If
    ElseIf ftr = 2 Or ftr = 3 Then
        If det > 0 And det < 6 Then
            getEOComplexity = "Low"
        ElseIf det > 5 And det < 20 Then
            getEOComplexity = "Average"
        ElseIf det > 19 Then
            getEOComplexity = "High"
        Else
            getEOComplexity = "Invalid DET"    
        End If
    ElseIf ftr > 3 Then
        If det > 0 And det < 6 Then
            getEOComplexity = "Average"
        ElseIf det > 5 And det < 20 Then
            getEOComplexity = "High"
        ElseIf det > 19 Then
            getEOComplexity = "High"
        Else
            getEOComplexity = "Invalid DET"    
        End If
    Else
        getEOComplexity = "Invalid FTR"
    End If

End Function

' Return the External Input's (EI) unadjusted points as 4, 5, 7 or 
' -1 (in case there is any invalid argument)
Function getEOUnadjustedPoints(ByVal ftr As Integer, ByVal det As Integer) As Integer
    If ftr = 1 Then
        If det > 0 And det < 6 Then
            getEOUnadjustedPoints = 4
        ElseIf det > 5 And det < 20 Then
            getEOUnadjustedPoints = 4
        ElseIf det > 19 Then
            getEOUnadjustedPoints = 5
        Else
            getEOUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr = 2 Or ftr = 3 Then
        If det > 0 And det < 6 Then
            getEOUnadjustedPoints = 4
        ElseIf det > 5 And det < 20 Then
            getEOUnadjustedPoints = 5
        ElseIf det > 19 Then
            getEOUnadjustedPoints = 7
        Else
            getEOUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr > 3 Then
        If det > 0 And det < 6 Then
            getEOUnadjustedPoints = 5
        ElseIf det > 5 And det < 20 Then
            getEOUnadjustedPoints = 7
        ElseIf det > 19 Then
            getEOUnadjustedPoints = 7
        Else
            getEOUnadjustedPoints = "Invalid DET"    
        End If
    Else
        getEOUnadjustedPoints = "Invalid FTR"
    End If

End Function