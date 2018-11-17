' Return the External Input's (EI) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getEIComplexity(ByVal ftr As Integer, ByVal det As Integer) As String
    If ftr = 0 Or ftr = 1 Then
        If det > 0 And det < 5 Then
            getEIComplexity = "Low"
        ElseIf det > 4 And det < 16 Then
            getEIComplexity = "Low"
        ElseIf det > 15 Then
            getEIComplexity = "Average"
        Else
            getEIComplexity = "Invalid DET"    
        End If
    ElseIf ftr = 2 Then
        If det > 0 And det < 5 Then
            getEIComplexity = "Low"
        ElseIf det > 4 And det < 16 Then
            getEIComplexity = "Average"
        ElseIf det > 15 Then
            getEIComplexity = "High"
        Else
            getEIComplexity = "Invalid DET"    
        End If
    ElseIf ftr > 2 Then
        If det > 0 And det < 5 Then
            getEIComplexity = "Average"
        ElseIf det > 4 And det < 16 Then
            getEIComplexity = "High"
        ElseIf det > 15 Then
            getEIComplexity = "High"
        Else
            getEIComplexity = "Invalid DET"    
        End If
    Else
        getEIComplexity = "Invalid FTR"
    End If

End Function

' Return the External Input's (EI) unadjusted points as 3, 4, 6 or 
' -1 (in case there is any invalid argument)
Function getEIUnadjustedPoints(ByVal ftr As Integer, ByVal det As Integer) As Integer
    If ftr = 0 Or ftr = 1 Then
        If det > 0 And det < 5 Then
            getEIUnadjustedPoints = 3
        ElseIf det > 4 And det < 16 Then
            getEIUnadjustedPoints = 3
        ElseIf det > 15 Then
            getEIUnadjustedPoints = 4
        Else
            getEIUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr = 2 Then
        If det > 0 And det < 5 Then
            getEIUnadjustedPoints = 3
        ElseIf det > 4 And det < 16 Then
            getEIUnadjustedPoints = 4
        ElseIf det > 15 Then
            getEIUnadjustedPoints = 6
        Else
            getEIUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr > 2 Then
        If det > 0 And det < 5 Then
            getEIUnadjustedPoints = 4
        ElseIf det > 4 And det < 16 Then
            getEIUnadjustedPoints = 6
        ElseIf det > 15 Then
            getEIUnadjustedPoints = 6
        Else
            getEIUnadjustedPoints = "Invalid DET"    
        End If
    Else
        getEIUnadjustedPoints = "Invalid FTR"
    End If

End Function