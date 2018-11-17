' Return the External Interface File's (EIF) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getEIFComplexity(ByVal ret As Integer, ByVal det As Integer) As String
    If ret = 1 Then
        If det > 0 And det < 20 Then
            getEIFComplexity = "Low"
        ElseIf det > 19 And det < 51 Then
            getEIFComplexity = "Low"
        ElseIf det > 50 Then
            getEIFComplexity = "Average"
        Else
            getEIFComplexity = "Invalid DET"    
        End If
    ElseIf ret > 1 And ret < 6 Then
        If det > 0 And det < 20 Then
            getEIFComplexity = "Low"
        ElseIf det > 19 And det < 51 Then
            getEIFComplexity = "Average"
        ElseIf det > 50 Then
            getEIFComplexity = "High"
        Else
            getEIFComplexity = "Invalid DET"    
        End If
    ElseIf ret > 5 Then
        If det > 0 And det < 20 Then
            getEIFComplexity = "Average"
        ElseIf det > 19 And det < 51 Then
            getEIFComplexity = "High"
        ElseIf det > 50 Then
            getEIFComplexity = "High"
        Else
            getEIFComplexity = "Invalid DET"    
        End If
    Else
        getEIFComplexity = "Invalid RET"
    End If

End Function

' Return the External Interface File's (EIF) unadjusted points as 5, 7, 10 or 
' -1 (in case there is any invalid argument)
Function getEIFUnadjustedPoints(ByVal ret As Integer, ByVal det As Integer) As Integer
    If ret = 1 Then
        If det > 0 And det < 20 Then
            getEIFUnadjustedPoints = 5
        ElseIf det > 19 And det < 51 Then
            getEIFUnadjustedPoints = 5
        ElseIf det > 50 Then
            getEIFUnadjustedPoints = 7
        Else
            getEIFUnadjustedPoints = -1    
        End If
    ElseIf ret > 1 And ret < 6 Then
        If det > 0 And det < 20 Then
            getEIFUnadjustedPoints = 5
        ElseIf det > 19 And det < 51 Then
            getEIFUnadjustedPoints = 7
        ElseIf det > 50 Then
            getEIFUnadjustedPoints = 10
        Else
            getEIFUnadjustedPoints = -1    
        End If
    ElseIf ret > 5 Then
        If det > 0 And det < 20 Then
            getEIFUnadjustedPoints = 7
        ElseIf det > 19 And det < 51 Then
            getEIFUnadjustedPoints = 10
        ElseIf det > 50 Then
            getEIFUnadjustedPoints = 10
        Else
            getEIFUnadjustedPoints = -1    
        End If
    Else
        getEIFUnadjustedPoints = -1
    End If

End Function