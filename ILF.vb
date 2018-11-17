' Return the Internal Logic File's (ILF) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getILFComplexity(ByVal ret As Integer, ByVal det As Integer) As String
    If ret = 1 Then
        If det > 0 And det < 20 Then
            getILFComplexity = "Low"
        ElseIf det > 19 And det < 51 Then
            getILFComplexity = "Low"
        ElseIf det > 50 Then
            getILFComplexity = "Average"
        Else
            getILFComplexity = "Invalid DET"    
        End If
    ElseIf ret > 1 And ret < 6 Then
        If det > 0 And det < 20 Then
            getILFComplexity = "Low"
        ElseIf det > 19 And det < 51 Then
            getILFComplexity = "Average"
        ElseIf det > 50 Then
            getILFComplexity = "High"
        Else
            getILFComplexity = "Invalid DET"    
        End If
    ElseIf ret > 5 Then
        If det > 0 And det < 20 Then
            getILFComplexity = "Average"
        ElseIf det > 19 And det < 51 Then
            getILFComplexity = "High"
        ElseIf det > 50 Then
            getILFComplexity = "High"
        Else
            getILFComplexity = "Invalid DET"    
        End If
    Else
        getILFComplexity = "Invalid RET"
    End If

End Function

' Return the Internal Logic File's (ILF) unadjusted points as 7, 10, 15 or 
' -1 (in case there is any invalid argument)
Function getILFUnadjustedPoints(ByVal ret As Integer, ByVal det As Integer) As Integer
    If ret = 1 Then
        If det > 0 And det < 20 Then
            getILFUnadjustedPoints = 7
        ElseIf det > 19 And det < 51 Then
            getILFUnadjustedPoints = 7
        ElseIf det > 50 Then
            getILFUnadjustedPoints = 10
        Else
            getILFUnadjustedPoints = -1    
        End If
    ElseIf ret > 1 And ret < 6 Then
        If det > 0 And det < 20 Then
            getILFUnadjustedPoints = 7
        ElseIf det > 19 And det < 51 Then
            getILFUnadjustedPoints = 10
        ElseIf det > 50 Then
            getILFUnadjustedPoints = 15
        Else
            getILFUnadjustedPoints = -1    
        End If
    ElseIf ret > 5 Then
        If det > 0 And det < 20 Then
            getILFUnadjustedPoints = 10
        ElseIf det > 19 And det < 51 Then
            getILFUnadjustedPoints = 15
        ElseIf det > 50 Then
            getILFUnadjustedPoints = 15
        Else
            getILFUnadjustedPoints = -1    
        End If
    Else
        getILFUnadjustedPoints = -1
    End If

End Function