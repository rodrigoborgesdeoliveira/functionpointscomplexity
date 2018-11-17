' Return the External Inquiry Input's (EQI) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getEQIComplexity(ByVal ftr As Integer, ByVal det As Integer) As String
    If ftr = 0 Or ftr = 1 Then
        If det > 0 And det < 5 Then
            getEQIComplexity = "Low"
        ElseIf det > 4 And det < 16 Then
            getEQIComplexity = "Low"
        ElseIf det > 15 Then
            getEQIComplexity = "Average"
        Else
            getEQIComplexity = "Invalid DET"    
        End If
    ElseIf ftr = 2 Then
        If det > 0 And det < 5 Then
            getEQIComplexity = "Low"
        ElseIf det > 4 And det < 16 Then
            getEQIComplexity = "Average"
        ElseIf det > 15 Then
            getEQIComplexity = "High"
        Else
            getEQIComplexity = "Invalid DET"    
        End If
    ElseIf ftr > 2 Then
        If det > 0 And det < 5 Then
            getEQIComplexity = "Average"
        ElseIf det > 4 And det < 16 Then
            getEQIComplexity = "High"
        ElseIf det > 15 Then
            getEQIComplexity = "High"
        Else
            getEQIComplexity = "Invalid DET"    
        End If
    Else
        getEQIComplexity = "Invalid FTR"
    End If

End Function

' Return the External Inquiry Input's (EQI) unadjusted points as 3, 4, 6 or 
' -1 (in case there is any invalid argument)
Function getEQIUnadjustedPoints(ByVal ftr As Integer, ByVal det As Integer) As Integer
    If ftr = 0 Or ftr = 1 Then
        If det > 0 And det < 5 Then
            getEQIUnadjustedPoints = 3
        ElseIf det > 4 And det < 16 Then
            getEQIUnadjustedPoints = 3
        ElseIf det > 15 Then
            getEQIUnadjustedPoints = 4
        Else
            getEQIUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr = 2 Then
        If det > 0 And det < 5 Then
            getEQIUnadjustedPoints = 3
        ElseIf det > 4 And det < 16 Then
            getEQIUnadjustedPoints = 4
        ElseIf det > 15 Then
            getEQIUnadjustedPoints = 6
        Else
            getEQIUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr > 2 Then
        If det > 0 And det < 5 Then
            getEQIUnadjustedPoints = 4
        ElseIf det > 4 And det < 16 Then
            getEQIUnadjustedPoints = 6
        ElseIf det > 15 Then
            getEQIUnadjustedPoints = 6
        Else
            getEQIUnadjustedPoints = "Invalid DET"    
        End If
    Else
        getEQIUnadjustedPoints = "Invalid FTR"
    End If

End Function

' Return the External Inquiry Output's (EQO) complexity as Low, Average, High or 
' Invalid (in case there is any invalid argument)
Function getEQOComplexity(ByVal ftr As Integer, ByVal det As Integer) As String
    If ftr = 1 Then
        If det > 0 And det < 6 Then
            getEQOComplexity = "Low"
        ElseIf det > 5 And det < 20 Then
            getEQOComplexity = "Low"
        ElseIf det > 19 Then
            getEQOComplexity = "Average"
        Else
            getEQOComplexity = "Invalid DET"    
        End If
    ElseIf ftr = 2 Or ftr = 3 Then
        If det > 0 And det < 6 Then
            getEQOComplexity = "Low"
        ElseIf det > 5 And det < 20 Then
            getEQOComplexity = "Average"
        ElseIf det > 19 Then
            getEQOComplexity = "High"
        Else
            getEQOComplexity = "Invalid DET"    
        End If
    ElseIf ftr > 3 Then
        If det > 0 And det < 6 Then
            getEQOComplexity = "Average"
        ElseIf det > 5 And det < 20 Then
            getEQOComplexity = "High"
        ElseIf det > 19 Then
            getEQOComplexity = "High"
        Else
            getEQOComplexity = "Invalid DET"    
        End If
    Else
        getEQOComplexity = "Invalid FTR"
    End If

End Function

' Return the External Inquiry Output's (EQO) unadjusted points as 3, 4, 6 or 
' -1 (in case there is any invalid argument)
Function getEQOUnadjustedPoints(ByVal ftr As Integer, ByVal det As Integer) As Integer
    If ftr = 1 Then
        If det > 0 And det < 6 Then
            getEQOUnadjustedPoints = 3
        ElseIf det > 5 And det < 20 Then
            getEQOUnadjustedPoints = 3
        ElseIf det > 19 Then
            getEQOUnadjustedPoints = 4
        Else
            getEQOUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr = 2 Or ftr = 3 Then
        If det > 0 And det < 6 Then
            getEQOUnadjustedPoints = 3
        ElseIf det > 5 And det < 20 Then
            getEQOUnadjustedPoints = 4
        ElseIf det > 19 Then
            getEQOUnadjustedPoints = 6
        Else
            getEQOUnadjustedPoints = "Invalid DET"    
        End If
    ElseIf ftr > 3 Then
        If det > 0 And det < 6 Then
            getEQOUnadjustedPoints = 4
        ElseIf det > 5 And det < 20 Then
            getEQOUnadjustedPoints = 6
        ElseIf det > 19 Then
            getEQOUnadjustedPoints = 6
        Else
            getEQOUnadjustedPoints = "Invalid DET"    
        End If
    Else
        getEQOUnadjustedPoints = "Invalid FTR"
    End If

End Function