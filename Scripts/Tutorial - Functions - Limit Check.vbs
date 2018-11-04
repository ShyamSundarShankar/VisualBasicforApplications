Attribute VB_Name = "Module6"
Function LimitCheck(Approver As String, TBType As String, Limit As Double)

Select Case LCase(Approver)

Case "raja"

    Select Case UCase(TBType)
    
    Case "BS"
    
    If Limit > 1000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "PL"
    
    If Limit > 100000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "OFFBS"
    
    If Limit > 10000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    End Select

Case "sundar"

    Select Case UCase(TBType)
    
    Case "BS"
    
    If Limit > 500000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "PL"
    
    If Limit > 50000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "OFFBS"
    
    If Limit > 5000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    End Select
    
Case "mohan"

    Select Case UCase(TBType)
    
    Case "BS"
    
    If Limit > 2000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "PL"
    
    If Limit > 200000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "OFFBS"
    
    If Limit > 20000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    End Select
Case "kannan"

    Select Case UCase(TBType)
    
    Case "BS"
    
    If Limit > 1500000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "PL"
    
    If Limit > 150000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "OFFBS"
    
    If Limit > 15000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    End Select

Case "saran"

    Select Case UCase(TBType)
    
    Case "BS"
    
    If Limit > 2500000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "PL"
    
    If Limit > 250000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    Case "OFFBS"
    
    If Limit > 25000000 Then
    LimitCheck = "Rejected"
    Else
    LimitCheck = "Accepted"
    End If
    
    End Select


End Select

End Function

