Attribute VB_Name = "Module5"
Function TaxCals(Age As Integer, Income As Double, Optional Savings As Double)
If Savings > 150000 Then
Savings = 150000
End If
Income = Income - Savings

Select Case Age

Case Is <= 60

    Select Case Income
    Case Is < 250000
    TaxCals = 0
    Case 250001 To 500000
    TaxCals = (Income - 250000) * 0.1
    Case 500001 To 1000000
    TaxCals = 25000 + (Income - 500000) * 0.2
    Case Is > 1000000
    TaxCals = 125000 + (Income - 1000000) * 0.3
    End Select

Case 61 To 80

    Select Case Income
    Case Is < 300000
    TaxCals = 0
    Case 300001 To 500000
    TaxCals = (Income - 300000) * 0.1
    Case 500001 To 1000000
    TaxCals = 20000 + (Income - 500000) * 0.2
    Case Is > 1000000
    TaxCals = 120000 + (Income - 1000000) * 0.3
    End Select

Case Is > 80

    Select Case Income
    Case Is < 500000
    TaxCals = 0
    Case 500001 To 1000000
    TaxCals = (Income - 500000) * 0.2
    Case Is > 1000000
    TaxCals = 100000 + (Income - 1000000) * 0.3
    End Select
    
End Select
End Function











