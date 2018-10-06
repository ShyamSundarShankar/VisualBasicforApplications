Attribute VB_Name = "Module1"
Function test()
test = "Excel"
End Function

Function Discount(Amount As Double)
Select Case Amount
Case Is < 10000
Discount = 0
Case 10000 To 20000
Discount = 0.1
Case 20001 To 30000
Discount = 0.2
Case 30001 To 40000
Discount = 0.3
Case 40001 To 50000
Discount = 0.4
Case Is > 50000
Discount = 0.5
End Select
End Function

Function DiscountValue(Amount As Double)
Select Case Amount
Case Is < 10000
DiscountValue = 0
Case 10000 To 20000
DiscountValue = Amount * 0.1
Case 20001 To 30000
DiscountValue = Amount * 0.2
Case 30001 To 40000
DiscountValue = Amount * 0.3
Case 40001 To 50000
DiscountValue = Amount * 0.4
Case Is > 50000
DiscountValue = Amount * 0.5
End Select
End Function

Function NetValue(Amount As Double)
Select Case Amount
Case Is < 10000
NetValue = Amount - (Amount * 0)
Case 10000 To 20000
NetValue = Amount - (Amount * 0.1)
Case 20001 To 30000
NetValue = Amount - (Amount * 0.2)
Case 30001 To 40000
NetValue = Amount - (Amount * 0.3)
Case 40001 To 50000
NetValue = Amount - (Amount * 0.4)
Case Is > 50000
NetValue = Amount - (Amount * 0.5)
End Select
End Function

Function SelectionProcess(Qulification As String)

Select Case UCase(Qulification)

Case "CA", "CWA", "ACS"
SelectionProcess = "Audit"
Case "BCOM", "MCOM"
SelectionProcess = "Finance"
Case "BA", "MA"
SelectionProcess = "Admin"
Case "BBA", "MBA"
SelectionProcess = "Marketing"
Case "BL", "ML"
SelectionProcess = "Legal"
Case "&"
SelectionProcess = "IT"
Case Else
SelectionProcess = "Others"
End Select
End Function




























































