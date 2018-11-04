Attribute VB_Name = "Module1"
Function Test1()
Test = "Excel"
End Function

Sub MessageBox()
MsgBox "Hi"
End Sub

Sub MessageBox1()
MsgBox "Completed", Title:="Status"
End Sub

Sub Variable()
Dim X As Double
Dim Y As Double
Dim Z As Double

X = 10000
Y = 200
Z = X * Y

MsgBox Z
End Sub

Sub Variable1()
Dim EmpName As String
Dim EmpAge As Integer
Dim EmpIncome As Double

EmpName = InputBox("Please enter the Employee Name", "Name")
EmpAge = InputBox("Please enter the Employee Age", "Age")
EmpIncome = InputBox("Please enter the Employee Income", "Income")

Range("A2").Value = EmpName
Range("B2").Value = EmpAge
Range("C2").Value = EmpIncome

End Sub

Sub SelectCell()
Range("C4").Select
End Sub

Sub SelectMultipleCells()
Range("C5 , E5 , G5").Select
End Sub

Sub SelectAllCells()
Cells.Select
End Sub

Sub SelectCurrentRegion()
ActiveCell.CurrentRegion.Select
End Sub

Sub SelectRange()
Range("C5:D13").Select
End Sub

Sub SelectRange1()
Range("C5", "D13").Select
End Sub

Sub ActivecellValue()
ActiveCell.Value = "Excel"
End Sub

Sub RangeValue()
Range("C8").Value = "VBA"
End Sub

Sub CellValue()
Cells(8, 4).Value = "XXX"
End Sub

Sub ActivecellAddress()
MsgBox ActiveCell.Address
End Sub

Sub SelectionAddress()
MsgBox Selection.Address
End Sub

Sub SelectionAddressWithoutDollor()
MsgBox Selection.Address(0, 0)
End Sub

Sub SelectRow()
Rows("9").Select
End Sub

Sub SelectRows()
Rows("9:13").Select
End Sub

Sub SelectMultipleRows()
Range("5:6, 9:11, 13:14").Select
End Sub

Sub SelectColumn()
Columns("C").Select
End Sub

Sub SelectColumns()
Columns("C:D").Select
End Sub

Sub SelectMultipleColumns()
Range("B:C , E:F,  H:I").Select
End Sub

Sub SelectActivecellRow()
ActiveCell.EntireRow.Select
End Sub

Sub SelectActiveCEllColumn()
ActiveCell.EntireColumn.Select
End Sub

Sub MoveEndDown()
ActiveCell.End(xlDown).Select
End Sub

Sub MoveEndUp()
ActiveCell.End(xlUp).Select
End Sub

Sub MoveEndRight()
ActiveCell.End(xlToRight).Select
End Sub

Sub MoveEndLeft()
ActiveCell.End(xlToLeft).Select
End Sub

Sub SelectionDown()
Range(Selection, Selection.End(xlDown)).Select
End Sub

Sub SelectionUp()
Range(Selection, Selection.End(xlUp)).Select
End Sub

Sub SelectionRight()
Range(Selection, Selection.End(xlToRight)).Select
End Sub

Sub SelectionLeft()
Range(Selection, Selection.End(xlToLeft)).Select
End Sub

Sub SelectEntireData()
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
End Sub


Sub OffsetDown()
ActiveCell.Offset(2, 0).Select
End Sub

Sub OffsetUp()
ActiveCell.Offset(-3, 0).Select
End Sub

Sub OffsetRight()
ActiveCell.Offset(0, 3).Select
End Sub

Sub OffsetLeft()
ActiveCell.Offset(0, -3).Select
End Sub

Sub OffsetMethod()
ActiveCell.Offset(-5, -2).Select
End Sub

Sub SelectionEntireData()
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.Offset(2, 0)).Select
End Sub
