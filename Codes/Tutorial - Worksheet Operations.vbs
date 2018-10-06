Attribute VB_Name = "Module1"
Sub Consol1()

For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
Next

End Sub

Sub Consol2()
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Range("A2").Select
If ActiveCell.Value <> "" Then
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
End If
Next

End Sub

Sub Consol3()

For i = 1 To Worksheets.Count
If Worksheets(i).Visible = True Then
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
End If
Next
End Sub

Sub Consol4()
Dim SheetName As String
Application.ScreenUpdating = False
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
SheetName = ActiveSheet.Name
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = SheetName
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
End If
Next

End Sub

Sub Consol5()

For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Range("A100000").Select
ActiveCell.End(xlUp).Select
Range(ActiveCell.Address, "A2").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
Next

End Sub

Sub Test1()
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Range("A1").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select
Range("G1").Value = "Excel"
End If
Next
End Sub

Sub Consol6()
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
Columns("C").Cut
Columns("B").Insert
Columns("E").Cut
Columns("C").Insert
Range("A2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.Offset(0, 2)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
Next
End Sub

Sub Consol7()
Application.ScreenUpdating = False
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then

Cells.Find(What:="Month").Select
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert
Cells.Find(What:="Region").Select
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert
Cells.Find(What:="Amount").Select
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert
Range("E2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.Offset(0, -2)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
Next
End Sub

Sub Consol8()
Application.ScreenUpdating = False
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
On Error Resume Next
Cells.Find(What:="Month").Select
If ActiveCell.Value <> "Month" Then
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Month"
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "Blank"
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown
Else
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert

End If
Cells.Find(What:="Region").Select
If ActiveCell.Value <> "Region" Then
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Region"
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "Blank"
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Else
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert
End If
Cells.Find(What:="Amount").Select
If ActiveCell.Value <> "Amount" Then
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.Value = "Amount"
ActiveCell.Offset(1, 0).Select
ActiveCell.Value = "Blank"
ActiveCell.Offset(0, -1).Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(0, 1).Select
Range(Selection, Selection.End(xlUp)).Select
Selection.FillDown

Else
ActiveCell.EntireColumn.Cut
Range("Z1").Select
ActiveCell.End(xlToLeft).Select
ActiveCell.Offset(0, 1).Select
ActiveCell.EntireColumn.Insert
End If
Range("E2").Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.Offset(0, -2)).Select
Selection.Copy
Worksheets("Main").Select
Range("A100000").Select
ActiveCell.End(xlUp).Select
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
End If
Next
End Sub












































