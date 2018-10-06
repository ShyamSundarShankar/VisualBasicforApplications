Attribute VB_Name = "Module2"
Sub SelectSheet()
Worksheets("March").Select
End Sub

Sub SelectSheetBasedOnInputBox()
Dim SheetName As String
SheetName = InputBox("Please enter the Sheet Name")
Worksheets(SheetName).Select
End Sub

Sub CountWorksheets()
ActiveCell.Value = Worksheets.Count
End Sub

Sub CreateWorksheet()
Worksheets.Add
End Sub

Sub CreateSheetAfterActiveSheet()
Worksheets.Add After:=ActiveSheet
End Sub

Sub CreateSheetBeforeAfter()
Worksheets.Add Before:=Worksheets(3)
End Sub

Sub CreateSheetInEnd()
Worksheets.Add After:=Worksheets(Worksheets.Count)
End Sub

Sub CreateMoreWorksheets()
Worksheets.Add Count:=3
End Sub

Sub CreateMoreWorksheets1()
For i = 1 To 3
Worksheets.Add
Next
End Sub

Sub ChangeWorksheetName()
ActiveSheet.Name = "Excel"
End Sub

Sub ChangeParticularSheetName()
Worksheets("Excel").Name = "VBA"
End Sub

Sub SelectAllSheets()
For i = 1 To Worksheets.Count
Worksheets(i).Select
Range("B1").Value = "Excel"
Next
End Sub

Sub CreateMoreSheetWithName()
Dim SheetName As String
Range("E2").Select
Do While ActiveCell.Value <> ""
SheetName = ActiveCell.Value
Worksheets.Add After:=Worksheets(Worksheets.Count)
ActiveSheet.Name = SheetName
Worksheets("Main").Select
ActiveCell.Offset(1, 0).Select
Loop
End Sub

Sub GetAllSheetName()
Dim SheetName As String
Application.ScreenUpdating = False
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
SheetName = ActiveSheet.Name
Worksheets("Main").Select
ActiveCell.Value = SheetName
ActiveCell.Offset(1, 0).Select
End If
Next
End Sub

Sub GetAllSheetNameVisibleSheets()
Dim SheetName As String
Application.ScreenUpdating = False
For i = 1 To Worksheets.Count
If Worksheets(i).Visible = True Then
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
SheetName = ActiveSheet.Name
Worksheets("Main").Select
ActiveCell.Value = SheetName
ActiveCell.Offset(1, 0).Select
End If
End If
Next
End Sub


Sub DeleteSheet()
Application.DisplayAlerts = False
ActiveSheet.Delete
End Sub

Sub DeleteParticularSheet()
Worksheets("July").Delete
End Sub

Sub DeleteAllSheetExceptActivesheet()
Dim TempSheet As Worksheet
Application.DisplayAlerts = False
For Each TempSheet In Worksheets
If TempSheet.Name <> "Main" Then
TempSheet.Delete
End If
Next
End Sub

Sub HideSheet()
ActiveSheet.Visible = False
End Sub

Sub HideParticularSheet()
Worksheets("July").Visible = False
End Sub

Sub UnHideParticularSheet()
Worksheets("July").Visible = True
End Sub

Sub UnhideAllSheets()
For i = 1 To Worksheets.Count
Worksheets(i).Visible = True
Next
End Sub



Sub HideAllSheetExceptActivesheet1()
For i = 1 To Worksheets.Count
Worksheets(i).Select
If ActiveSheet.Name <> "Main" Then
ActiveSheet.Visible = False
End If
Next
End Sub

