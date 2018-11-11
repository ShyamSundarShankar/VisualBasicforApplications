Dim wordapp
Dim worddoc
Dim WB, owb
Dim Test

Set WB = ThisWorkbook.Sheets("Index")
Set owb = ThisWorkbook.Sheets("Log")
Set wwb = ThisWorkbook.Sheets("Scrubber")
owblastrow = owb.Cells(Rows.Count, 1).End(xlUp).Row

Sheets("Log").Select
Range("E1048576").End(xlUp).Offset(1, -4).Select
For Actrow = ActiveCell.Row To owblastrow
wwb.Cells.Clear
Set wordapp = CreateObject("Word.Application")
Filename = UCase(owb.Cells(Actrow, 1))
If Err.Number = 5792 Then
GoTo Enddd
End If
Set worddoc = wordapp.Documents.Open(Filename, ReadOnly:=False)
ThisWorkbook.Activate
wwb.Activate
worddoc.Content.Copy
worddoc.Content.Copy
wwb.Activate
wwb.Range("A1").Select
Application.Wait (Now() + TimeValue("00:00:10"))
wwb.Paste
wwb.Paste

On Error Resume Next
Cells.Find(What:="Provider:", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select
If Err.Number = 0 Then
On Error GoTo 0
ProvName = ActiveCell.Value
If ProvName = "Provider: " Or ProvName = "Provider:" Then
ProvName = ActiveCell.Offset(0, 1).Value
Else
End If
Cells.Find(What:="Date of Service:", After:=ActiveCell, LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False).Select
TripNo = ActiveCell.Value
If TripNo = "Date of Service:" Then
Cells.Find(What:="Our Trip", After:=Range("A1"), LookAt:=xlPart).Select
Cells.Find(What:="Our Trip #:", After:=ActiveCell, LookAt:=xlPart).Select
TripNo = ActiveCell.Offset(0, 1).Value
Else
End If
Else
ProvName = ""
TripNo = ""
End If

Enddd:
wordapp.Quit Savechange = False
For Each wShape In wwb.Shapes
    wShape.Delete
Next wShape
wwb.Cells.Delete
wwb.Range("A1").Select
owb.Cells(Actrow, 5).Value = ProvName
owb.Cells(Actrow, 6).Value = TripNo
ActiveWorkbook.Save

Next Actrow