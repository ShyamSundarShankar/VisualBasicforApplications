Attribute VB_Name = "Module1"
Sub SubmitData()

Application.ScreenUpdating = False
Sheets("FrontEnd").Select

If Range("B28").Value = 0 Then
MsgBox "Please enter all fields.", vbCritical, "Technology Helpdesk Manager"
End If

Range("B27").Select
Range(Selection, Selection.End(xlToRight)).Copy
Sheets("BackEnd").Select
Range("C1048576").End(xlUp).Offset(1, 0).PasteSpecial xlValues
TID = ActiveCell.Offset(0, -1).Value
Range("A1").Select

Sheets("FrontEnd").Select
Range("C5, C7, C9, C11, C13, C15, C17, C19, C21, C23").ClearContents
Range("C5").Select

MsgBox TID & " Ticket has been successfully created.", vbInformation, "Technology Helpdesk Manager"

End Sub
