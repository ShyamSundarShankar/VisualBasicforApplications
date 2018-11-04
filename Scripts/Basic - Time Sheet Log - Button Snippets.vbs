Attribute VB_Name = "Module1"
Private Sub Workbook_Open()

Application.DisplayFormulaBar = False
Application.DisplayFullScreen = True
Application.DisplayScrollBars = False
Application.DisplayStatusBar = False

Worksheets("Data").Select
MsgBox "Hi " & Cells(1, 3).Value & "! Have a Great Day Ahead..", Title:="Log"

End Sub

Sub StartTime()

MsgBox "Are you sure that your task is Started?", vbYesNo, Title:="Log"
If vbYes Then

Application.ScreenUpdating = False
ActiveCell.Value = "=Now()"
ActiveCell.Copy
ActiveCell.PasteSpecial (xlPasteValues)
ActiveCell.Offset(0, 1).Select
Application.CutCopyMode = False


Else
MsgBox "Please Relax..", Title:="Log"

End If
End Sub

Sub EndTime()

MsgBox "Are you sure that your task is Ended?", vbYesNo, Title:="Log"
If vbYes Then

Application.ScreenUpdating = False
ActiveCell.Value = "=Now()"
ActiveCell.Copy
ActiveCell.PasteSpecial (xlPasteValues)
ActiveCell.Offset(1, -6).Select
Range(Selection, Selection.End(xlUp)).Select
Range(Selection, Selection.End(xlToLeft)).Select
Selection.FillDown
ActiveCell.End(xlDown).Select
ActiveCell.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
Application.CutCopyMode = False

Else
MsgBox "Please Continue..", Title:="Log"

End If
End Sub
Function UserName()

UserName = Environ$("UserName")
End Function

Sub Mail()

ActiveWorkbook.Save
Application.DisplayAlerts = True

Dim OutApp As Object
Dim OutMail As Object
Dim EmailAddr As String
Dim Subj As String
Dim BodyText As String
Dim HOURLY As String

EmailAddr = Cells(2, 1).Value
EmailAddr1 = Cells(3, 1).Value

Subj = Cells(1, 1).Value
HOURLY = ""
BodyText = Cells(1, 2).Value

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

With OutMail

.to = EmailAddr
.CC = EmailAddr1
.BCC = ""
.Subject = Subj
.Body = BodyText & vbNewLine & vbNewLine & "Please see attached the Production Log As On Dated " & Format(Date, "mm/dd/yyyy") & "." & vbNewLine & vbNewLine & "Please do writeback for Clarifications or Questions if any." & vbNewLine & vbNewLine & "Best Regards - " & Cells(5, 1).Value


.Attachments.Add ActiveWorkbook.FullName

.Display 'or use .send
End With
End Sub

Sub Help()

MsgBox "Do you have any Glitches/Challenges in filling the Log?", vbYesNo, Title:="Log"

Dim IssueType As String

If vbYes Then
IssueType = InputBox("Please specify the Issue:", "Log")
Range("B2").Value = IssueType

Call HelpMail

Else

MsgBox "Good!", Title:="Log"
End If
End Sub

Sub HelpMail()

ActiveWorkbook.Save
Application.DisplayAlerts = True

Dim OutApp As Object
Dim OutMail As Object
Dim EmailAddr As String
Dim Subj As String
Dim BodyText As String
Dim HOURLY As String

EmailAddr = "ShyamSundarSOrange.0005@Gmail.Com"
EmailAddr1 = "KishoreOrange0001@Gmail.Com"


Subj = "Issue: Production Log"
HOURLY = ""
BodyText = Cells(1, 2).Value

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(0)

With OutMail

.to = EmailAddr
.CC = EmailAddr1
.BCC = ""
.Subject = Subj
.Body = BodyText & vbNewLine & vbNewLine & "I'm unable to fill the Log as there is a issue in terms of " & Cells(2, 2).Value & " ." & vbNewLine & vbNewLine & "Best Regards - " & Cells(5, 1).Value

.Display 'or use .send

End With
End Sub

Sub Instructions()

Dim InsText As String

InsText = Cells(3, 2).Value
MsgBox InsText, Title:="-----------------------------------Log Instructions-----------------------------------"

End Sub

Sub OnExit()

Application.DisplayAlerts = False
Application.DisplayFormulaBar = True
Application.DisplayFullScreen = False
Application.DisplayScrollBars = True
Application.DisplayStatusBar = True

Sheets("Data").Select
Range("a5").Select

ActiveWorkbook.Save
ActiveWorkbook.Close

End Sub

Sub DataSheet()

Sheets("Data").Select

End Sub

Sub RecipientSheet()

Sheets("Mail Recipients").Select

End Sub

