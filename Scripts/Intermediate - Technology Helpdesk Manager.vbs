Attribute VB_Name = "Module1"
Public MSubject, MID, xRG
Sub SubmitData()

Application.ScreenUpdating = False
Sheets("Interface").Select
Sheets("Data").Visible = True

'Data Validation
If Range("P12").Value > 0 Then
MsgBox "Hey " & Application.WorksheetFunction.Proper(Environ("Username")) & ", Please enter all fields to Log the Ticket.", vbCritical, "IT Helpdesk Manager"
Exit Sub
End If

'Copying Data to Data
Range("C12:O12").Copy
Worksheets("Data").Range("B1048576").End(xlUp).Offset(1, 0).PasteSpecial xlValues
TID = Range("D6").Value

'Storing Values
Sheets("Interface").Select: MID = Range("C11").Value: Cate = Range("J6").Value
MSubject = "New " & Cate & " Ticket #" & Range("D6").Value & " has been Raised by " & Range("G8").Value

'Initiating Mail
Call LogNotify

'Clearing Contents
Range("D8, D10, G6, G8, G10, J6, J8, J10, M10").ClearContents
Range("P6").MergeArea.ClearContents: Range("D8").Select

MSubject = "": MID = "": xRG = ""
Application.CutCopyMode = False: Sheets("Data").Visible = False
MsgBox "Ticket # " & TID & " has been Successfully Created & Mailed.", vbInformation, "IT Helpdesk Manager"

End Sub
Sub LogNotify()

Dim ol As Object
Dim olEmail As Object
Dim olInsp As Object
Dim wd As Object
Dim rCol As Collection, r As Range, i As Integer

Set ol = GetObject(Class:="Outlook.Application")
Set olEmail = ol.CreateItem(0)
Set rCol = New Collection

With rCol
    .Add Sheets("Interface").Range("C6:M10")
End With

With olEmail
    .To = MID
    .Subject = MSubject
    .HTMLBody = "<html><body style=""font-family:calibri"">" & _
                "<p>Hello!<br><br>Thank you for Contacting us! We have logged in your Issue/Request and below is the details of the same. Our Engineer would get this addressed as soon as possible." & _
                "</p></body></html>"
    Set olInsp = .GetInspector
    If olInsp.EditorType = 4 Then
        Set wd = olInsp.WordEditor
        For i = 1 To rCol.Count
            Set r = rCol.Item(i): r.Copy
            wd.Range.InsertParagraphAfter
            wd.Paragraphs(wd.Paragraphs.Count).Range.PasteAndFormat 16
        Next
    End If
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "We will keep you posted once the Ticket is Closed. Please reply back or reach us for Concerns/Clarifications if any."
    wd.Range.InsertParagraphAfter
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "Regards - IT Helpdesk"
    wd.Paragraphs.Last.Range.Sentences.Last.Font.Bold = True
    wd.Range.InsertParagraphAfter
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "Call: 555/044-4207 2489 | Mail: itsupport@comnet.org.in"
    wd.Paragraphs.Last.Range.Sentences.Last.Font.Bold = True
    .Display
    .Send
End With

End Sub
Sub Search()

Application.ScreenUpdating = False
Sheets("Interface").Select
Sheets("Data").Visible = True

'Searh Validation
TID = Range("D15").Value
Sheets("Data").Select
On Error Resume Next
Range("B:B").Find(what:=TID, LookAt:=xlWhole).Select
If Err.Number = 91 Then
MsgBox "Hey " & Application.WorksheetFunction.Proper(Environ("Username")) & ", It's an Invalid Ticket #.", vbCritical, "IT Helpdesk Manager"
Sheets("Interface").Select
Range("D15").Select
On Error GoTo 0
Exit Sub
End If
On Error GoTo 0

'Search Result
Sheets("Interface").Select
Range("D17").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],2,0),"""")"
Range("D19").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],3,0),"""")"
Range("G15").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],4,0),"""")"
Range("G17").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],5,0),"""")"
Range("G19").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],6,0),"""")"
Range("J15").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],7,0),"""")"
Range("J17").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],8,0),"""")"
Range("J19").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],9,0),"""")"
Range("M15").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],10,0),"""")"
Range("M19").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],13,0),"""")"
Range("P15").Value = "=IFERROR(VLOOKUP(D15,Data[[#All],[Ticket '#]:[TAT Status]],12,0),"""")"

'Pastespecial
Worksheets("Interface").Range("D15:M19").Copy
Worksheets("Interface").Range("D15:M19").PasteSpecial xlValues
Worksheets("Interface").Range("P15").Copy
Worksheets("Interface").Range("P15").PasteSpecial xlValues
Range("M17").Value = "=IF(COUNTIFS(Data[[#All],[Category]],""Issue"",Data[[#All],[Workstation]],G15)=0,"""",COUNTIFS(Data[[#All],[Category]],""Issue"",Data[[#All],[Workstation]],G15))"

Range("D15").Select
Application.CutCopyMode = False
Sheets("Data").Visible = False
MsgBox "Ticket # " & TID & " has been found.", vbInformation, "IT Helpdesk Manager"

End Sub
Sub UpdateData()

Application.ScreenUpdating = False
Sheets("Interface").Select
Sheets("Data").Visible = True

'Data Validation
If Range("P21").Value > 0 Then
MsgBox "Hey " & Application.WorksheetFunction.Proper(Environ("Username")) & ", Please enter all fields to Log the Ticket.", vbCritical, "IT Helpdesk Manager"
Exit Sub
End If

'Copying Data to Data
TID = Range("D15").Value
Range("C21:O21").Copy
Sheets("Data").Select
Range("B:B").Find(what:=TID, LookAt:=xlWhole).Select
ActiveCell.PasteSpecial xlValues

'Storing Values
Sheets("Interface").Select
MID = Range("C20").Value: Cate = Range("J15").Value
MSubject = "Your " & Cate & " Ticket #" & Range("D15").Value & " has been updated as " & Range("M19").Value

'Initiating Mail
Call UpdateNotify

'Clearing Contents
Range("D15, D17, D19, G15, G17, G19, J15, J17, J19, M15, M19").ClearContents
Range("P15").MergeArea.ClearContents: Range("D15").Select

MSubject = "": MID = "": xRG = ""
Application.CutCopyMode = False: Sheets("Data").Visible = False
MsgBox "Ticket # " & TID & " has been Successfully Updated & Mailed.", vbInformation, "IT Helpdesk Manager"

End Sub
Sub UpdateNotify()

Dim ol As Object
Dim olEmail As Object
Dim olInsp As Object
Dim wd As Object
Dim rCol As Collection, r As Range, i As Integer

Set ol = GetObject(Class:="Outlook.Application")
Set olEmail = ol.CreateItem(0)
Set rCol = New Collection

With rCol
    .Add Sheets("Interface").Range("C15:M19")
End With

With olEmail
    .To = MID
    .Subject = MSubject
    .HTMLBody = "<html><body style=""font-family:calibri"">" & _
                "<p>Hello!<br><br>Thank you for Contacting us! We have changed the Status of your Issue/Request and below is the details of the same." & _
                "</p></body></html>"
    Set olInsp = .GetInspector
    If olInsp.EditorType = 4 Then
        Set wd = olInsp.WordEditor
        For i = 1 To rCol.Count
            Set r = rCol.Item(i): r.Copy
            wd.Range.InsertParagraphAfter
            wd.Paragraphs(wd.Paragraphs.Count).Range.PasteAndFormat 16
        Next
    End If
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "Please reply back or reach us for Concerns/Clarifications if any."
    wd.Range.InsertParagraphAfter
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "Regards - IT Helpdesk"
    wd.Paragraphs.Last.Range.Sentences.Last.Font.Bold = True
    wd.Range.InsertParagraphAfter
    wd.Paragraphs(wd.Paragraphs.Count).Range.Text = "Call: 555/044-4207 2489 | Mail: itsupport@comnet.org.in"
    wd.Paragraphs.Last.Range.Sentences.Last.Font.Bold = True
    .Display
    .Send
End With

End Sub
Sub StatusFilter()

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With
Sheets("Data").Visible = True

Dim WB As Workbook: Set WB = ThisWorkbook
Dim iFace As Worksheet: Set iFace = WB.Sheets("Interface")
Dim DataS As Worksheet: Set DataS = WB.Sheets("Data")
Dim iCriteria As String: iCriteria = "<>" & iFace.Range("Q22")
Dim DValue As String
Sheets.Add After:=Sheets("Data")
ActiveSheet.Name = "Temp"
Set TSheet = WB.Sheets("Temp")
DataS.Cells.Copy
TSheet.Range("A1").PasteSpecial xlValues
TSheet.Activate
TSheet.Range("1:1").AutoFilter 14, iCriteria

With ActiveSheet.AutoFilter.Range
    .Offset(1).Resize(.Rows.Count - 1).EntireRow.Delete
End With
TSheet.AutoFilterMode = False

CValue = TSheet.Range("A2").Value
If CValue = "" Then
    TSheet.Delete
    iFace.Activate
    Range("O25:Q34").ClearContents
    Range("Q22").Select
    MsgBox "Lucky! No Tickets are in this Criteria!!", vbInformation, "IT Helpdesk Manager"
    Exit Sub
End If

Range("A:A, C:H, J:J, L:P").EntireColumn.Delete
Range("a1").End(xlToRight).EntireColumn.Cut: Range("A1").EntireColumn.Insert
Range("A2:C11").Copy: iFace.Activate: Range("O25").PasteSpecial xlValues
Range("Q22").Select: TSheet.Delete

Application.ScreenUpdating = True: Application.DisplayAlerts = True
Sheets("Data").Visible = False
MsgBox "Remember! Only Top 10 Tickets in this Criteria is shown.", vbInformation, "IT Helpdesk Manager"

End Sub
Sub StatusMailer()

If MsgBox("Are you Sure to Publish the Status?", vbYesNo, vbInformation, "IT Helpdesk Manager") = vbNo Then
    MsgBox "Status Not Published.", vbInformation, "IT Helpdesk Manager"
    Exit Sub
End If

Const PR_ATTACH_MIME_TAG = "http://schemas.microsoft.com/mapi/proptag/0x370E001E"
Const PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
Dim xOutApp As Object
Dim xOutMail As Object
Dim xStartMsg As String
Dim xEndMsg As String
Dim xChartName As String
Dim xChartPath As String
Dim xPath As String
Dim xChart As ChartObject
On Error Resume Next
xChartName = "TicketTrend"
DynamicSubj = "IT Helpdesk Ticket Status as of " & Format(Now(), "DD/MM/YYYY")
Dim Recepient As String: Recepient = Sheets("Setup").Range("X1").Value
If xChartName = "" Then Exit Sub
Set xChart = Sheets("Interface").ChartObjects(xChartName)
If xChart Is Nothing Then Exit Sub
Set xOutApp = CreateObject("Outlook.Application")
Set xOutMail = xOutApp.CreateItem(0)
xStartMsg = "<font size='3' color='black'> Hello All!" & "<br> <br>" & "Please see below mentioned the IT Helpdesk Ticket Status as of " & Format(Now(), "DD/MM/YYYY") & "." & "<br> <br> </font>"
xEndMsg = "<font size='3' color='black'> Regards - IT Helpdesk" & "<br> <br> </font>"
xChartPath = ThisWorkbook.Path & "\" & "LastStatus.bmp"
xPath = "<p align='Left'><img src=" / "cid:" & "LastStatus.bmp" & "  width=700 height=500 > <br> <br>"
xChart.Chart.Export xChartPath, Filtername:="bmp"
With xOutMail
    .To = Recepient
    .Subject = DynamicSubj
    .Attachments.Add xChartPath, olByValue, 0
    .HTMLBody = xStartMsg & "<img src='cid:LastStatus.bmp'" & "width='800' height='300'><br>" & "<br>" & xEndMsg
    .Display
    .Send
End With
Kill xChartPath
Set xOutMail = Nothing
Set xOutApp = Nothing

MsgBox "Status has been Published.", vbInformation, "IT Helpdesk Manager"

End Sub
