Attribute VB_Name = "Module1"
Public SMEName, EndingPhrase, ActualEndPhrase, FindPhrase, ReplacePhrase
Global StartTime As Date
Sub Scrub()

Dim RowADD As Integer
Dim SourceFileName, SourceFilePath, EndRange, NewSheet, StartRange, StopRange As String
Dim I As Long

StartTime = Now()

With Application
.ScreenUpdating = False
.DisplayAlerts = False
.EnableCancelKey = False
.EnableEvents = False
End With

MainFN = ActiveWorkbook.Name

'Validating Input data
Sheets("RawData").Select
Application.StatusBar = "Validating Input data.."
If Range("A2").Value = "" Then
    Sheets("Index").Select
    MsgBox "There is no data to process.", vbOKOnly, "Input Data"
    Exit Sub
    Else
End If

Sheets("RawData").Select
TotalRecords = Range("C:C").Count - 1
Range("C:C").Copy
Range("D1").PasteSpecial xlValues
Application.CutCopyMode = False

'Cleaning Salutations
Application.StatusBar = "Cleaning Salutations.."
Call CleanIt
Application.CutCopyMode = False

'Identifying SME Names
Application.StatusBar = "Identifying SME Names.."
Sheets("Reference").Select
Range("D2").Select
ICount = 0

Do Until ActiveCell.Value = ""

I = ICount + 1
EndingPhrase = ActiveCell.Value
ActualEndPhrase = ActiveCell.Offset(0, 1).Value
Sheets("RawData").Select
Range("1:1").Select
Selection.AutoFilter Field:=3, Criteria1:="<>*|*"
Range("D:D").SpecialCells(xlCellTypeVisible).Select
Selection.Replace What:=EndingPhrase, Replacement:=ActualEndPhrase, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
ActiveSheet.AutoFilterMode = False
Sheets("Reference").Select
ActiveCell.Offset(1, 0).Select
ICount = I + 1
Application.StatusBar = "Identified " & ICount & " Keywords.."

Loop

'Deleting Suffixes
Application.StatusBar = "Deleting Suffixes.."
Sheets("RawData").Select
Columns("D:D").Select
Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
    TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
    Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
    :="|", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
Columns("E:E").Select
Range(Selection, Selection.End(xlToRight)).Select
Selection.Delete Shift:=xlToLeft
Range("D:D").Select

'Sorting Data
Application.StatusBar = "Sorting Data.."
Selection.Sort Key1:=Range("D1"), Order1:=xlAscending
Range("D1").Value = "Scrubbed Company Name"
Range("A1").Select

'Confrmation Message
Application.StatusBar = False
Range("Index").Select
Range("A1").Select
EndTime = Format((Now() - StartTime), "HH:MM:SS")

MsgBox "Data Scrubbed Succesfully in " & EndTime, vbOKOnly, "SME Name Scrubber"

End Sub

Sub CleanIt()

Dim FindPhrase, ReplacePhrase As String

Sheets("Reference").Select
Range("A2").Select

Do Until ActiveCell.Value = ""

FindPhrase = ActiveCell.Value
ReplacePhrase = ActiveCell.Offset(0, 1).Value
Sheets("RawData").Select
Range("D:D").Select
Selection.Replace What:=FindPhrase, Replacement:=ReplacePhrase, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False
Sheets("Reference").Select
ActiveCell.Offset(1, 0).Select

Loop

Sheets("RawData").Select
Range("D2").Select

End Sub
