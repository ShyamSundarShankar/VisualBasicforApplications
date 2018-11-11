Sub SummaryPivot()

Sheets("DashboardSummary").Select
ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Sheets("DashboardRawData").Range("A1").CurrentRegion, Version:=xlPivotTableVersion14).CreatePivotTable TableDestination:=Sheets("DashboardSummary").Range("A1"), TableName:="DashboardSummary", DefaultVersion:=xlPivotTableVersion14
Set PvtTbl = Sheets("DashboardSummary").PivotTables("DashboardSummary")

'Alligning Required Fields
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Go-Live Month")
    .Orientation = xlRowField
    .Position = 1
End With
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Vertical")
    .Orientation = xlRowField
    .Position = 2
End With
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Process")
    .Orientation = xlRowField
    .Position = 3
End With
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Practice Name")
    .Orientation = xlRowField
    .Position = 4
End With
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Facility")
    .Orientation = xlRowField
    .Position = 5
End With
With Sheets("DashboardSummary").PivotTables("DashboardSummary").PivotFields("Status")
    .Orientation = xlRowField
    .Position = 6
End With

'Alligning Required Values
Sheets("DashboardSummary").PivotTables("DashboardSummary").AddDataField ActiveSheet.PivotTables("DashboardSummary").PivotFields("Process"), "####", xlCount

'Summarizing Pivot
With PvtTbl
 .RepeatAllLabels Repeat:=xlRepeatLabels
 .RowAxisLayout xlTabularRow
 .ColumnGrand = False
 .RowGrand = False
 .PivotFields("Go-Live Month").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
 .PivotFields("Vertical").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
 .PivotFields("Process").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
 .PivotFields("Practice Name").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
 .PivotFields("Facility").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
 .PivotFields("Status").Subtotals(1) = Array(False, False, False, False, False, False, False, False, False, False, False, False)
End With

ActiveWorkbook.ShowPivotTableFieldList = False
Range("A1").CurrentRegion.Copy
Range("A1").PasteSpecial xlValues
Range("A1").Select
ActiveSheet.ListObjects.Add(xlSrcRange, ActiveCell.CurrentRegion).Name = "tblSummaryData"

End Sub
