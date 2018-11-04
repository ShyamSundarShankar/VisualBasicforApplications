Attribute VB_Name = "Module2"
Sub SlipSave()
'Saves active worksheet as pdf using rep name

With ActiveSheet
    .ExportAsFixedFormat Type:=xlTypePDF, filename:=Range("H1").Value, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    
End With
End Sub

Sub PREEmp()

Range("E8").Select
ActiveCell.Value = ActiveCell.Value - 1

End Sub
Sub NXTEmp()

Range("E8").Select
ActiveCell.Value = ActiveCell.Value + 1

End Sub

Sub Send_Files()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim sh As Worksheet
    Dim cell As Range
    Dim FileCell As Range
    Dim rng As Range

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set sh = Sheets("Controls")

    Set OutApp = CreateObject("Outlook.Application")

    For Each cell In sh.Columns("B").Cells.SpecialCells(xlCellTypeConstants)

        'Enter the path/file names in the C:Z column in each row
        Set rng = sh.Cells(cell.Row, 1).Range("e2:Z2")

        If cell.Value Like "?*@?*.?*" And _
           Application.WorksheetFunction.CountA(rng) > 0 Then
            Set OutMail = OutApp.CreateItem(0)

            With OutMail
                .to = cell.Value
                .Subject = "Testfile"
                .Body = "Hi " & cell.Offset(0, -1).Value

                For Each FileCell In rng.SpecialCells(xlCellTypeConstants)
                    If Trim(FileCell) <> "" Then
                        If Dir(FileCell.Value) <> "" Then
                            .Attachments.Add FileCell.Value
                        End If
                    End If
                Next FileCell

                .Display  'Or use .Display
            End With

            Set OutMail = Nothing
        End If
    Next cell

    Set OutApp = Nothing
    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With
End Sub

