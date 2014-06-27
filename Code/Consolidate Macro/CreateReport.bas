Attribute VB_Name = "CreateReport"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : CreatePiv
' Date : 6/18/2014
' Desc : Create a pivot table out of the combined data
'---------------------------------------------------------------------------------------
Sub CreatePiv(Destination As Worksheet, PivFields As Variant)
    Dim TotalRows As Long
    Dim NoSubs As Variant
    Dim i As Integer


    Sheets("Combined").Select
    NoSubs = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    TotalRows = Rows(Rows.Count).End(xlUp).Row

    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, _
                                      SourceData:=Range(Cells(1, 1), Cells(TotalRows, 16)), _
                                      Version:=xlPivotTableVersion14).CreatePivotTable _
                                      TableDestination:=Destination.Range("A1"), _
                                      TableName:="PivotTable1", _
                                      DefaultVersion:=xlPivotTableVersion14
    Destination.Select
    Range("A1").Select

    With ActiveSheet.PivotTables("PivotTable1")
        For i = 0 To UBound(PivFields)
            .PivotFields(PivFields(i)).Orientation = xlRowField
            .PivotFields(PivFields(i)).Position = i + 1
            .PivotFields(PivFields(i)).Subtotals = NoSubs
            .PivotFields(PivFields(i)).LayoutForm = xlTabular
        Next
        .ColumnGrand = False
    End With

    ActiveSheet.UsedRange.Copy
    Range("A1").PasteSpecial xlPasteValues
    Range("A1").Select
End Sub

