Attribute VB_Name = "FormatData"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : Highlight
' Date : 6/19/2014
' Desc : Highlight items with more than one sub item in the second column
'---------------------------------------------------------------------------------------
Sub Highlight(Destination As Worksheet)
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim i As Long
    Dim j As Long


    Destination.Select
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column

    For i = 2 To TotalRows
        If Cells(i, 1).Value = "" Then
            Range(Cells(i, 1), Cells(i, TotalCols)).Interior.Color = RGB(255, 255, 0)

            j = i
            Do Until Cells(j, 1).Value <> ""
                j = j - 1
                Range(Cells(j, 1), Cells(j, TotalCols)).Interior.Color = RGB(255, 255, 0)
            Loop
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : CleanPiv
' Date : 6/20/2014
' Desc : Removes items without discrepancies from the table
'---------------------------------------------------------------------------------------
Sub CleanPiv(PivSheet As Worksheet)
    Dim TotalCols As Integer
    Dim TotalRows As Long
    Dim ColHeaders As Variant
    Dim i As Long
    Dim j As Long


    PivSheet.Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    TotalRows = ActiveSheet.UsedRange.Rows.Count
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Remove VMI orders that do not have any discrepancies
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="="
    ActiveSheet.UsedRange.AutoFilter Field:=2, Criteria1:="="
    ActiveSheet.UsedRange.AutoFilter Field:=3, Criteria1:="="
    ActiveSheet.UsedRange.AutoFilter Field:=4, Criteria1:="="

    'Remove remaining data
    Cells.Delete

    'Reinsert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders

    FilterData Field:=1, Operator:=xlFilterNoFill
    TotalRows = ActiveSheet.UsedRange.Rows.Count

    'Remove rows without discrepancies
    For i = TotalRows To 2 Step -1
        If Cells(i, 1).Interior.Color = 65535 And _
           Cells(i + 1, 1).Value <> "" And _
           Cells(i, 1).Value <> "" Then
            Rows(i).Delete
        End If

        If i = TotalRows And Cells(i, 1).Interior.Color = 65535 Then
            Rows(i).Delete
        End If

        If Cells(i, 1).Interior.Color = 16777215 Then
            Rows(i).Delete
        End If
    Next

    'Remove background colors
    With ActiveSheet.UsedRange.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

'---------------------------------------------------------------------------------------
' Proc : FilterData
' Date : 6/20/2014
' Desc : Filter and remove remaining data
'---------------------------------------------------------------------------------------
Private Sub FilterData(Field As Integer, Operator As XlAutoFilterOperator)
    Dim TotalCols As Integer
    Dim ColHeaders As Variant


    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value

    'Filter data
    ActiveSheet.UsedRange.AutoFilter Field:=Field, Operator:=Operator

    'Remove unfiltered data
    Cells.Delete

    'Reinsert column headers
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub
