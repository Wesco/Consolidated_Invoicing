Attribute VB_Name = "Imports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ImportReports
' Date : 6/18/2014
' Desc : Import consolidated invoicing reports
'---------------------------------------------------------------------------------------
Sub ImportReports(Path As String)
    Dim File As String
    Dim TotalRows As Long
    Dim TotalCols As Integer
    Dim ColHeaders As Variant


    If Right(Path, 1) <> "\" Then Path = Path + "\"
    If Not FolderExists(Path) Then RecMkDir Path

    File = Dir(Path)
    If File = "" Then Err.Raise CustErr.EMPTYFOLDER, "Combine", Path & " is empty."

    Do While File <> ""
        Workbooks.Open Path & File

        If TotalRows > 0 Then
            TotalRows = ThisWorkbook.Sheets("Combined").UsedRange.Rows.Count + 1
        Else
            TotalRows = ThisWorkbook.Sheets("Combined").UsedRange.Rows.Count
        End If

        ActiveSheet.UsedRange.Copy Destination:=ThisWorkbook.Sheets("Combined").Range("A" & TotalRows)
        ActiveWorkbook.Close
        File = Dir()
    Loop

    Sheets("Combined").Select
    TotalCols = Columns(Columns.Count).End(xlToLeft).Column
    ColHeaders = Range(Cells(1, 1), Cells(1, TotalCols)).Value
    ActiveSheet.UsedRange.AutoFilter Field:=1, Criteria1:="Cust#", Operator:=xlFilterValues
    Cells.Delete
    Rows(1).Insert
    Range(Cells(1, 1), Cells(1, TotalCols)).Value = ColHeaders
End Sub

