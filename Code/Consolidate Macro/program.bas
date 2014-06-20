Attribute VB_Name = "program"
'---------------------------------------------------------------------------------------
' Module    : Program
' Author    : TReische
' Date      : 6/18/2014
' Purpose   : Combine consolidated invoicing reports and provide
'           : pricing and supplier reports
'---------------------------------------------------------------------------------------

Option Explicit

Public Const VersionNumber = "1.0.1"
Public Const RepositoryName = "Consolidated_Invoicing"

'---------------------------------------------------------------------------------------
' Type : Enum
' Date : 6/20/2014
' Desc : List of custom errors
'---------------------------------------------------------------------------------------
Private Enum CustErr
    EMPTYFOLDER = 50000
    REPORTCHANGED = 50001
End Enum

'---------------------------------------------------------------------------------------
' Proc : Main
' Date : 6/19/2014
' Desc : Combine consolidated reports, remove files, and export two reports
'---------------------------------------------------------------------------------------
Sub Main()
    Dim ReportDate As String
    Dim UserProfile As String
    Dim CombinePath As String
    Dim ExportPath As String
    
    
    CombinePath = Environ("USERPROFILE") & "\My Documents\Consolidated Spend Report Emails\"
    ExportPath = Environ("USERPROFILE") & "\My Documents\Consolidated Spend Reports\"

    On Error GoTo ERR_HANDLER
    Combine Path:=CombinePath
    CheckHeaders
    ReportDate = Format(Sheets("Combined").Range("E2").Value, "mmm yyyy")
    CreatePiv Sheets("Discrepancy"), Array("Stock Code", "2nd Tier Supplier", "Price", "Description", "VMI Order #")
    Highlight Sheets("Discrepancy")
    CleanPiv Sheets("Discrepancy")

    ExportSheet "Discrepancy", ExportPath & "Discrepancy Report " & ReportDate & ".xlsx", xlOpenXMLWorkbook
    ExportSheet "Combined", ExportPath & "Consolidated Report " & ReportDate & ".csv", xlCSV
    DeleteFiles Path:=CombinePath
    On Error GoTo 0
    Exit Sub

ERR_HANDLER:
    If Err.Number = CustErr.EMPTYFOLDER Then
        MsgBox "No files found.", vbOKOnly, "Macro Aborted"
    ElseIf Err.Number = 1004 Then
        ThisWorkbook.Activate
        Resume Next
    Else
        MsgBox "Error " & Err.Number & " occurred in " & Err.Source & "." & vbCrLf & Err.Description, vbOKOnly, "Macro Aborted"
    End If

    Clean
End Sub

'---------------------------------------------------------------------------------------
' Proc : CheckHeaders
' Date : 6/20/2014
' Desc : Checks headers to make sure nothing has changed
'---------------------------------------------------------------------------------------
Private Sub CheckHeaders()
    Dim ColHeaders As Variant
    Dim i As Integer
    
    ColHeaders = Array("Cust#", "Plant", "2nd Tier Supplier", "Contract#", "Invoice Date", _
                       "VMI Order #", "Order Line", "Stock Code", "Description", "Qty", "Price", _
                       "Extended Price", "Invoice Number", "2nd Tier Supplier Invoice#", _
                       "2nd Tier Supplier Inv date", "Packing List No.")
                       
    For i = 1 To UBound(ColHeaders)
        If Cells(1, i).Value <> ColHeaders(i - 1) Then
            Err.Raise CustErr.REPORTCHANGED, "CheckHeaders", "The consolidated invoice report has changed."
        End If
    Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : Combine
' Date : 6/18/2014
' Desc : Combine consolidated invoicing reports
'---------------------------------------------------------------------------------------
Private Sub Combine(Path As String)
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

'---------------------------------------------------------------------------------------
' Proc : CleanPiv
' Date : 6/20/2014
' Desc : Removes items without discrepancies from the table
'---------------------------------------------------------------------------------------
Private Sub CleanPiv(PivSheet As Worksheet)
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

'---------------------------------------------------------------------------------------
' Proc : CreatePiv
' Date : 6/18/2014
' Desc : Create a pivot table out of the combined data
'---------------------------------------------------------------------------------------
Private Sub CreatePiv(Destination As Worksheet, PivFields As Variant)
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

'---------------------------------------------------------------------------------------
' Proc : Highlight
' Date : 6/19/2014
' Desc : Highlight items with more than one sub item in the second column
'---------------------------------------------------------------------------------------
Private Sub Highlight(Destination As Worksheet)
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
' Proc : ExportSheets
' Date : 6/19/2014
' Desc : Copies sheets to a new workbook and saves it
'---------------------------------------------------------------------------------------
Private Sub ExportSheet(SheetName As Variant, FileName As String, SaveType As XlFileFormat)
    Dim DispAlerts As Boolean
    Dim File As Variant

    DispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(SheetName).Copy
    ActiveWorkbook.SaveAs FileName, SaveType
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close

    Application.DisplayAlerts = DispAlerts
End Sub

'---------------------------------------------------------------------------------------
' Proc : DeleteFiles
' Date : 6/18/2014
' Desc : Deletes all files in the specified folder
'---------------------------------------------------------------------------------------
Private Sub DeleteFiles(Path As String)
    Dim File As String


    File = Dir(Path)
    Do While File <> ""
        On Error GoTo DELETE_FAILED
        Kill Path & File
        On Error GoTo 0
        File = Dir()
    Loop
    Exit Sub

DELETE_FAILED:
    MsgBox "Failed to delete " & File, vbOKOnly, "Delete Failed"
    Resume Next
End Sub

'---------------------------------------------------------------------------------------
' Proc : Clean
' Date : 6/18/2014
' Desc : Removes any data produced by the macro at run time
'---------------------------------------------------------------------------------------
Sub Clean()
    Dim s As Worksheet


    ThisWorkbook.Activate
    For Each s In ThisWorkbook.Sheets
        If s.Name <> "Macro" Then
            s.Select
            s.AutoFilterMode = False
            s.Cells.Delete
            s.Range("A1").Select
        End If
    Next

    Sheets("Macro").Select
    Range("C7").Select
End Sub
