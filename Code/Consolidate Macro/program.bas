Attribute VB_Name = "program"
'---------------------------------------------------------------------------------------
' Module    : Program
' Author    : TReische
' Date      : 6/18/2014
' Purpose   : Combine consolidated invoicing reports and provide
'           : pricing and supplier reports
'---------------------------------------------------------------------------------------

Option Explicit

Public Const VersionNumber = "1.0.2"
Public Const RepositoryName = "Consolidated_Invoicing"

'---------------------------------------------------------------------------------------
' Type : Enum
' Date : 6/20/2014
' Desc : List of custom errors
'---------------------------------------------------------------------------------------
Enum CustErr
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
    Dim ImportPath As String
    Dim ExportPath As String


    ImportPath = Environ("USERPROFILE") & "\My Documents\Consolidated Spend Report Emails\"
    ExportPath = Environ("USERPROFILE") & "\My Documents\Consolidated Spend Reports\"

    On Error GoTo ERR_HANDLER
    ImportReports Path:=ImportPath
    CheckHeaders
    ReportDate = Format(Sheets("Combined").Range("E2").Value, "mmm yyyy")
    CreatePiv Sheets("Discrepancy"), Array("Stock Code", "2nd Tier Supplier", "Price", "Description", "VMI Order #")
    Highlight Sheets("Discrepancy")
    CleanPiv Sheets("Discrepancy")

    ExportSheet "Discrepancy", ExportPath, "Discrepancy Report " & ReportDate & ".xlsx", xlOpenXMLWorkbook
    ExportSheet "Combined", ExportPath, "Consolidated Report " & ReportDate & ".csv", xlCSV
    DeleteFiles Path:=ImportPath
    On Error GoTo 0
    Clean
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
