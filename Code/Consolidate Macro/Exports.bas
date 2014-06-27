Attribute VB_Name = "Exports"
Option Explicit

'---------------------------------------------------------------------------------------
' Proc : ExportSheets
' Date : 6/19/2014
' Desc : Copies sheets to a new workbook and saves it
'---------------------------------------------------------------------------------------
Sub ExportSheet(SheetName As Variant, Path As String, FileName As String, SaveType As XlFileFormat)
    Dim DispAlerts As Boolean
    Dim File As Variant

    DispAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    
    If Not FolderExists(Path) Then
        RecMkDir Path
    End If
    
    If Right(Path, 1) <> "\" Then
        Path = Path + "\"
    End If

    ThisWorkbook.Activate
    ThisWorkbook.Sheets(SheetName).Copy
    ActiveWorkbook.SaveAs Path & FileName, SaveType
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close

    Application.DisplayAlerts = DispAlerts
End Sub
