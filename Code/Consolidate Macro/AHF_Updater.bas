Attribute VB_Name = "AHF_Updater"
Option Explicit

Private Declare PtrSafe Function ShellExecute _
        Lib "shell32.dll" Alias "ShellExecuteA" ( _
        ByVal hWnd As Long, _
        ByVal Operation As String, _
        ByVal FileName As String, _
        Optional ByVal Parameters As String, _
        Optional ByVal Directory As String, _
        Optional ByVal WindowStyle As Long = vbMaximizedFocus _
      ) As Long

'---------------------------------------------------------------------------------------
' Proc : CheckForUpdates
' Date : 4/24/2013
' Desc : Checks to see if the macro is up to date
'---------------------------------------------------------------------------------------
Sub CheckForUpdates(RepoName As String, LocalVer As String)
    Dim RemoteVer As Variant
    Dim RegEx As Variant
    Dim Result As Integer

    On Error GoTo UPDATE_ERROR
    Set RegEx = CreateObject("VBScript.RegExp")

    'Try to get the contents of the text file
    RemoteVer = DownloadTextFile("https://raw.github.com/Wesco/" & RepoName & "/master/Version.txt")
    RemoteVer = Replace(RemoteVer, vbLf, "")
    RemoteVer = Replace(RemoteVer, vbCr, "")

    'Expression to verify the data retrieved is a version number
    RegEx.Pattern = "^[0-9]+\.[0-9]+\.[0-9]+$"

    If RegEx.Test(RemoteVer) Then
        If Not RemoteVer = LocalVer Then
            Result = MsgBox("An update is available. Would you like to download the latest version now?", vbYesNo, "Update Available")
            If Result = vbYes Then
                'Opens github release page in the default browser, maximised with focus by default
                ShellExecute 0, "Open", "http://github.com/Wesco/" & RepoName & "/releases/"
                ThisWorkbook.Saved = True
                If Workbooks.Count = 1 Then
                    Application.Quit
                Else
                    ThisWorkbook.Close
                End If
            End If
        End If
    End If
    On Error GoTo 0
    Exit Sub

UPDATE_ERROR:
    If MsgBox("An error occured while checking for updates." & vbCrLf & vbCrLf & _
              "Would you like to open the website to download the latest version?", vbYesNo) = vbYes Then
        ShellExecute 0, "Open", "http://github.com/Wesco/" & RepoName & "/releases/"
        If Workbooks.Count = 1 Then
            Application.Quit
        Else
            ThisWorkbook.Close
        End If
    End If
End Sub

'---------------------------------------------------------------------------------------
' Proc : DownloadTextFile
' Date : 4/25/2013
' Desc : Returns the contents of a text file from a website
'---------------------------------------------------------------------------------------
Private Function DownloadTextFile(URL As String) As String
    Dim success As Boolean
    Dim responseText As String
    Dim oHTTP As Variant

    Set oHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")

    oHTTP.Open "GET", URL, False
    oHTTP.Send
    success = oHTTP.WaitForResponse()

    If Not success Then
        DownloadTextFile = ""
        Exit Function
    End If

    responseText = oHTTP.responseText
    Set oHTTP = Nothing

    DownloadTextFile = responseText
End Function
