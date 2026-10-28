Attribute VB_Name = "LiveSync"
'==============================================================
' Excel Live Sync (demo)
' Watches a local Excel/CSV source file and refreshes the
' Power Query table in this workbook a few seconds after the
' source is saved. No buttons, no Windows scheduled tasks.
'
' How it works: every CHECK_SECONDS the macro reads the source
' file's "last saved" time. When it changes, all queries are
' refreshed and the run is written to the Log sheet.
'==============================================================
Option Explicit

Private Const CHECK_SECONDS As Integer = 2
Private Const MAX_LOG_ROWS As Long = 500

Public NextCheck As Date
Private LastStamp As Date
Private WatchedPath As String
Private Running As Boolean

' Called from Workbook_Open
Public Sub StartWatcher()
    WatchedPath = GetSourcePath()
    If FileExists(WatchedPath) Then
        LastStamp = FileDateTime(WatchedPath)
        LogEvent "Watcher started", WatchedPath
    Else
        LastStamp = 0
        LogEvent "Source not found (will keep checking)", WatchedPath
    End If
    Running = True
    ScheduleNext
End Sub

' Called from Workbook_BeforeClose
Public Sub StopWatcher()
    Running = False
    On Error Resume Next
    Application.OnTime EarliestTime:=NextCheck, Procedure:=ProcName(), Schedule:=False
    On Error GoTo 0
End Sub

Private Sub ScheduleNext()
    If Not Running Then Exit Sub
    NextCheck = Now + TimeSerial(0, 0, CHECK_SECONDS)
    Application.OnTime EarliestTime:=NextCheck, Procedure:=ProcName()
End Sub

Private Function ProcName() As String
    ProcName = "'" & ThisWorkbook.Name & "'!LiveSync.CheckSource"
End Function

Public Sub CheckSource()
    Dim p As String
    Dim s As Date
    On Error GoTo Fail

    p = GetSourcePath()
    If p <> WatchedPath Then
        ' Source path changed in Settings: start watching the new file
        WatchedPath = p
        LastStamp = 0
        LogEvent "Source path changed", p
    End If

    If FileExists(p) Then
        s = FileDateTime(p)
        If s <> LastStamp Then
            RefreshData
            ' Only remember the new time after a successful refresh,
            ' so a file that was still being written is retried
            LastStamp = s
            ThisWorkbook.Worksheets("Settings").Range("LastRefresh").Value = Now
            LogEvent "Refreshed", "Source saved at " & Format(s, "yyyy-mm-dd hh:nn:ss")
        End If
    End If

    ScheduleNext
    Exit Sub

Fail:
    LogEvent "Error (will retry)", Err.Description
    Err.Clear
    ScheduleNext
End Sub

Private Sub RefreshData()
    Dim c As WorkbookConnection
    Dim prevScreen As Boolean
    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    For Each c In ThisWorkbook.Connections
        If c.Type = xlConnectionTypeOLEDB Then
            ' Wait for each refresh to finish before continuing
            c.OLEDBConnection.BackgroundQuery = False
        End If
        c.Refresh
    Next c
    Application.ScreenUpdating = prevScreen
End Sub

' Reads the SourcePath named cell. A file name alone is treated as
' "same folder as this workbook", so both files can be moved together.
Public Function GetSourcePath() As String
    Dim raw As String
    raw = Trim$(CStr(ThisWorkbook.Worksheets("Settings").Range("SourcePath").Value))
    If InStr(raw, ":\") > 0 Or Left$(raw, 2) = "\\" Then
        GetSourcePath = raw
    Else
        GetSourcePath = ThisWorkbook.Path & Application.PathSeparator & raw
    End If
End Function

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(p) > 0 And Len(Dir(p)) > 0)
    On Error GoTo 0
End Function

Private Sub LogEvent(ByVal action As String, ByVal detail As String)
    Dim ws As Worksheet
    Dim r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Log")
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r > MAX_LOG_ROWS + 1 Then
        ws.Rows(2).Delete
        r = r - 1
    End If
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = action
    ws.Cells(r, 3).Value = detail
End Sub
