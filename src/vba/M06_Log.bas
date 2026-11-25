Attribute VB_Name = "M06_Log"
Option Explicit
'==========================================================
' M06_Log : 実行ログと保管
'==========================================================

Public Sub WriteLog(ByVal procName As String, ByVal result As String, Optional ByVal msg As String = "")
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = LogSheet()
    If ws Is Nothing Then Exit Sub
    r = LastRow(ws) + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 1).NumberFormat = "yyyy/mm/dd hh:mm:ss"
    ws.Cells(r, 2).Value = Environ$("USERNAME")
    ws.Cells(r, 3).Value = procName
    ws.Cells(r, 4).Value = result
    ws.Cells(r, 5).Value = msg
    FormatBody ws.Range(ws.Cells(r, 1), ws.Cells(r, 5))
End Sub

' ブックと同じフォルダの「棚卸保管」に、日時付きのコピーを保存する
Public Sub WriteLogAndArchive(ByVal silent As Boolean)
    Dim folder As String, sep As String, fname As String, baseName As String
    sep = Application.PathSeparator
    If ThisWorkbook.Path = "" Then
        WriteLog "WriteLogAndArchive", "スキップ", "ブックが未保存のため保管コピーを作成しませんでした"
        Exit Sub
    End If
    folder = ThisWorkbook.Path & sep & "棚卸保管"
    If Dir(folder, vbDirectory) = "" Then MkDir folder
    baseName = ThisWorkbook.Name
    If InStrRev(baseName, ".") > 0 Then baseName = Left$(baseName, InStrRev(baseName, ".") - 1)
    fname = folder & sep & baseName & "_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsm"
    WriteLog "WriteLogAndArchive", "完了", "保管：" & fname
    ThisWorkbook.SaveCopyAs fname
End Sub
