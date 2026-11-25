Attribute VB_Name = "M01_Import"
Option Explicit
'==========================================================
' M01_Import : インプットExcelの取込・整形
'  インプットExcelの各シート（1行目が見出し）を、見出し名で列を対応付けて
'  ツール側の IN_ シート（4行目が見出し）へ取り込む
'==========================================================

' 戻り値：処理を続行する場合 True
Public Function ImportAndNormalize(ByVal silent As Boolean, Optional ByVal inputPath As String = "") As Boolean
    Dim f As Variant, ans As VbMsgBoxResult, cnt As Long

    If Len(inputPath) = 0 And Not silent Then
        ans = MsgBox("インプットExcelを選択して取り込みますか？" & vbCrLf & vbCrLf & _
                     "［はい］　　 ファイルを選んで取り込む" & vbCrLf & _
                     "［いいえ］　 各 IN_ シートの現在のデータで照合する" & vbCrLf & _
                     "［キャンセル］中止", vbYesNoCancel + vbQuestion, "データ取込")
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then
            f = Application.GetOpenFilename("Excelファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "インプットExcelを選択")
            If VarType(f) <> vbString Then Exit Function
            inputPath = CStr(f)
        End If
    End If

    If Len(inputPath) > 0 Then cnt = ImportWorkbook(inputPath)

    NormalizeIds
    If DataCount(SH_ASSET) = 0 Then Err.Raise vbObjectError + 4, , "IN_資産台帳 にデータがありません。インプットExcelを取り込んでください。"

    Dim src As String
    If cnt > 0 Then
        src = "取込 " & cnt & " シート（" & Mid$(inputPath, InStrRev(inputPath, Application.PathSeparator) + 1) & "）"
    Else
        src = "取込なし（既存データで照合）"
    End If
    WriteLog "ImportAndNormalize", "完了", src & _
             "／人事 " & DataCount(SH_HR) & "・台帳 " & DataCount(SH_ASSET) & _
             "・インベントリ " & DataCount(SH_INV) & "・契約 " & DataCount(SH_LIC) & "・利用実績 " & DataCount(SH_USE)
    ImportAndNormalize = True
End Function

Public Function DataCount(ByVal sheetName As String) As Long
    DataCount = LastRow(ThisWorkbook.Worksheets(sheetName)) - HEADER_ROW
End Function

Private Function ImportWorkbook(ByVal path As String) As Long
    Dim wb As Workbook, map As Variant, i As Long, n As Long, msg As String
    ' ツール側シート名, インプットExcel側シート名
    map = Array(SH_HR, "人事マスタ", SH_ASSET, "資産台帳", SH_INV, "インベントリ", _
                SH_LIC, "ライセンス契約", SH_USE, "ソフト利用実績")
    Set wb = Workbooks.Open(Filename:=path, ReadOnly:=True)
    On Error GoTo EH
    For i = 0 To UBound(map) Step 2
        CopyByHeader FindSheet(wb, CStr(map(i + 1))), ThisWorkbook.Worksheets(CStr(map(i)))
        n = n + 1
    Next i
    wb.Close SaveChanges:=False
    ImportWorkbook = n
    Exit Function
EH:
    msg = Err.Description
    wb.Close SaveChanges:=False
    Err.Raise vbObjectError + 2, , msg
End Function

Private Function FindSheet(ByVal wb As Workbook, ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Name = nm Then Set FindSheet = ws: Exit Function
    Next ws
    Err.Raise vbObjectError + 5, , "インプットExcelに『" & nm & "』シートがありません。"
End Function

' 見出し名で列を対応付けてコピー（インプット側の列順・余分な列は問わない）
Private Sub CopyByHeader(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim nCols As Long, lc As Long, lr As Long, j As Long, k As Long, r As Long
    Dim v As Variant, out() As Variant, colMap() As Long, missing As String

    nCols = 0
    Do While Len(CStr(dst.Cells(HEADER_ROW, nCols + 1).Value)) > 0
        nCols = nCols + 1
    Loop
    lc = 0
    Do While Len(CStr(src.Cells(1, lc + 1).Value)) > 0
        lc = lc + 1
    Loop
    lr = 1
    Do While Len(CStr(src.Cells(lr + 1, 1).Value)) > 0
        lr = lr + 1
    Loop
    If lc = 0 Then Err.Raise vbObjectError + 6, , "『" & src.Name & "』の1行目に見出しがありません。"

    v = src.Range(src.Cells(1, 1), src.Cells(lr, lc)).Value
    ReDim colMap(1 To nCols)
    For j = 1 To nCols
        For k = 1 To lc
            If Trim$(CStr(v(1, k))) = Trim$(CStr(dst.Cells(HEADER_ROW, j).Value)) Then colMap(j) = k: Exit For
        Next k
        If colMap(j) = 0 Then missing = missing & "『" & dst.Cells(HEADER_ROW, j).Value & "』"
    Next j
    If Len(missing) > 0 Then Err.Raise vbObjectError + 7, , "『" & src.Name & "』に必要な列がありません：" & missing

    ClearBody dst, nCols
    If lr < 2 Then Exit Sub
    ReDim out(1 To lr - 1, 1 To nCols)
    For r = 2 To lr
        For j = 1 To nCols
            out(r - 1, j) = v(r, colMap(j))
        Next j
    Next r
    dst.Range(dst.Cells(FIRST_ROW, 1), dst.Cells(FIRST_ROW + lr - 2, nCols)).Value = out
    ' 日付などの表示形式をインプット側から引き継ぐ
    On Error Resume Next
    For j = 1 To nCols
        dst.Range(dst.Cells(FIRST_ROW, j), dst.Cells(FIRST_ROW + lr - 2, j)).NumberFormat = src.Cells(2, colMap(j)).NumberFormat
    Next j
    On Error GoTo 0
    FormatBody dst.Range(dst.Cells(FIRST_ROW, 1), dst.Cells(FIRST_ROW + lr - 2, nCols))
End Sub

' 突合キー（社員番号・シリアル）の前後空白・全角・小文字を整える
Private Sub NormalizeIds()
    CleanColumn SH_HR, 1
    CleanColumn SH_ASSET, 1
    CleanColumn SH_ASSET, 5
    CleanColumn SH_ASSET, 6
    CleanColumn SH_INV, 2
    CleanColumn SH_INV, 3
    CleanColumn SH_USE, 1
End Sub

Private Sub CleanColumn(ByVal sheetName As String, ByVal col As Long)
    Dim ws As Worksheet, r As Long, v As Variant
    Set ws = ThisWorkbook.Worksheets(sheetName)
    For r = FIRST_ROW To LastRow(ws)
        v = ws.Cells(r, col).Value
        If VarType(v) = vbString Then
            v = UCase(StrConv(Trim(v), vbNarrow))
            If ws.Cells(r, col).Value <> v Then ws.Cells(r, col).Value = v
        End If
    Next r
End Sub
