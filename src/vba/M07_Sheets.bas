Attribute VB_Name = "M07_Sheets"
Option Explicit
'==========================================================
' M07_Sheets : 出力シートの自動生成
'  シートが無ければ、見出し・列幅・書式付きで新規作成する
'  既にあればそのまま返す（データ行のクリアは各処理側で行う）
'==========================================================

Public Function ResultSheet() As Worksheet
    Set ResultSheet = EnsureSheet(SH_RESULT, "OUT_照合結果（① で自動生成）", _
        "黄色の担当者コメント列は手入力。『除外』で始まる行は確認依頼の対象外。コメントは再実行しても引き継がれる", _
        Array("資産番号", "シリアル", "使用者社員番号", "使用者氏名", "所属", "台帳拠点", "検出ホスト名", _
              "最終通信日", "判定", "対応要否", "対応案", "優先度", "担当者コメント"), _
        Array(10, 11, 12, 13, 28, 12, 16, 12, 26, 8, 34, 7, 28), RGB(112, 173, 71))
End Function

Public Function LicenseSheet() As Worksheet
    Set LicenseSheet = EnsureSheet(SH_LICOUT, "OUT_ライセンス過不足（① で自動生成・数式で集計）", _
        "余剰＝保有数の20%以上が未使用（説明用の仮置き基準）。保有数・満了日は IN_ライセンス契約、利用数は IN_ソフト利用実績 を参照", _
        Array("ソフトウェア名", "保有数", "利用数", "過不足", "判定", "契約満了日", "満了まで(日)", "不足時の追加費用目安(円)"), _
        Array(32, 9, 9, 9, 10, 12, 12, 22), RGB(112, 173, 71))
End Function

Public Function RequestSheet() As Worksheet
    Set RequestSheet = EnsureSheet(SH_REQ, "OUT_部門別確認依頼（② で自動生成）", _
        "所属長へ送付し、黄色の回答欄の記入を依頼する", _
        Array("所属", "資産番号", "使用者氏名", "確認事項", "依頼内容", "回答期限", "部門回答", "回答日"), _
        Array(30, 10, 13, 26, 36, 12, 24, 12), RGB(112, 173, 71))
End Function

Public Function SummarySheet() As Worksheet
    Set SummarySheet = EnsureSheet(SH_SUMMARY, "OUT_サマリ（① で自動生成・課長報告用）", "", _
        Array("差異区分", "件数", "備考"), Array(30, 12, 32), RGB(112, 173, 71))
End Function

Public Function LogSheet() As Worksheet
    Set LogSheet = EnsureSheet(SH_LOG, "LOG_実行ログ（自動生成・自動追記）", "", _
        Array("日時", "実行者", "処理", "結果", "メッセージ"), Array(18, 14, 22, 10, 70), RGB(165, 165, 165))
End Function

Public Function SheetExists(ByVal nm As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then SheetExists = True: Exit Function
    Next ws
End Function

Public Function EnsureSheet(ByVal nm As String, ByVal ttl As String, ByVal subTtl As String, _
                            ByVal headers As Variant, ByVal widths As Variant, ByVal tabColor As Long) As Worksheet
    Dim ws As Worksheet, j As Long, hdr As Range

    If SheetExists(nm) Then
        Set EnsureSheet = ThisWorkbook.Worksheets(nm)
        Exit Function
    End If

    If nm <> SH_LOG And SheetExists(SH_LOG) Then
        Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(SH_LOG))   ' ログは常に末尾
    Else
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    End If
    ws.Name = nm
    ws.Range("A1").Value = ttl
    ws.Range("A2").Value = subTtl
    For j = 0 To UBound(headers)
        ws.Cells(HEADER_ROW, j + 1).Value = headers(j)
    Next j
    Set hdr = ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, UBound(headers) + 1))

    ' 見た目の設定（環境によって未対応の項目があっても処理は止めない）
    On Error Resume Next
    ws.Cells.Font.Name = "Meiryo UI"
    ws.Tab.Color = tabColor
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    ws.Range("A2").Font.Size = 9
    ws.Range("A2").Font.Color = RGB(89, 89, 89)
    With hdr
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = IIf(nm = SH_LOG, RGB(89, 89, 89), RGB(84, 130, 53))
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
    End With
    ws.Rows(HEADER_ROW).RowHeight = 30
    For j = 0 To UBound(widths)
        ws.Columns(j + 1).ColumnWidth = widths(j)
    Next j
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.FreezePanes = False
    ws.Range("A" & FIRST_ROW).Select
    ActiveWindow.FreezePanes = True
    hdr.AutoFilter
    On Error GoTo 0

    Set EnsureSheet = ws
    If nm <> SH_LOG Then WriteLog "EnsureSheet", "生成", nm & " を新規作成"
End Function
