Attribute VB_Name = "M02_Match"
Option Explicit
'==========================================================
' M02_Match : 突合・差異判定
'  資産台帳を起点に、インベントリ・人事マスタを結び付け、
'  差異区分・優先度・対応案を OUT_照合結果 に出力する
'==========================================================

Private Const DAYS_NO_CONTACT As Long = 60     ' 長期未通信とみなす日数
Private Const DAYS_LEASE_WARN As Long = 90     ' リース満了間近とみなす日数
Private Const RESULT_COLS As Long = 13

Public Sub MatchAndJudge(ByVal baseDate As Date)
    Dim wsHR As Worksheet, wsA As Worksheet, wsI As Worksheet, wsR As Worksheet
    Dim hr As Variant, ast As Variant, inv As Variant
    Dim cHR As New Collection, cInv As New Collection, cAst As New Collection, cCmt As New Collection
    Dim out() As Variant, n As Long, r As Long, k As Long, iH As Long, iI As Long
    Dim flags As String, acts As String, pri As String, cmt As Variant
    Dim stat As String, dtv As Date

    Set wsHR = ThisWorkbook.Worksheets(SH_HR)
    Set wsA = ThisWorkbook.Worksheets(SH_ASSET)
    Set wsI = ThisWorkbook.Worksheets(SH_INV)
    Set wsR = ResultSheet()

    hr = ReadBlock(wsHR, 7)
    ast = ReadBlock(wsA, 10)
    inv = ReadBlock(wsI, 6)
    BuildIndex cHR, hr, 1
    BuildIndex cInv, inv, 2
    BuildIndex cAst, ast, 5

    ' 前回の担当者コメントを引き継ぐ（キー：資産番号|シリアル）
    For r = FIRST_ROW To LastRow(wsR)
        If Len(CStr(wsR.Cells(r, 13).Value)) > 0 Then
            AddKey cCmt, CStr(wsR.Cells(r, 1).Value) & "|" & CStr(wsR.Cells(r, 2).Value), wsR.Cells(r, 13).Value
        End If
    Next r

    ReDim out(1 To RowsOf(ast) + RowsOf(inv), 1 To RESULT_COLS)

    ' --- 資産台帳を起点に判定 ---
    For r = 1 To RowsOf(ast)
        n = n + 1
        flags = "": acts = "": pri = ""
        out(n, 1) = ast(r, 1)
        out(n, 2) = ast(r, 5)
        out(n, 3) = ast(r, 6)
        out(n, 6) = ast(r, 7)

        If TryIndex(cHR, CStr(ast(r, 6)), iH) Then
            out(n, 4) = hr(iH, 2)
            out(n, 5) = hr(iH, 3)
            stat = CStr(hr(iH, 5))
            If stat = "退職" Then AddFlag flags, acts, "退職者割当", "端末回収・アカウント無効化確認"
            If stat = "休職" Then AddFlag flags, acts, "休職者割当", "所属長へ保管方針確認"
            If CStr(hr(iH, 4)) <> CStr(ast(r, 7)) Then AddFlag flags, acts, "拠点不一致(異動未反映)", "台帳の設置拠点を更新"
        End If

        If TryIndex(cInv, CStr(ast(r, 5)), iI) Then
            out(n, 7) = inv(iI, 1)
            If ToDate(inv(iI, 5), dtv) Then
                out(n, 8) = dtv
                If baseDate - dtv > DAYS_NO_CONTACT And CStr(ast(r, 10)) <> "予備" Then
                    AddFlag flags, acts, "長期未通信", "使用者へ利用状況ヒアリング"
                End If
            End If
            If Len(CStr(ast(r, 6))) > 0 And CStr(inv(iI, 3)) <> CStr(ast(r, 6)) Then
                AddFlag flags, acts, "使用者不一致", "又貸し・入替有無を確認"
            End If
        Else
            AddFlag flags, acts, "未検出", "現物確認(所在調査)"
        End If

        If ToDate(ast(r, 9), dtv) Then
            If dtv - baseDate <= DAYS_LEASE_WARN Then AddFlag flags, acts, "リース満了間近", "返却/再リース/入替を判断"
        End If

        If flags = "" Then
            out(n, 9) = "正常": out(n, 10) = "不要"
        Else
            out(n, 9) = flags: out(n, 10) = "要": out(n, 11) = acts
            If InStr(flags, "退職者割当") > 0 Or InStr(flags, "未検出") > 0 Then pri = "高" Else pri = "中"
        End If
        out(n, 12) = pri
        If TryValue(cCmt, CStr(out(n, 1)) & "|" & CStr(out(n, 2)), cmt) Then out(n, 13) = cmt
    Next r

    ' --- インベントリにあって台帳に無い機器 ---
    For r = 1 To RowsOf(inv)
        If Not TryIndex(cAst, CStr(inv(r, 2)), k) Then
            n = n + 1
            out(n, 1) = "(未登録)"
            out(n, 2) = inv(r, 2)
            out(n, 3) = inv(r, 3)
            If TryIndex(cHR, CStr(inv(r, 3)), iH) Then out(n, 4) = hr(iH, 2): out(n, 5) = hr(iH, 3)
            out(n, 7) = inv(r, 1)
            If ToDate(inv(r, 5), dtv) Then out(n, 8) = dtv
            out(n, 9) = "台帳未登録機器": out(n, 10) = "要"
            out(n, 11) = "持込・私物/検証機の可能性。登録または撤去"
            out(n, 12) = "高"
            If TryValue(cCmt, CStr(out(n, 1)) & "|" & CStr(out(n, 2)), cmt) Then out(n, 13) = cmt
        End If
    Next r

    WriteResult wsR, out, n
    WriteLog "MatchAndJudge", "完了", "出力 " & n & " 件（要対応 " & _
             Application.WorksheetFunction.CountIf(wsR.Range("J" & FIRST_ROW & ":J" & FIRST_ROW + n), "要") & " 件）"
End Sub

Private Sub WriteResult(ByVal ws As Worksheet, ByRef out() As Variant, ByVal n As Long)
    Dim r As Long, rng As Range
    ClearBody ws, RESULT_COLS
    If n = 0 Then Exit Sub
    Set rng = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(FIRST_ROW + n - 1, RESULT_COLS))
    rng.Value = out
    FormatBody rng
    ws.Range(ws.Cells(FIRST_ROW, 8), ws.Cells(FIRST_ROW + n - 1, 8)).NumberFormat = "yyyy/mm/dd"
    ws.Range(ws.Cells(FIRST_ROW, 13), ws.Cells(FIRST_ROW + n - 1, 13)).Interior.Color = RGB(255, 242, 204)
    For r = FIRST_ROW To FIRST_ROW + n - 1
        If ws.Cells(r, 10).Value = "要" Then ws.Range(ws.Cells(r, 9), ws.Cells(r, 12)).Interior.Color = RGB(252, 228, 214)
    Next r
End Sub

'---------------- 共通部品 ----------------
' 日付型・日付文字列・シリアル値のいずれでも日付として受け取る
Public Function ToDate(ByVal v As Variant, ByRef d As Date) As Boolean
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbDate Then
        d = v: ToDate = True
    ElseIf IsNumeric(v) Then
        If CDbl(v) > 0 Then d = CDate(CDbl(v)): ToDate = True
    ElseIf IsDate(v) Then
        d = CDate(v): ToDate = True
    End If
End Function
Public Function ReadBlock(ByVal ws As Worksheet, ByVal nCols As Long) As Variant
    Dim lr As Long, v As Variant
    lr = LastRow(ws)
    If lr < FIRST_ROW Then
        ReDim v(1 To 1, 1 To nCols)
        ReadBlock = v
        Exit Function
    End If
    v = ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lr, nCols)).Value
    ReadBlock = v
End Function

Public Function RowsOf(ByVal v As Variant) As Long
    If IsEmpty(v(1, 1)) And UBound(v, 1) = 1 Then RowsOf = 0 Else RowsOf = UBound(v, 1)
End Function

Public Sub BuildIndex(ByVal c As Collection, ByVal v As Variant, ByVal keyCol As Long)
    Dim r As Long
    For r = 1 To RowsOf(v)
        AddKey c, CStr(v(r, keyCol)), r
    Next r
End Sub

Public Sub AddKey(ByVal c As Collection, ByVal key As String, ByVal item As Variant)
    If Len(key) = 0 Then Exit Sub
    On Error Resume Next
    c.Add item, key
    On Error GoTo 0
End Sub

Public Function TryIndex(ByVal c As Collection, ByVal key As String, ByRef idx As Long) As Boolean
    Dim v As Variant
    If TryValue(c, key, v) Then idx = CLng(v): TryIndex = True
End Function

Public Function TryValue(ByVal c As Collection, ByVal key As String, ByRef v As Variant) As Boolean
    If Len(key) = 0 Then Exit Function
    On Error GoTo NF
    v = c.Item(key)
    TryValue = True
    Exit Function
NF:
    TryValue = False
End Function

Private Sub AddFlag(ByRef flags As String, ByRef acts As String, ByVal f As String, ByVal a As String)
    If flags = "" Then flags = f Else flags = flags & " / " & f
    If acts = "" Then acts = a Else acts = acts & "；" & a
End Sub

Public Sub ClearBody(ByVal ws As Worksheet, ByVal nCols As Long)
    Dim lr As Long
    lr = LastRow(ws)
    If lr < FIRST_ROW Then Exit Sub
    With ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(lr, nCols))
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
End Sub

Public Sub FormatBody(ByVal rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(191, 191, 191)
        .Font.Name = "Meiryo UI"
        .VerticalAlignment = xlTop
        .WrapText = True
    End With
End Sub
