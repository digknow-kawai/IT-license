Attribute VB_Name = "M05_Request"
Option Explicit
'==========================================================
' M05_Request : 部門別確認依頼リスト出力
'  OUT_照合結果 のうち『要』で、担当者コメントが『除外』で始まらない行を
'  所属順に並べて OUT_部門別確認依頼 へ出力する
'==========================================================

Private Const REQ_COLS As Long = 8
Private Const ANSWER_BIZDAYS As Long = 10      ' 回答期限（基準日からの営業日数）

Public Function ExportDeptRequests(ByVal baseDate As Date) As Long
    Dim wsR As Worksheet, wsQ As Worksheet, res As Variant
    Dim idx() As Long, n As Long, i As Long, j As Long, t As Long, r As Long
    Dim out() As Variant, due As Date, rng As Range

    Set wsR = ThisWorkbook.Worksheets(SH_RESULT)
    Set wsQ = RequestSheet()
    res = ReadBlock(wsR, 13)

    ReDim idx(1 To RowsOf(res) + 1)
    For r = 1 To RowsOf(res)
        If CStr(res(r, 10)) = "要" And Left$(Trim$(CStr(res(r, 13))), 2) <> "除外" Then
            n = n + 1: idx(n) = r
        End If
    Next r

    ' 所属 → 資産番号 の順に並べ替え（件数が少ないため単純挿入ソート）
    For i = 2 To n
        t = idx(i): j = i - 1
        Do While j >= 1
            If SortKey(res, idx(j)) <= SortKey(res, t) Then Exit Do
            idx(j + 1) = idx(j): j = j - 1
        Loop
        idx(j + 1) = t
    Next i

    due = AddBizDays(baseDate, ANSWER_BIZDAYS)
    ClearBody wsQ, REQ_COLS
    If n > 0 Then
        ReDim out(1 To n, 1 To REQ_COLS)
        For i = 1 To n
            r = idx(i)
            out(i, 1) = IIf(Len(CStr(res(r, 5))) = 0, "(所属不明)", res(r, 5))
            out(i, 2) = res(r, 1)
            out(i, 3) = res(r, 4)
            out(i, 4) = res(r, 9)
            out(i, 5) = res(r, 11)
            out(i, 6) = due
        Next i
        Set rng = wsQ.Range(wsQ.Cells(FIRST_ROW, 1), wsQ.Cells(FIRST_ROW + n - 1, REQ_COLS))
        rng.Value = out
        FormatBody rng
        wsQ.Range(wsQ.Cells(FIRST_ROW, 6), wsQ.Cells(FIRST_ROW + n - 1, 6)).NumberFormat = "yyyy/mm/dd"
        wsQ.Range(wsQ.Cells(FIRST_ROW, 7), wsQ.Cells(FIRST_ROW + n - 1, 8)).Interior.Color = RGB(255, 242, 204)
    End If

    WriteLog "ExportDeptRequests", "完了", "依頼 " & n & " 件／回答期限 " & Format(due, "yyyy/mm/dd")
    ExportDeptRequests = n
End Function

Private Function SortKey(ByVal res As Variant, ByVal r As Long) As String
    If Len(CStr(res(r, 5))) = 0 Then
        SortKey = "(所属不明)|" & CStr(res(r, 1))
    Else
        SortKey = CStr(res(r, 5)) & "|" & CStr(res(r, 1))
    End If
End Function

' 土日を除いて営業日を加算（祝日は考慮しない）
Private Function AddBizDays(ByVal d As Date, ByVal days As Long) As Date
    Dim c As Long
    Do While c < days
        d = d + 1
        If Weekday(d, vbMonday) <= 5 Then c = c + 1
    Loop
    AddBizDays = d
End Function
