Attribute VB_Name = "M04_License"
Option Explicit
'==========================================================
' M04_License : ライセンス集計・サマリ作成
'  OUT_ライセンス過不足 と OUT_サマリ は数式で作成する
'  （IN_ シートや担当者コメントを直すと自動で再計算される）
'==========================================================

Private Const SURPLUS_RATE As Double = 0.2     ' 余剰とみなす未使用率

Public Sub AggregateLicenses(ByVal baseDate As Date)
    Dim wsL As Worksheet, wsIn As Worksheet, n As Long, i As Long, r As Long, s As Long, shortCnt As Long
    Set wsIn = ThisWorkbook.Worksheets(SH_LIC)
    Set wsL = LicenseSheet()
    ClearBody wsL, 8
    n = DataCount(SH_LIC)
    For i = 1 To n
        r = FIRST_ROW + i - 1          ' 出力行
        s = FIRST_ROW + i - 1          ' IN_ライセンス契約 の行
        wsL.Cells(r, 1).Formula = "='" & SH_LIC & "'!A" & s
        wsL.Cells(r, 2).Formula = "='" & SH_LIC & "'!C" & s
        wsL.Cells(r, 3).Formula = "=COUNTIF('" & SH_USE & "'!$B:$B,A" & r & ")"
        wsL.Cells(r, 4).Formula = "=B" & r & "-C" & r
        wsL.Cells(r, 5).Formula = "=IF(D" & r & "<0,""不足"",IF(D" & r & ">=B" & r & "*" & SURPLUS_RATE & ",""余剰"",""適正""))"
        wsL.Cells(r, 6).Formula = "=IF('" & SH_LIC & "'!D" & s & "="""",""-"",'" & SH_LIC & "'!D" & s & ")"
        wsL.Cells(r, 7).Formula = "=IF(F" & r & "=""-"",""-"",F" & r & "-'" & SH_COVER & "'!$B$7)"
        wsL.Cells(r, 8).Formula = "=IF(D" & r & "<0,-D" & r & "*'" & SH_LIC & "'!E" & s & ",0)"
    Next i
    If n > 0 Then
        FormatBody wsL.Range(wsL.Cells(FIRST_ROW, 1), wsL.Cells(FIRST_ROW + n - 1, 8))
        wsL.Range(wsL.Cells(FIRST_ROW, 6), wsL.Cells(FIRST_ROW + n - 1, 6)).NumberFormat = "yyyy/mm/dd"
        wsL.Range(wsL.Cells(FIRST_ROW, 7), wsL.Cells(FIRST_ROW + n - 1, 7)).NumberFormat = "0"
        wsL.Range(wsL.Cells(FIRST_ROW, 8), wsL.Cells(FIRST_ROW + n - 1, 8)).NumberFormat = "#,##0;-#,##0;-"
    End If
    Application.Calculate
    For r = FIRST_ROW To FIRST_ROW + n - 1
        If wsL.Cells(r, 5).Value = "不足" Then shortCnt = shortCnt + 1
    Next r
    WriteLog "AggregateLicenses", "完了", n & " ソフト（不足 " & shortCnt & " 件）"
End Sub

Public Sub BuildSummary(ByVal baseDate As Date)
    Dim ws As Worksheet, cats As Variant, i As Long, r As Long
    Dim rngI As String, rngJ As String, rngL As String
    Set ws = SummarySheet()
    ClearBody ws, 3
    ws.Range("A2").Value = "基準日：" & Format(baseDate, "yyyy/mm/dd") & "　（数値は OUT_照合結果・OUT_ライセンス過不足 から数式で集計）"

    rngI = "'" & SH_RESULT & "'!$I$" & FIRST_ROW & ":$I$5000"
    rngJ = "'" & SH_RESULT & "'!$J$" & FIRST_ROW & ":$J$5000"
    rngL = "'" & SH_RESULT & "'!$L$" & FIRST_ROW & ":$L$5000"
    r = FIRST_ROW
    PutRow ws, r, "照合対象（総件数）", "=COUNTA('" & SH_RESULT & "'!$A$" & FIRST_ROW & ":$A$5000)", "", True
    PutRow ws, r, "うち 正常", "=COUNTIF(" & rngI & ",""正常"")", "", True
    PutRow ws, r, "うち 対応要", "=COUNTIF(" & rngJ & ",""要"")", "", True
    PutRow ws, r, "対応要の比率", "=IF(B" & FIRST_ROW & "=0,0,B" & FIRST_ROW + 2 & "/B" & FIRST_ROW & ")", "", True
    ws.Cells(r - 1, 2).NumberFormat = "0.0%"
    cats = Array("退職者割当", "休職者割当", "拠点不一致", "未検出", "長期未通信", "使用者不一致", "リース満了間近", "台帳未登録機器")
    For i = 0 To UBound(cats)
        PutRow ws, r, "　" & cats(i), "=COUNTIF(" & rngI & ",""*" & cats(i) & "*"")", "1台に複数区分が付く場合あり", False
    Next i
    PutRow ws, r, "優先度『高』", "=COUNTIF(" & rngL & ",""高"")", "", True
    PutRow ws, r, "ライセンス不足ソフト数", "=COUNTIF('" & SH_LICOUT & "'!$E$" & FIRST_ROW & ":$E$500,""不足"")", "", True
    PutRow ws, r, "ライセンス追加費用目安(円)", "=SUM('" & SH_LICOUT & "'!$H$" & FIRST_ROW & ":$H$500)", "", True
    ws.Cells(r - 1, 2).NumberFormat = "#,##0"
    FormatBody ws.Range(ws.Cells(FIRST_ROW, 1), ws.Cells(r - 1, 3))
    Application.Calculate
    WriteLog "BuildSummary", "完了", "要対応 " & ws.Cells(FIRST_ROW + 2, 2).Value & " 件"
End Sub

Private Sub PutRow(ByVal ws As Worksheet, ByRef r As Long, ByVal label As String, ByVal fml As String, _
                   ByVal note As String, ByVal isBold As Boolean)
    ws.Cells(r, 1).Value = label
    ws.Cells(r, 2).Formula = fml
    ws.Cells(r, 3).Value = note
    ws.Cells(r, 1).Font.Bold = isBold
    r = r + 1
End Sub
