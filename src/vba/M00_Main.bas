Attribute VB_Name = "M00_Main"
Option Explicit
'==========================================================
' M00_Main : 全体の実行制御
'  IT資産・ライセンス 定期棚卸 照合ツール
'  手順1 : Main_RunInventory     （インプットExcel取込 → 突合・差異判定 → ライセンス集計 → サマリ）
'  手順2 : Step2_ExportRequests  （担当者レビュー後に、確認依頼の作成とログ保存）
'  ※ OUT_ シートと LOG_ シートは、無ければ自動生成する
'==========================================================

Public Const SH_COVER As String = "00_表紙"
Public Const SH_HR As String = "IN_人事マスタ"
Public Const SH_ASSET As String = "IN_資産台帳"
Public Const SH_INV As String = "IN_インベントリ"
Public Const SH_LIC As String = "IN_ライセンス契約"
Public Const SH_USE As String = "IN_ソフト利用実績"
Public Const SH_RESULT As String = "OUT_照合結果"
Public Const SH_LICOUT As String = "OUT_ライセンス過不足"
Public Const SH_REQ As String = "OUT_部門別確認依頼"
Public Const SH_SUMMARY As String = "OUT_サマリ"
Public Const SH_LOG As String = "LOG_実行ログ"

Public Const HEADER_ROW As Long = 4
Public Const FIRST_ROW As Long = 5
Public Const BASE_DATE_CELL As String = "B7"

' 手順1：ボタン「① 照合を実行」から呼び出す
Public Sub Main_RunInventory()
    RunStep1 False
End Sub

' 手順2：ボタン「② 確認依頼を作成」から呼び出す
Public Sub Step2_ExportRequests()
    RunStep2 False
End Sub

Public Sub RunStep1(ByVal silent As Boolean, Optional ByVal inputPath As String = "")
    Dim baseDate As Date
    On Error GoTo EH
    baseDate = GetBaseDate()

    Application.ScreenUpdating = False
    WriteLog "Main_RunInventory", "開始", "基準日 " & Format(baseDate, "yyyy/mm/dd")

    If Not ImportAndNormalize(silent, inputPath) Then   ' M01 取込・整形
        Application.ScreenUpdating = True
        Exit Sub
    End If
    MatchAndJudge baseDate          ' M02 突合・差異判定
    AggregateLicenses baseDate      ' M04 ライセンス集計
    BuildSummary baseDate           ' M04 サマリ

    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_RESULT).Activate
    On Error GoTo 0
    If Not silent Then
        MsgBox "照合が完了しました。" & vbCrLf & _
               "OUT_照合結果 の『要』行を確認し、除外する行は担当者コメントに『除外』と記入してから" & vbCrLf & _
               "② 確認依頼を作成 を実行してください。", vbInformation, "棚卸 照合ツール"
    End If
    Exit Sub
EH:
    Application.ScreenUpdating = True
    WriteLog "Main_RunInventory", "エラー", Err.Description
    If Not silent Then MsgBox "処理を中断しました：" & Err.Description, vbCritical, "棚卸 照合ツール"
End Sub

Public Sub RunStep2(ByVal silent As Boolean)
    Dim n As Long
    On Error GoTo EH
    If Not SheetExists(SH_RESULT) Then
        Err.Raise vbObjectError + 3, , "OUT_照合結果 がありません。先に ① 照合を実行 してください。"
    End If
    Application.ScreenUpdating = False
    n = ExportDeptRequests(GetBaseDate())   ' M05 部門別確認依頼
    WriteLogAndArchive silent               ' M06 ログ保存
    Application.ScreenUpdating = True
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_REQ).Activate
    On Error GoTo 0
    If Not silent Then MsgBox "確認依頼を " & n & " 件作成しました。", vbInformation, "棚卸 照合ツール"
    Exit Sub
EH:
    Application.ScreenUpdating = True
    WriteLog "Step2_ExportRequests", "エラー", Err.Description
    If Not silent Then MsgBox "処理を中断しました：" & Err.Description, vbCritical, "棚卸 照合ツール"
End Sub

Public Function GetBaseDate() As Date
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_COVER).Range(BASE_DATE_CELL).Value
    If Not IsDate(v) Then Err.Raise vbObjectError + 1, , "表紙の棚卸基準日（" & BASE_DATE_CELL & "）が日付ではありません。"
    GetBaseDate = CDate(v)
End Function

' 対象シートの最終データ行（A列が空になる直前の行）
Public Function LastRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    r = HEADER_ROW
    Do While Len(CStr(ws.Cells(r + 1, 1).Value)) > 0
        r = r + 1
    Loop
    LastRow = r
End Function

' 表紙に実行ボタンを配置（ブックを開いたときに無ければ作成）
Public Sub EnsureButtons()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_COVER)
    If ws Is Nothing Then Exit Sub
    AddButton ws, "btnStep1", "① 照合を実行", "Main_RunInventory", ws.Range("E7")
    AddButton ws, "btnStep2", "② 確認依頼を作成", "Step2_ExportRequests", ws.Range("E9")
End Sub

Private Sub AddButton(ByVal ws As Worksheet, ByVal nm As String, ByVal caption As String, _
                      ByVal macroName As String, ByVal pos As Range)
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.Shapes(nm)
    On Error GoTo 0
    If Not shp Is Nothing Then Exit Sub
    Set shp = ws.Shapes.AddShape(5, pos.Left, pos.Top, 170, 30)   ' 5 = 角丸四角形
    shp.Name = nm
    shp.Fill.ForeColor.RGB = RGB(31, 78, 121)
    shp.Line.Visible = False
    shp.TextFrame.Characters.Text = caption
    shp.TextFrame.Characters.Font.Color = RGB(255, 255, 255)
    shp.TextFrame.Characters.Font.Bold = True
    shp.TextFrame.HorizontalAlignment = xlHAlignCenter
    shp.TextFrame.VerticalAlignment = xlVAlignCenter
    shp.OnAction = macroName
End Sub
