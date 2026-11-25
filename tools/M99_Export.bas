Attribute VB_Name = "M99_Export"
Option Explicit
'==========================================================
' M99_Export : VBAモジュールを src\vba へ一括エクスポート（Git管理用）
'  事前に［トラストセンター］→［マクロの設定］→
'  「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」をオンにする
'  ブックはリポジトリ直下の tool フォルダに置く想定（..\src\vba へ出力）
'==========================================================
Public Sub ExportModules()
    Dim c As Object, ext As String, root As String, dirPath As String, n As Long
    root = ThisWorkbook.Path & "\.."
    If Dir(root & "\src", vbDirectory) = "" Then root = ThisWorkbook.Path
    dirPath = root & "\src\vba\"
    If Dir(root & "\src", vbDirectory) = "" Then MkDir root & "\src"
    If Dir(dirPath, vbDirectory) = "" Then MkDir dirPath
    For Each c In ThisWorkbook.VBProject.VBComponents
        ext = ""
        Select Case c.Type
            Case 1: ext = ".bas"                    ' 標準モジュール
            Case 2: ext = ".cls"                    ' クラスモジュール
            Case 3: ext = ".frm"                    ' ユーザーフォーム
            Case 100                                ' ThisWorkbook・シート（コードがあるものだけ）
                If c.CodeModule.CountOfLines > 1 Then ext = ".cls"
        End Select
        If ext <> "" And c.Name <> "M99_Export" Then
            c.Export dirPath & c.Name & ext
            n = n + 1
        End If
    Next c
    MsgBox n & " 個のモジュールをエクスポートしました。" & vbCrLf & dirPath, vbInformation
End Sub
