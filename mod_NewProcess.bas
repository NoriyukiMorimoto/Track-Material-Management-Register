Attribute VB_Name = "mod_NewProcess"
Option Explicit

' 工事番号データを保持する共有変数
Public SharedMasterData As Variant

' マスタファイルのフルパス
Public Const MASTER_FILE_PATH As String = "C:\Users\n-morimoto\大鉄工業株式会社\本社現場サポート室 - 本社現場サポート室\人員リスト・組織図\出張所＆JR管理室.xlsx"

'===========================================================
' 指定された列のデータを7行目から最終行までクリアする
'===========================================================
Public Sub ClearTargetData(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim col As Variant
    
    ' D列（工事番号が入力される列など）を基準に最終行を判定
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If lastRow < 7 Then Exit Sub
    
    On Error Resume Next
    ' 指定された列（A, C, H, J, M）をクリア
    For Each col In Array("A", "C", "H", "J", "M")
        ws.Range(ws.Cells(7, col), ws.Cells(lastRow, col)).ClearContents
    Next col
    On Error GoTo 0
    
    ' D2（管理室）・D3（工事番号）・E3（工事件名）もクリア
    Application.EnableEvents = False
    ws.Range("D2").Value = ""
    ws.Range("D2").Validation.Delete
    ws.Range("D3").Value = ""
    ws.Range("E3").Value = ""
    Application.EnableEvents = True
    
    ' 工事番号キャッシュをリセット
    SharedMasterData = Empty
    
    ' シート名を「原図」に戻す
    On Error Resume Next
    ws.Name = "原図"
    If Err.Number <> 0 Then
        MsgBox "シート名を「原図」に戻せませんでした。" & vbCrLf & _
               "同名のシートが既に存在する可能性があります。", vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

'===========================================================
' 工事番号（D3）確定時のメイン処理（プレースホルダ）
'===========================================================
Public Sub RunDataProcessForD3(ByVal ws As Worksheet)
    ' ここにデータの読み込みや計算などのメインロジックを記述します
    MsgBox ws.Range("D3").Value & " のデータ処理を開始します。", vbInformation
End Sub

