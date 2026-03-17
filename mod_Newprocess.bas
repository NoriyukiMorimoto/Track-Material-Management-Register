Attribute VB_Name = "mod_Newprocess"
Option Explicit

' 工事番号データを保持する共有変数
Public SharedMasterData As Variant

' マスタファイルのフルパス
Public Const MASTER_FILE_PATH As String = "C:\Users\n-morimoto\大鉄工業株式会社\本社現場サポート室 - 本社現場サポート室\人員リスト・組織図\出張所＆JR管理室.xlsx"

' 単価表ベースフォルダ定数
Private Const UNIT_PRICE_FOLDER As String = "帳票7_支払金額計算シート - 帳票7_支払金額計算シート"
Private Const UNIT_PRICE_TABLE As String = "★購入充当材料単価表"
Private Const MASTER_SHEET_NAME As String = "単価適用線区"
Private Const HAYAKI_FOLDER As String = "早期発注"
Private Const SEKKEI_FOLDER As String = "設計変更"
Private Const ZAIRAISEN_FOLDER As String = "在来線"
Private Const SHINKANSEN_FOLDER As String = "新幹線"

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

    ' シート名を「原図」に戻す（既に「原図」なら何もしない）
    If ws.Name <> "原図" Then
        On Error Resume Next
        ws.Name = "原図"
        If Err.Number <> 0 Then
            MsgBox "シート名を「原図」に戻せませんでした。" & vbCrLf & _
                   "同名のシートが既に存在する可能性があります。", vbExclamation
            Err.Clear
        End If
        On Error GoTo 0
    End If

    ' 単価シート・設計変更単価シートをクリア
    Call ClearUnitPriceSheets(ws.Parent)
End Sub

'===========================================================
' 単価・設計変更単価シートの内容をクリア（シートは削除しない）
'===========================================================
Private Sub ClearUnitPriceSheets(ByVal wb As Workbook)
    Dim sheetNames(1) As String
    sheetNames(0) = "単価"
    sheetNames(1) = "設計変更単価"

    Dim i As Integer
    Dim ws As Worksheet
    For i = 0 To 1
        Set ws = Nothing
        On Error Resume Next
        Set ws = wb.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Application.EnableEvents = False
            ws.Cells.Clear
            Application.EnableEvents = True
            Set ws = Nothing
        End If
    Next i
End Sub


'===========================================================
' ボタン用：単価シートへの早期発注データ取込
'===========================================================
Public Sub ImportSoukiData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim officeName As String
    Dim nendo As Long
    officeName = Trim(ws.Range("E1").Value)
    If officeName = "" Then
        MsgBox "F1に出張所名を入力してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    nendo = CLng(ws.Range("B1").Value)
    On Error GoTo 0
    If nendo = 0 Then
        MsgBox "B1に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    ' 対象シートの存在確認
    Dim wsTanka As Worksheet
    If Not GetOrErrorSheet(ws.Parent, "単価", wsTanka) Then Exit Sub

    ' データ有無チェック
    Dim hasData As Boolean
    hasData = (wsTanka.Cells.Find("*") Is Nothing) = False
    Dim b1c1 As String
    b1c1 = Trim(CStr(ws.Range("B1").Value)) & Trim(CStr(ws.Range("C1").Value))
    If hasData Then
        Dim tankaHeader As String
        tankaHeader = GetSheetHeaderText(wsTanka.Parent, "単価")
        Dim resp As VbMsgBoxResult
        resp = MsgBox("既に「" & tankaHeader & "」のデータが入力されていますが、取込を実行してもいいですか？", _
                      vbOKCancel + vbExclamation, "確認")
        If resp = vbCancel Then Exit Sub
    Else
        MsgBox b1c1 & "_早期発注購入充当データを取り込みます。", vbInformation, "取込開始"
    End If

    ' マスタから線区情報を取得
    Dim shisha As String
    Dim tankaFolder As String
    Dim isShinkansen As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo CleanupSouki

    Dim hayakiPath As String
    hayakiPath = BuildFolderPath(nendo, HAYAKI_FOLDER, isShinkansen, shisha, tankaFolder)
    If hayakiPath = "" Then
        MsgBox "まだ、早期発注単価表は作成されていません。", vbExclamation
        GoTo CleanupSouki
    End If

    wsTanka.Cells.Clear
    Dim soukiOk As Boolean
    soukiOk = ImportUnitPriceFiles(hayakiPath, wsTanka, "TableStyleMedium2")
    If Not soukiOk Then
        MsgBox "まだ、早期発注単価表は作成されていません。", vbExclamation
        GoTo CleanupSouki
    End If

CleanupSouki:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    If soukiOk Then
        Dim msgS As String
        msgS = BuildCompleteMessage(ws.Parent, True, False)
        If msgS <> "" Then MsgBox msgS, vbInformation, "完了"
    End If
End Sub

'===========================================================
' ボタン用：設計変更単価シートへの設計変更データ取込
'===========================================================
Public Sub ImportSekkeiData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim officeName As String
    Dim nendo As Long
    officeName = Trim(ws.Range("E1").Value)
    If officeName = "" Then
        MsgBox "F1に出張所名を入力してください。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    nendo = CLng(ws.Range("B1").Value)
    On Error GoTo 0
    If nendo = 0 Then
        MsgBox "B1に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    ' 対象シートの存在確認
    Dim wsSekkei As Worksheet
    If Not GetOrErrorSheet(ws.Parent, "設計変更単価", wsSekkei) Then Exit Sub

    ' データ有無チェック
    Dim hasData As Boolean
    hasData = (wsSekkei.Cells.Find("*") Is Nothing) = False
    Dim b1c1 As String
    b1c1 = Trim(CStr(ws.Range("B1").Value)) & Trim(CStr(ws.Range("C1").Value))
    If hasData Then
        Dim sekkeiHeader As String
        sekkeiHeader = GetSheetHeaderText(wsSekkei.Parent, "設計変更単価")
        Dim resp As VbMsgBoxResult
        resp = MsgBox("既に「" & sekkeiHeader & "」のデータが入力されていますが、取込を実行してもいいですか？", _
                      vbOKCancel + vbExclamation, "確認")
        If resp = vbCancel Then Exit Sub
    Else
        MsgBox b1c1 & "_設計変更購入充当データを取り込みます。", vbInformation, "取込開始"
    End If

    ' マスタから線区情報を取得
    Dim shisha As String
    Dim tankaFolder As String
    Dim isShinkansen As Boolean

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo CleanupSekkei

    Dim sekkeiPath As String
    sekkeiPath = BuildFolderPath(nendo, SEKKEI_FOLDER, isShinkansen, shisha, tankaFolder)
    If sekkeiPath = "" Then
        MsgBox "まだ、設計変更単価表は作成されていません。", vbExclamation
        GoTo CleanupSekkei
    End If

    wsSekkei.Cells.Clear
    Dim sekkeiOk As Boolean
    sekkeiOk = ImportUnitPriceFiles(sekkeiPath, wsSekkei, "TableStyleMedium7")
    If Not sekkeiOk Then
        MsgBox "まだ、設計変更単価表は作成されていません。", vbExclamation
        GoTo CleanupSekkei
    End If

CleanupSekkei:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate

    If sekkeiOk Then
        Dim msgK As String
        msgK = BuildCompleteMessage(ws.Parent, False, True)
        If msgK <> "" Then MsgBox msgK, vbInformation, "完了"
    End If
End Sub

'===========================================================
' F1確定時のメイン処理（Sheet20から呼び出し）
'===========================================================
Public Sub RunF1Process(ByVal ws As Worksheet)
    Dim officeName As String
    Dim nendo As Long
    officeName = Trim(ws.Range("E1").Value)
    If officeName = "" Then Exit Sub

    On Error Resume Next
    nendo = CLng(ws.Range("B1").Value)
    On Error GoTo 0
    If nendo = 0 Then
        MsgBox "B1に年度が入力されていません。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' マスタから線区情報を取得
    Dim shisha As String        ' D列：支社
    Dim tankaFolder As String   ' F列：単価適用保線区（フォルダ名）
    Dim isShinkansen As Boolean
    If Not GetSenku(officeName, shisha, tankaFolder, isShinkansen) Then GoTo Cleanup

    ' 単価シート取込（早期発注）
    Dim hayakiPath As String
    hayakiPath = BuildFolderPath(nendo, HAYAKI_FOLDER, isShinkansen, shisha, tankaFolder)
    If hayakiPath <> "" Then
        Dim wsT As Worksheet
        If Not GetOrErrorSheet(ws.Parent, "単価", wsT) Then GoTo Cleanup
        ImportUnitPriceFiles hayakiPath, wsT, "TableStyleMedium2"
    End If

    ' 設計変更単価シート取込（設計変更）
    Dim sekkeiPath As String
    sekkeiPath = BuildFolderPath(nendo, SEKKEI_FOLDER, isShinkansen, shisha, tankaFolder)
    If sekkeiPath <> "" Then
        Dim wsS As Worksheet
        If Not GetOrErrorSheet(ws.Parent, "設計変更単価", wsS) Then GoTo Cleanup
        ImportUnitPriceFiles sekkeiPath, wsS, "TableStyleMedium7"
    End If

Cleanup:
    ' 画面更新を復元（再計算・EnableEvents復元はSheet20側で行う）
    Application.ScreenUpdating = True
End Sub

'===========================================================
' 完了メッセージ文字列を組み立てて返す
' 単価・設計変更単価シートのA1:G5の文字列を取得し整形
'===========================================================
Public Function BuildCompleteMessage(ByVal wb As Workbook, _
                                      Optional ByVal includeTanka As Boolean = True, _
                                      Optional ByVal includeSekkei As Boolean = True) As String
    Dim msgTanka As String
    Dim msgSekkei As String
    If includeTanka Then msgTanka = GetSheetHeaderText(wb, "単価")
    If includeSekkei Then msgSekkei = GetSheetHeaderText(wb, "設計変更単価")

    If msgTanka <> "" And msgSekkei <> "" Then
        BuildCompleteMessage = msgTanka & vbCrLf & _
                               msgSekkei & vbCrLf & _
                               "上記2つの購入充当データを適用しました。"
    ElseIf msgTanka <> "" Then
        BuildCompleteMessage = msgTanka & "の購入充当データを適用しました。"
    ElseIf msgSekkei <> "" Then
        BuildCompleteMessage = msgSekkei & "の購入充当データを適用しました。"
    End If
End Function

'===========================================================
' 指定シートのA1:G5の空でないセル値を連結して返す
'===========================================================
Private Function GetSheetHeaderText(ByVal wb As Workbook, _
                                     ByVal sheetName As String) As String
    GetSheetHeaderText = ""
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    Dim arr As Variant
    arr = ws.Range("A1:G1").Value

    Dim c As Long
    Dim parts() As String
    ReDim parts(0)
    Dim cnt As Long
    cnt = 0

    For c = 1 To 7
        Dim v As String
        v = Trim(CStr(arr(1, c)))
        If v <> "" Then
            ReDim Preserve parts(cnt)
            parts(cnt) = v
            cnt = cnt + 1
        End If
    Next c

    If cnt > 0 Then
        GetSheetHeaderText = Join(parts, " ")
    End If
End Function

'===========================================================
' マスタファイルから出張所名に対応する線区情報を取得
' 戻り値：True=成功 / False=失敗（エラーメッセージ表示済）
'===========================================================
Private Function GetSenku(ByVal officeName As String, _
                           ByRef shisha As String, _
                           ByRef tankaFolder As String, _
                           ByRef isShinkansen As Boolean) As Boolean
    GetSenku = False
    shisha = ""
    tankaFolder = ""
    isShinkansen = False

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(MASTER_FILE_PATH) Then
        MsgBox "マスタファイルが見つかりませんでした。" & vbCrLf & MASTER_FILE_PATH, vbCritical
        Set fso = Nothing
        Exit Function
    End If
    Set fso = Nothing

    Dim masterWb As Workbook
    Dim masterWs As Worksheet
    On Error Resume Next
    Set masterWb = Workbooks.Open(MASTER_FILE_PATH, ReadOnly:=True)
    On Error GoTo 0
    If masterWb Is Nothing Then
        MsgBox "マスタファイルを開けませんでした。", vbCritical
        Exit Function
    End If

    On Error Resume Next
    Set masterWs = masterWb.Sheets(MASTER_SHEET_NAME)
    On Error GoTo 0
    If masterWs Is Nothing Then
        MsgBox "マスタファイルに「" & MASTER_SHEET_NAME & "」シートが見つかりません。", vbCritical
        masterWb.Close False
        Exit Function
    End If

    ' C列を検索して一致行のD列・F列を取得
    Dim lastRow As Long
    lastRow = masterWs.Cells(masterWs.Rows.Count, "C").End(xlUp).Row
    Dim i As Long
    Dim found As Boolean
    found = False
    For i = 2 To lastRow
        If Trim(CStr(masterWs.Cells(i, "C").Value)) = officeName Then
            shisha = Trim(CStr(masterWs.Cells(i, "D").Value))
            tankaFolder = Trim(CStr(masterWs.Cells(i, "F").Value))
            found = True
            Exit For
        End If
    Next i

    masterWb.Close False

    If Not found Then
        MsgBox "出張所「" & officeName & "」がマスタに見つかりませんでした。" & vbCrLf & _
               "シート：" & MASTER_SHEET_NAME & " / C列を確認してください。", vbExclamation
        Exit Function
    End If
    If tankaFolder = "" Then
        MsgBox "出張所「" & officeName & "」の単価適用保線区（F列）が空です。", vbExclamation
        Exit Function
    End If

    ' 新幹線判定：出張所名（C列）に「新幹線」が含まれる場合
    isShinkansen = (InStr(officeName, "新幹線") > 0)

    GetSenku = True
End Function

'===========================================================
' フォルダパスを組み立てて返す
' 存在しない場合は "" を返す（設計変更フォルダなし＝正常）
'===========================================================
Private Function BuildFolderPath(ByVal nendo As Long, _
                                  ByVal category As String, _
                                  ByVal isShinkansen As Boolean, _
                                  ByVal shisha As String, _
                                  ByVal tankaFolder As String) As String
    BuildFolderPath = ""
    Dim userName As String
    userName = Environ("USERNAME")
    Dim basePath As String
    basePath = "C:\Users\" & userName & "\大鉄工業株式会社\" & _
               UNIT_PRICE_FOLDER & "\" & UNIT_PRICE_TABLE & "\" & _
               CStr(nendo) & "\" & category & "\"

    Dim fullPath As String
    If isShinkansen Then
        fullPath = basePath & SHINKANSEN_FOLDER & "\" & tankaFolder
    Else
        fullPath = basePath & ZAIRAISEN_FOLDER & "\" & shisha & "\" & tankaFolder
    End If

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(fullPath) Then
        BuildFolderPath = fullPath & "\"
    End If
    Set fso = Nothing
End Function

'===========================================================
' 対象シートを取得する（存在しない場合はエラーメッセージ）
' 戻り値：True=取得成功 / False=シートなし
'===========================================================
Private Function GetOrErrorSheet(ByVal wb As Workbook, _
                                  ByVal sheetName As String, _
                                  ByRef ws As Worksheet) As Boolean
    GetOrErrorSheet = False
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Sheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "「" & sheetName & "」シートが見つかりません。" & vbCrLf & _
               "シートを作成してから再実行してください。", vbCritical
        Exit Function
    End If
    GetOrErrorSheet = True
End Function

'===========================================================
' 指定フォルダのExcelファイルを取込み、対象シートに貼り付け
' ・通常ファイル → テーブル解除→値・書式コピー→列幅・行高コピー
' ・"-2"付きファイル → 6行目以降を最終行の下に追記
' ・全データ転記後に1つのテーブルとして再設定
' ・印刷範囲クリア・表示を標準ビューに変更
'===========================================================
Private Function ImportUnitPriceFiles(ByVal folderPath As String, _
                                       ByVal wsDest As Worksheet, _
                                       Optional ByVal tableStyleName As String = "TableStyleMedium2") As Boolean
    ImportUnitPriceFiles = False  ' ファイルなし or 失敗
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        Set fso = Nothing
        Exit Function
    End If

    ' ファイルを通常/-2に仕分け
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    Dim normalFile As String
    Dim sub2File As String
    normalFile = ""
    sub2File = ""
    Dim f As Object
    Dim ext As String
    For Each f In folder.Files
        ext = LCase(fso.GetExtensionName(f.Name))
        If ext = "xls" Or ext = "xlsx" Or ext = "xlsm" Then
            If InStr(f.Name, "-2") > 0 Then
                sub2File = f.Path
            Else
                normalFile = f.Path
            End If
        End If
    Next f
    Set fso = Nothing

    If normalFile = "" Then Exit Function  ' ファイルなし→Falseのまま返る

    ' テーブルスタイル・ヘッダー行番号（固定値：別途ファイルを開かず高速化）
    Dim tableStyle As String
    Dim tableHeaderRow As Long
    tableStyle = tableStyleName       ' 呼び出し元からスタイルを受け取る
    tableHeaderRow = 5                ' ヘッダー行は5行目固定

    ' 送り先シートをクリア
    wsDest.Cells.Clear

    ' 通常ファイルを転記（A1から）
    Dim normalLastRow As Long
    normalLastRow = 0
    Call PasteFromFileFast(normalFile, wsDest, 1, normalLastRow)

    ' "-2"ファイルを追記（6行目以降を最終行の下に）
    Dim totalLastRow As Long
    totalLastRow = normalLastRow
    If sub2File <> "" And normalLastRow > 0 Then
        Call PasteFromFileFast(sub2File, wsDest, normalLastRow + 1, totalLastRow)
    End If

    ' テーブルとして再設定（ヘッダー行?全体最終行）
    If tableStyle <> "" And totalLastRow >= tableHeaderRow + 1 Then
        Call ResetAsTable(wsDest, tableHeaderRow, totalLastRow, tableStyle)
    End If

    ' 印刷範囲をクリア
    wsDest.PageSetup.PrintArea = ""

    ' 表示を標準ビューに変更
    wsDest.Parent.Windows(1).View = xlNormalView

    ' --- 書式設定 ---
    ' 全体（A1:G最終行）のフォントを「BIZ UDゴシック」サイズ11に設定
    With wsDest.Range(wsDest.Cells(1, 1), wsDest.Cells(totalLastRow, 7))
        .Font.Name = "BIZ UDゴシック"
        .Font.Size = 11
    End With

    ' A1:G1を結合してフォントサイズ14・中央揃えに設定
    With wsDest.Range("A1:G1")
        .Merge
        .Font.Name = "BIZ UDゴシック"
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' A5:G5（ヘッダー行）を水平・垂直中央揃えに設定
    With wsDest.Range("A5:G5")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' A列・E列・G列（6行目以降）を水平・垂直中央揃えに設定
    Dim centerCols As Variant
    Dim cc As Variant
    For Each cc In Array(1, 5, 7)  ' A=1, E=5, G=7
        With wsDest.Range(wsDest.Cells(6, cc), wsDest.Cells(totalLastRow, cc))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    Next cc

    ImportUnitPriceFiles = True  ' 正常完了
End Function

'===========================================================
' 取込元ファイルからテーブル情報（スタイル・ヘッダー行）を取得
'===========================================================
Private Sub GetTableInfo(ByVal filePath As String, _
                          ByRef tableStyle As String, _
                          ByRef tableHeaderRow As Long)
    tableStyle = ""
    tableHeaderRow = 1

    Dim srcWb As Workbook
    On Error Resume Next
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If srcWb Is Nothing Then Exit Sub

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = srcWb.Sheets("単価")
    On Error GoTo 0
    If Not srcWs Is Nothing Then
        If srcWs.ListObjects.Count > 0 Then
            tableStyle = srcWs.ListObjects(1).tableStyle
            tableHeaderRow = srcWs.ListObjects(1).HeaderRowRange.Row
        End If
    End If

    srcWb.Close False
End Sub

'===========================================================
' ファイルを開いてA1:G2500のスナップショットを取得し高速転記
' ・A1:G2500を配列で一括取得→即閉じる→貼り付け
' ・有効データの最終行までを対象とする
' ・destLastRowに転記後の最終行を返す
'===========================================================
Private Sub PasteFromFileFast(ByVal filePath As String, _
                               ByVal wsDest As Worksheet, _
                               ByVal destStartRow As Long, _
                               ByRef destLastRow As Long)
    destLastRow = 0

    Const SNAP_COLS As Long = 7     ' 列数固定（A?G列）

    ' ファイルを開いてスナップショット取得後すぐに閉じる
    Dim srcWb As Workbook
    On Error Resume Next
    Set srcWb = Workbooks.Open(filePath, ReadOnly:=True)
    On Error GoTo 0
    If srcWb Is Nothing Then
        MsgBox "ファイルを開けませんでした。" & vbCrLf & filePath, vbCritical
        Exit Sub
    End If

    Dim srcWs As Worksheet
    On Error Resume Next
    Set srcWs = srcWb.Sheets("単価")
    On Error GoTo 0
    If srcWs Is Nothing Then
        MsgBox "取込ファイルに「単価」シートが見つかりませんでした。" & vbCrLf & filePath, vbExclamation
        srcWb.Close False
        Exit Sub
    End If

    ' A列の最終行を取得してから配列で一括取得
    Dim snapRows As Long
    snapRows = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    If snapRows < 1 Then
        srcWb.Close False
        Exit Sub
    End If

    Dim dataArr As Variant
    dataArr = srcWs.Range("A1:G" & snapRows).Value

    ' 列幅をスナップ（通常ファイル時のみ）
    Dim colWidths(1 To SNAP_COLS) As Double
    If destStartRow = 1 Then
        Dim c As Long
        For c = 1 To SNAP_COLS
            colWidths(c) = srcWs.Columns(c).ColumnWidth
        Next c
    End If

    ' ファイルを即閉じる
    srcWb.Close False
    Set srcWs = Nothing
    Set srcWb = Nothing

    ' "-2"ファイルは6行目から取込（1?5行目除外）
    Dim srcStartRow As Long
    Dim srcLastRow As Long
    Dim r As Long

    If destStartRow > 1 Then
        srcStartRow = 6
    Else
        srcStartRow = 1
    End If

    srcLastRow = snapRows

    If srcLastRow = 0 Or srcLastRow < srcStartRow Then Exit Sub

    Dim rowCount As Long
    rowCount = srcLastRow - srcStartRow + 1

    ' 列幅を設定（通常ファイル時のみ）
    If destStartRow = 1 Then
        For c = 1 To SNAP_COLS
            wsDest.Columns(c).ColumnWidth = colWidths(c)
        Next c
    End If

    ' 転記用配列を切り出し
    Dim writeArr() As Variant
    ReDim writeArr(1 To rowCount, 1 To SNAP_COLS)
    Dim col As Long
    For r = 1 To rowCount
        For col = 1 To SNAP_COLS
            writeArr(r, col) = dataArr(srcStartRow + r - 1, col)
        Next col
    Next r

    ' 一括書き込み
    wsDest.Range(wsDest.Cells(destStartRow, 1), _
                 wsDest.Cells(destStartRow + rowCount - 1, SNAP_COLS)).Value = writeArr

    ' A列：数値書式
    wsDest.Range(wsDest.Cells(destStartRow, 1), _
                 wsDest.Cells(destStartRow + rowCount - 1, 1)).NumberFormat = "0"

    ' F列：桁区切り数値書式
    wsDest.Range(wsDest.Cells(destStartRow, 6), _
                 wsDest.Cells(destStartRow + rowCount - 1, 6)).NumberFormat = "#,##0"

    destLastRow = destStartRow + rowCount - 1
End Sub

'===========================================================
' 指定範囲をテーブルとして再設定
'===========================================================
Private Sub ResetAsTable(ByVal ws As Worksheet, _
                          ByVal headerRow As Long, _
                          ByVal lastRow As Long, _
                          ByVal tableStyle As String)
    ' 既存テーブルがあれば解除
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        lo.Unlist
    Next lo

    ' ヘッダー行から最終行・最終列までをテーブルとして設定
    Dim lastCol As Long
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    Dim tblRange As Range
    Set tblRange = ws.Range(ws.Cells(headerRow, 1), ws.Cells(lastRow, lastCol))

    Dim newTable As ListObject
    Set newTable = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    On Error Resume Next
    newTable.tableStyle = tableStyle
    On Error GoTo 0
End Sub

'===========================================================
' 工事番号（D3）確定時のメイン処理（プレースホルダ）
'===========================================================
Public Sub RunDataProcessForD3(ByVal ws As Worksheet)
    ' ここにデータの読み込みや計算などのメインロジックを記述します
    MsgBox ws.Range("D3").Value & " のデータ処理を開始します。", vbInformation
End Sub


