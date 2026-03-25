Attribute VB_Name = "Module1"
Option Explicit

'残数量を前期繰越へ
'
'今期注文数量、2/4で使用をクリア
'2024/12/05 O.Kanai
Sub ボタン1_Click()
    Dim i As Long
    Dim iMax As Long
    Dim wkNum As Long
    
    If MsgBox("残数量を前期繰越へ設定しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    i = 7 '明細開始の行
    Do
        If i > 10000 Then Exit Do  '周り過ぎないように
        
        '品目コードが入力ある分のみ処理
        If Range("A" & i) = "" Then
            Exit Do
        Else
            '残数量を取得し、前期繰越へ設定
            wkNum = Val(Range("S" & i).Value)  '残数量(S列)
            Range("K" & i).Value = wkNum       '前期繰越(K列)
            
            '今期注文数量、2/4で使用をクリア
            Range("M" & i).ClearContents       '今期注文(M列)
            Range("Q" & i).ClearContents       '2/4で使用(Q列)
        End If
    
        i = i + 1
    Loop

    iMax = i - 1 '最大値を保存
    i = iMax
    Do
        If i < 7 Then
            Exit Do
        Else
            '残数量が０の場合に行を削除
            If Val(Range("S" & i).Value) = 0 Then '残数量(S列)
                Rows(i).Delete
            End If
        End If
        i = i - 1
    Loop
End Sub

'支払金額計算取り込み
'
'
'2025/01/22 O.Kanai
Sub ボタン2_Click()
    '支払金額計算取り込み
    '単なる入力チェック＋フォーム呼び出し（本処理は UserForm1 側）
    If Range("D3").Text = "" Then
        MsgBox "工番・件名を入力してください", vbCritical
        Exit Sub
    End If
    UserForm1.Show
End Sub

'最終施行指示取り込み
'
'
'2025/01/22 O.Kanai
Sub ボタン3_Click()
    Dim i As Long, k As Long
    Dim wkNum As Double
    Dim keyCellDest As String
    Dim srcWorkbook As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim destRow As Long
    Dim conditionValue As String
    Dim startRowSource As Long
    Dim startRowDest As Long
    Dim serchColSource As Long
    Dim serchColDest As Long
    Dim keyColSource As Long
    Dim keyColSource2 As Long
    Dim keyColSource3 As Long
    Dim outputCount As Long
    Dim outputCount2 As Long
    Dim keyData As String '品目コード
    Dim keyData2 As String '管理室
    Dim keyKanri As String
    Dim insertFlag As Boolean
    Dim wkSekouCnt As Double

    keyCellDest = "D2"
    
    '入力チェック
    If Range(keyCellDest).Text = "" Then
        MsgBox "管理室を入力してください", vbCritical
        Exit Sub
    End If
    
    If MsgBox("最終施行指示を取り込みしますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' 1. ファイル選択
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="最終施行指示を取得するExcelファイルを選択してください")
    If filePath = "False" Then
        MsgBox "操作がキャンセルされました。", vbExclamation
        GoTo Cleanup
    End If

    ' 2. ファイルを開く
    Set srcWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    sheetName = srcWorkbook.ActiveSheet.Name
    
    ' 3. 各種設定
    serchColSource = 1   'A列：整理番号
    serchColDest = 1     'A列：品目コード
    keyColSource = 77    'BY列：工種分類
    keyColSource2 = 1    'A列：整理番号
    keyColSource3 = 69   'BQ列：管理室
    startRowSource = 26
    startRowDest = 7
    
    Set wsSource = srcWorkbook.Sheets(sheetName)
    Set wsDest = ThisWorkbook.ActiveSheet
    
    ' 最終行取得（xlUpで安定化）
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, serchColSource).End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, serchColDest).End(xlUp).Row
    If lastRowDest < startRowDest Then lastRowDest = startRowDest - 1
    
    conditionValue = "" '未使用だが残す
    destRow = lastRowDest
    If startRowSource >= lastRowSource Then GoTo NoData

    outputCount = 0   '更新件数
    outputCount2 = 0  '追加件数
    
    keyKanri = wsDest.Range(keyCellDest).Value
    
    ' データ取得
    For i = startRowSource To lastRowSource
        keyData = ""
        insertFlag = False
        If wsSource.Cells(i, keyColSource).Value = "購入充当" And _
           InStr(keyKanri, wsSource.Cells(i, keyColSource3).Value) > 0 Then   '購入充当 & 管理室一致
           
            insertFlag = True
            keyData = wsSource.Cells(i, keyColSource2).Value
            keyData2 = wsSource.Cells(i, keyColSource3).Value
            If keyData <> "" Then
                wkSekouCnt = Val(wsSource.Cells(i, 43).Value)  ' 施工数量（列43）
                
                For k = startRowDest To lastRowDest  ' 既存データを検索
                    If keyData = wsDest.Cells(k, 1).Value And _
                       keyData2 = wsDest.Cells(k, 8).Value Then 'キー一致（品目コード・管理室）
                       
                        If wsDest.Cells(k, 17).Value = "" Or _
                           (wsDest.Cells(k, 17).Value > "" And _
                           (wsDest.Cells(k, 17).Value <> wsDest.Cells(k, 11).Value And _
                            wsDest.Cells(k, 17).Value <> wsDest.Cells(k, 13).Value)) Then
                            
                            '注文数量 優先
                            If wsDest.Cells(k, 11).Value <> "" Then
                                If wkSekouCnt <= (Val(wsDest.Cells(k, 11).Value) - Val(wsDest.Cells(k, 17).Value)) Then
                                    wsDest.Cells(k, 17).Value = Val(wsDest.Cells(k, 17).Value) + wkSekouCnt
                                    wkSekouCnt = 0
                                    outputCount = outputCount + 1
                                    insertFlag = False
                                    Exit For
                                Else
                                    wkSekouCnt = wkSekouCnt - Val(wsDest.Cells(k, 11).Value) + Val(wsDest.Cells(k, 17).Value)
                                    wsDest.Cells(k, 17).Value = wsDest.Cells(k, 11).Value
                                    outputCount = outputCount + 1
                                End If
                            ElseIf wsDest.Cells(k, 13).Value <> "" Then
                                '納品数量 次優先
                                If wkSekouCnt <= (Val(wsDest.Cells(k, 13).Value) - Val(wsDest.Cells(k, 17).Value)) Then
                                    wsDest.Cells(k, 17).Value = Val(wsDest.Cells(k, 17).Value) + wkSekouCnt
                                    wkSekouCnt = 0
                                    outputCount = outputCount + 1
                                    insertFlag = False
                                    Exit For
                                Else
                                    wkSekouCnt = wkSekouCnt - Val(wsDest.Cells(k, 13).Value) + Val(wsDest.Cells(k, 17).Value)
                                    wsDest.Cells(k, 17).Value = wsDest.Cells(k, 13).Value
                                    outputCount = outputCount + 1
                                End If
                            End If
                        End If
                    End If
                Next k
                
                ' 既存で割当できなかった分を追加
                If insertFlag = True And wkSekouCnt > 0 Then
                    destRow = destRow + 1
                    lastRowDest = destRow
                    
                    ' 必要に応じてテンプレ行コピー（書式が要る場合のみ）
                    ' wsDest.Rows(startRowDest).Copy wsDest.Rows(destRow)
                    
                    wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, 1).Value  ' 品目コード
                    wsDest.Cells(destRow, 3).Value = ""                          ' 納品日
                    wsDest.Cells(destRow, 8).Value = wsSource.Cells(i, keyColSource3).Value ' 管理室
                    wsDest.Cells(destRow, 17).Value = wkSekouCnt                  ' 施工数量
                    wsDest.Cells(destRow, 10).Value = wsSource.Cells(i, 48).Value ' 購入単価
                    ' クリア
                    wsDest.Cells(destRow, 11).Value = "" ' 注文数量
                    wsDest.Cells(destRow, 13).Value = "" ' 納品数量
                    outputCount2 = outputCount2 + 1
                End If
            End If
        End If
    Next i
    
    If outputCount = 0 And outputCount2 = 0 Then GoTo NoData
    
    MsgBox "データをシートから取得しました：" & sheetName & vbCrLf & _
           "更新データ " & outputCount & " 件" & vbCrLf & _
           "追加データ " & outputCount2 & " 件", vbInformation
    GoTo Cleanup

NoData:
    MsgBox "対象の購入充当のデータが見つかりません: " & sheetName, vbExclamation
    GoTo Cleanup

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
Cleanup:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWorkbook Is Nothing Then srcWorkbook.Close False
    Set srcWorkbook = Nothing
End Sub

'施行通知書取込
'
'
'2025/07/02 O.Kanai
Sub ボタン4_Click()
    Dim i As Long, k As Long
    Dim wkNum As Double
    Dim keyCellDest As String
    Dim srcWorkbook As Workbook
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim destRow As Long
    Dim conditionValue As String
    Dim startRowSource As Long
    Dim startRowDest As Long
    Dim serchColSource As Long
    Dim serchColDest As Long
    Dim keyColSource As Long
    Dim keyColSource2 As Long
    Dim keyColSource3 As Long
    Dim outputCount As Long
    Dim outputCount2 As Long
    Dim keyData As String '品目コード
    Dim keyKanri As String
    Dim insertFlag As Boolean
    Dim wkSekouCnt As Double

    keyCellDest = "D2"
    
    If MsgBox("施行通知書を取り込みしますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    
    ' 1. ファイル選択
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="施行通知書を取得するExcelファイルを選択してください")
    If filePath = "False" Then
        MsgBox "操作がキャンセルされました。", vbExclamation
        GoTo Cleanup
    End If

    ' 2. ファイルを開く
    Set srcWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    sheetName = srcWorkbook.ActiveSheet.Name
    
    ' 3. 各種設定
    serchColSource = 1   'A列：整理番号
    serchColDest = 1     'A列：品目コード
    keyColSource = 133   'EC列：工種分類
    keyColSource2 = 1    'A列：整理番号
    keyColSource3 = 69   'BQ列：管理室（追加時のみ使用）
    startRowSource = 26
    startRowDest = 7
    
    Set wsSource = srcWorkbook.Sheets(sheetName)
    Set wsDest = ThisWorkbook.ActiveSheet
    
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, serchColSource).End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, serchColDest).End(xlUp).Row
    If lastRowDest > 10000 Then
        If wsDest.Cells(startRowDest, 1).Value = "" Then
            lastRowDest = startRowDest - 1
        Else
            lastRowDest = startRowDest
        End If
    End If
    If lastRowDest < startRowDest Then lastRowDest = startRowDest - 1
    
    destRow = lastRowDest
    If startRowSource >= lastRowSource Then GoTo NoData

    outputCount = 0   '更新件数
    outputCount2 = 0  '追加件数
    
    keyKanri = wsDest.Range(keyCellDest).Value
    
    ' データ取得
    For i = startRowSource To lastRowSource
        keyData = ""
        insertFlag = False
        If wsSource.Cells(i, keyColSource).Value = "購入充当" Then   '購入充当
            insertFlag = True
            keyData = wsSource.Cells(i, keyColSource2).Value
            If keyData <> "" Then
                '数値取得（今回が無ければ前回）
                If wsSource.Cells(i, 79).Value = "" Then  'CA列：今回
                    wkSekouCnt = Val(wsSource.Cells(i, 70).Value)  'BR列：前回
                Else
                    wkSekouCnt = Val(wsSource.Cells(i, 79).Value)  'CA列：今回
                End If
                If wkSekouCnt <> 0 Then
                    For k = startRowDest To lastRowDest
                        If keyData = wsDest.Cells(k, 1).Value Then 'キー一致（品目コード）
                            If wsDest.Cells(k, 17).Value = "" Or _
                               (wsDest.Cells(k, 17).Value > "" And _
                               (wsDest.Cells(k, 17).Value <> wsDest.Cells(k, 11).Value And _
                                wsDest.Cells(k, 17).Value <> wsDest.Cells(k, 13).Value)) Then
                                
                                If wsDest.Cells(k, 11).Value <> "" Then
                                    If wkSekouCnt <= (Val(wsDest.Cells(k, 11).Value) - Val(wsDest.Cells(k, 17).Value)) Then
                                        wsDest.Cells(k, 17).Value = Val(wsDest.Cells(k, 17).Value) + wkSekouCnt
                                        wkSekouCnt = 0
                                        outputCount = outputCount + 1
                                        insertFlag = False
                                        Exit For
                                    Else
                                        wkSekouCnt = wkSekouCnt - Val(wsDest.Cells(k, 11).Value) + Val(wsDest.Cells(k, 17).Value)
                                        wsDest.Cells(k, 17).Value = wsDest.Cells(k, 11).Value
                                        outputCount = outputCount + 1
                                    End If
                                ElseIf wsDest.Cells(k, 13).Value <> "" Then
                                    If wkSekouCnt <= (Val(wsDest.Cells(k, 13).Value) - Val(wsDest.Cells(k, 17).Value)) Then
                                        wsDest.Cells(k, 17).Value = Val(wsDest.Cells(k, 17).Value) + wkSekouCnt
                                        wkSekouCnt = 0
                                        outputCount = outputCount + 1
                                        insertFlag = False
                                        Exit For
                                    Else
                                        wkSekouCnt = wkSekouCnt - Val(wsDest.Cells(k, 13).Value) + Val(wsDest.Cells(k, 17).Value)
                                        wsDest.Cells(k, 17).Value = wsDest.Cells(k, 13).Value
                                        outputCount = outputCount + 1
                                    End If
                                End If
                            End If
                        End If
                    Next k
                    
                    If insertFlag = True And wkSekouCnt > 0 Then
                        destRow = destRow + 1
                        lastRowDest = destRow
                        
                        ' 必要に応じてテンプレ行コピー
                        ' wsDest.Rows(startRowDest).Copy wsDest.Rows(destRow)
                        
                        wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, 1).Value ' 品目コード
                        wsDest.Cells(destRow, 3).Value = ""                         ' 納品日
                        wsDest.Cells(destRow, 8).Value = wsSource.Cells(i, keyColSource3).Value ' 管理室
                        wsDest.Cells(destRow, 17).Value = wkSekouCnt               ' 施工数量
                        wsDest.Cells(destRow, 10).Value = wsSource.Cells(i, 48).Value ' 購入単価
                        ' クリア
                        wsDest.Cells(destRow, 11).Value = "" ' 注文数量
                        wsDest.Cells(destRow, 13).Value = "" ' 納品数量
                        
                        outputCount2 = outputCount2 + 1
                    End If
                End If
            End If
        End If
    Next i
    
    If outputCount = 0 And outputCount2 = 0 Then GoTo NoData
    
    MsgBox "データをシートから取得しました：" & sheetName & vbCrLf & _
           "更新データ " & outputCount & " 件" & vbCrLf & _
           "追加データ " & outputCount2 & " 件", vbInformation
    GoTo Cleanup

NoData:
    MsgBox "対象の購入充当のデータが見つかりません: " & sheetName, vbExclamation
    GoTo Cleanup

ErrHandler:
    MsgBox "エラーが発生しました: " & Err.Description, vbExclamation
Cleanup:
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWorkbook Is Nothing Then srcWorkbook.Close False
    Set srcWorkbook = Nothing
End Sub

