Attribute VB_Name = "mod_最終施行指示取込"
Option Explicit

'最終施行指示取り込み
'
'
'2025/01/22 O.Kanai
Public Sub 最終施行指示取込()
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
    If ActiveSheet.Range(keyCellDest).Text = "" Then
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
           InStr(keyKanri, wsSource.Cells(i, keyColSource3).Value) > 0 Then
           
            insertFlag = True
            keyData = wsSource.Cells(i, keyColSource2).Value
            keyData2 = wsSource.Cells(i, keyColSource3).Value
            If keyData <> "" Then
                wkSekouCnt = Val(wsSource.Cells(i, 43).Value)  ' 施工数量（列43）
                
                For k = startRowDest To lastRowDest
                    If keyData = wsDest.Cells(k, 1).Value And _
                       keyData2 = Cells(k, 7).Value Then
                       
                        If Cells(k, 16).Value = "" Or _
                           (Cells(k, 16).Value > "" And _
                           (Cells(k, 16).Value <> Cells(k, 10).Value And _
                            Cells(k, 16).Value <> Cells(k, 12).Value)) Then
                            
                            '注文数量 優先
                            If Cells(k, 10).Value <> "" Then
                                If wkSekouCnt <= (Val(Cells(k, 10).Value) - Val(Cells(k, 16).Value)) Then
                                    Cells(k, 16).Value = Val(Cells(k, 16).Value) + wkSekouCnt
                                    wkSekouCnt = 0
                                    outputCount = outputCount + 1
                                    insertFlag = False
                                    Exit For
                                Else
                                    wkSekouCnt = wkSekouCnt - Val(Cells(k, 10).Value) + Val(Cells(k, 16).Value)
                                    Cells(k, 16).Value = Cells(k, 10).Value
                                    outputCount = outputCount + 1
                                End If
                            ElseIf Cells(k, 12).Value <> "" Then
                                '納品数量 次優先
                                If wkSekouCnt <= (Val(Cells(k, 12).Value) - Val(Cells(k, 16).Value)) Then
                                    Cells(k, 16).Value = Val(Cells(k, 16).Value) + wkSekouCnt
                                    wkSekouCnt = 0
                                    outputCount = outputCount + 1
                                    insertFlag = False
                                    Exit For
                                Else
                                    wkSekouCnt = wkSekouCnt - Val(Cells(k, 12).Value) + Val(Cells(k, 16).Value)
                                    Cells(k, 16).Value = Cells(k, 12).Value
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
                    wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, 1).Value
                    wsDest.Cells(destRow, 3).Value = ""
                    Cells(destRow, 7).Value = wsSource.Cells(i, keyColSource3).Value
                    Cells(destRow, 16).Value = wkSekouCnt
                    Cells(destRow, 9).Value = wsSource.Cells(i, 48).Value
                    Cells(destRow, 10).Value = ""
                    Cells(destRow, 12).Value = ""
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


