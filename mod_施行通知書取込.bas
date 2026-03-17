Attribute VB_Name = "mod_施行通知書取込"
Option Explicit

'施行通知書取込
'
'
'2025/07/02 O.Kanai
Public Sub 施行通知書取込()
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
        If wsSource.Cells(i, keyColSource).Value = "購入充当" Then
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
                        If keyData = wsDest.Cells(k, 1).Value Then
                            If Cells(k, 16).Value = "" Or _
                               (Cells(k, 16).Value > "" And _
                               (Cells(k, 16).Value <> Cells(k, 10).Value And _
                                Cells(k, 16).Value <> Cells(k, 12).Value)) Then
                                
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


