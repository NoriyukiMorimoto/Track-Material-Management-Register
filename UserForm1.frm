VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "支払金額シート選択"
   ClientHeight    =   4965
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
    Dim srcWorkbook As Workbook
    Dim srcSheet As Worksheet
    Dim destSheet As Worksheet
    Dim filePath As String
    Dim sheetName As String
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    Dim ws As Worksheet

    ' 1. ファイル選択ダイアログを表示
    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", _
        Title:="データを取得するExcelファイルを選択してください")

    ' キャンセルされた場合
    If filePath = "False" Then
        MsgBox "操作がキャンセルされました。", vbExclamation
        Exit Sub
    End If

    lblFileName.Caption = filePath

    ' 3. ファイルを開く
    On Error Resume Next
    Application.ScreenUpdating = False ' 画面更新を停止
    Set srcWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    If srcWorkbook Is Nothing Then
        MsgBox "ファイルを開けませんでした: " & filePath, vbExclamation
        Application.ScreenUpdating = True ' 画面更新を再開
        'srcWorkbook.Close False
        Exit Sub
    End If

    ' 4. シートを取得
    ComboBox1.Clear
    For Each ws In srcWorkbook.Sheets
        'シートをコンボボックスに設定
        ComboBox1.AddItem ws.Name
    Next ws
    
    Application.ScreenUpdating = True ' 画面更新を再開
    'srcWorkbook.Close False


End Sub


Private Sub CommandButton2_Click()
'選択されたシートからデータを取得
    Dim srcWorkbook As Workbook
    Dim srcSheet As Worksheet
    Dim sheetName As String
    Dim filePath As String
    
    If ComboBox1.Text = "" Then
        MsgBox "シートを選択してください。", vbExclamation
        Exit Sub
    Else
        sheetName = ComboBox1.Text
        filePath = lblFileName.Caption
    End If
    
    If MsgBox("シートからデータを追加しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub
    
    Application.ScreenUpdating = False ' 画面更新を停止
    Set srcWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    
    Set srcSheet = srcWorkbook.Sheets(sheetName)
    If srcSheet Is Nothing Then
        MsgBox "指定したシートが見つかりません: " & sheetName, vbExclamation
        'srcWorkbook.Close False
        Exit Sub
    End If

    ' 4. データ範囲を取得
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim i As Long, destRow As Long
    Dim conditionValue As String
    Dim startRowSource As Long
    Dim startRowDest As Long
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    
    Dim serchColSource As Long
    Dim serchColDest As Long
    Dim keyCellDest As String
    Dim keyColSource As Long
    Dim outputCount As Long

    serchColSource = 4 'D列：請求書番号
    serchColDest = 1 'A列：品目コード
    keyCellDest = "D3" '工番・件名(D3セルにある工番・件名を条件とする)
    keyColSource = 21 'U列：工事番号
    
    ' シートを設定
    Set wsSource = srcWorkbook.Sheets(sheetName)
    Set wsDest = ThisWorkbook.ActiveSheet
    
    startRowSource = 6
    ' データソースの最終行を取得
    lastRowSource = wsSource.Cells(startRowSource, serchColSource).End(xlDown).Row
    
    startRowDest = 7
    ' 結果を貼り付ける先の最終行を取得
    If wsDest.Cells(startRowDest, serchColDest) = "" Then
        lastRowDest = startRowDest - 1
    Else
        lastRowDest = wsDest.Cells(startRowDest, serchColDest).End(xlDown).Row
    End If
    ' 条件を取得
    conditionValue = wsDest.Range(keyCellDest).Value
    
    ' 結果を書き込む行番号を初期化
    destRow = lastRowDest ' データを書き込む先の開始行 (最後の行)
    
    ' データが無い場合終了
    If startRowSource >= lastRowSource Then Exit Sub
    
    outputCount = 0
    ' 条件に合うデータを取得
    For i = startRowSource To lastRowSource  ' データを取得
        If wsSource.Cells(i, keyColSource).Value = conditionValue Then
            If destRow = 6 Then
                '一行もない場合はコピーしない
            Else
                
                '行をコピー
                wsDest.Rows(destRow).Copy
            
                ' 次の行に挿入
                wsDest.Rows(destRow + 1).Insert Shift:=xlDown
            
                ' コピー元の選択を解除
                Application.CutCopyMode = False
            End If
            
            destRow = destRow + 1
            
            ' 条件に合致した行をコピー
            wsDest.Cells(destRow, 1).Value = wsSource.Cells(i, 7).Value   ' 品目コード
            wsDest.Cells(destRow, 3).Value = wsSource.Cells(i, 19).Value  ' 物品受領日（納品日）
            'wsDest.Cells(destRow, 7).Value = wsSource.Cells(i, 11).Value ' 単位
            wsDest.Cells(destRow, 8).Value = wsSource.Cells(i, 27).Value  ' 管理室　/*ADD　2025/01/29　*/
            wsDest.Cells(destRow, 10).Value = wsSource.Cells(i, 10).Value ' 購入単価
            wsDest.Cells(destRow, 13).Value = wsSource.Cells(i, 12).Value ' 購入数量
            '// クリア 2025/01/28
            wsDest.Cells(destRow, 11).Value = "" ' 注文数量
            wsDest.Cells(destRow, 17).Value = "" ' 施工数量
            '//
            
            outputCount = outputCount + 1
        End If
    Next i
    
    If outputCount = 0 Then
        MsgBox "対象の工番・件名のデータが見つかりません: " & sheetName, vbExclamation
        'srcWorkbook.Close False
        Application.ScreenUpdating = True ' 画面更新を再開
        Exit Sub
    Else
        MsgBox "データをシートから取得しました：" & sheetName & vbCrLf & "取得データ" & outputCount & "件", vbInformation
    End If
    Application.ScreenUpdating = True ' 画面更新を再開
    'srcWorkbook.Close False
    Unload Me

End Sub

Private Sub CommandButton3_Click()
'キャンセル
    Unload Me
End Sub
