Attribute VB_Name = "mod_支払シート取込"
Option Explicit

' 支払金額計算シート取込
'VBA改修 n-morimoto

Private Const SHAREPOINT_FOLDER As String = "大鉄工業株式会社"
Private Const PAYMENT_FOLDER As String = "帳票7_支払金額計算シート - 帳票7_支払金額計算シート"
Private Const SRC_START_ROW As Long = 6
Private Const DEST_START_ROW As Long = 7

'===========================================================
' ベースフォルダパスを取得
'===========================================================
Private Function GetBaseFolder() As String
    Dim userName As String
    userName = Environ("USERNAME")
    GetBaseFolder = "C:\Users\" & userName & "\" & SHAREPOINT_FOLDER & "\" & PAYMENT_FOLDER & "\"
End Function

'===========================================================
' メイン：ボタンから呼び出し
'===========================================================
Public Sub 支払シート取込()
    Dim wsDest As Worksheet
    Dim projectNo As String
    Dim officeName As String
    Dim baseFolder As String
    Dim targetFolder As String
    Dim filePath As String
    Dim srcWorkbook As Workbook

    Set wsDest = ActiveSheet
    projectNo = Trim(wsDest.Range("D3").Value)
    officeName = Trim(wsDest.Range("E1").Value)

    If projectNo = "" Then
        MsgBox "D3に工事番号を入力してください。", vbCritical
        Exit Sub
    End If
    If officeName = "" Then
        MsgBox "F1に出張所名を入力してください。", vbCritical
        Exit Sub
    End If

    baseFolder = GetBaseFolder()
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseFolder) Then
        MsgBox "支払金額計算シートフォルダが見つかりませんでした。" & vbCrLf & _
               "パス：" & baseFolder & vbCrLf & vbCrLf & _
               "SharePointの同期状態を確認してください。", vbCritical
        Set fso = Nothing
        Exit Sub
    End If
    targetFolder = GetOfficeFolder(baseFolder, officeName, fso)
    Set fso = Nothing
    If targetFolder = "" Then
        MsgBox "出張所フォルダが見つかりませんでした。" & vbCrLf & _
               "出張所名：" & officeName, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ChDir targetFolder
    On Error GoTo 0

    filePath = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*.xlsm", _
        Title:="支払金額計算シートを選択 [" & officeName & "]")
    If filePath = "False" Then
        MsgBox "キャンセルされました。", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set srcWorkbook = Workbooks.Open(filePath, ReadOnly:=True)
    frmMonthSelector.SetSheetList srcWorkbook
    frmMonthSelector.Show

    If frmMonthSelector.IsCancelled Then
        srcWorkbook.Close False
        MsgBox "キャンセルされました。", vbExclamation
        GoTo Cleanup
    End If

    Dim addCount As Long
    Dim skipCount As Long
    Call ProcessTransfer(srcWorkbook, frmMonthSelector.selectedSheets, _
                         wsDest, projectNo, addCount, skipCount)

    'srcWorkbook.Close False  'テスト中はコメントアウト

    MsgBox "処理が完了しました。" & vbCrLf & _
           "追記件数：" & addCount & " 件" & vbCrLf & _
           "スキップ件数：" & skipCount & " 件（重複）", vbInformation
    GoTo Cleanup

ErrHandler:
    MsgBox "エラーが発生しました：" & Err.Description, vbCritical
    On Error Resume Next
    If Not srcWorkbook Is Nothing Then srcWorkbook.Close False
Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Set srcWorkbook = Nothing
End Sub

'===========================================================
' 出張所サブフォルダを特定する
'===========================================================
Private Function GetOfficeFolder(ByVal baseFolder As String, _
                                  ByVal officeName As String, _
                                  ByVal fso As Object) As String
    Dim folder As Object
    Dim subFolder As Object
    Set folder = fso.GetFolder(baseFolder)
    For Each subFolder In folder.SubFolders
        If InStr(subFolder.Name, officeName) > 0 Then
            GetOfficeFolder = subFolder.Path & "\"
            Exit Function
        End If
    Next subFolder
    GetOfficeFolder = ""
End Function

'===========================================================
' 日付文字列から曜日など余分な文字を除去
'===========================================================
Private Function CleanDate(ByVal dateStr As String) As String
    Dim pos As Long
    pos = InStr(dateStr, " ")
    If pos > 0 Then
        CleanDate = Left(dateStr, pos - 1)
    Else
        CleanDate = dateStr
    End If
End Function

'===========================================================
' 転記処理本体
' データ収集→ソート→重複チェック→一括書込
'===========================================================
Private Sub ProcessTransfer(ByVal srcWb As Workbook, _
                             ByVal selectedSheets As Collection, _
                             ByVal wsDest As Worksheet, _
                             ByVal projectNo As String, _
                             ByRef addCount As Long, _
                             ByRef skipCount As Long)
    addCount = 0
    skipCount = 0

    ' --- Step1: 対象行を収集（Collectionを使用）---
    ' 各要素構成：0=日付シリアル, 1=品目コード数値, 2=G, 3=Sクリーン, 4=H, 5=I, 6=AA, 7=J, 8=L, 9=AB, 10=W
    Dim dataCol As New Collection
    Dim SEP As String: SEP = Chr(9)

    Dim sheetName As Variant
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dateVal As Double
    Dim codeVal As Double
    Dim rec As String
    Dim srcData As Variant
    Dim r As Long
    Dim cleanedDate As String

    For Each sheetName In selectedSheets
        Set ws = Nothing
        On Error Resume Next
        Set ws = srcWb.Sheets(CStr(sheetName))
        On Error GoTo 0
        If ws Is Nothing Then GoTo NextSheet

        lastRow = ws.Cells(ws.Rows.Count, "U").End(xlUp).Row
        If lastRow < SRC_START_ROW Then GoTo NextSheet

srcData = ws.Range(ws.Cells(SRC_START_ROW, 1), _
                   ws.Cells(lastRow, 28)).Value

        For r = 1 To UBound(srcData, 1)
            If Trim(CStr(srcData(r, 21))) = projectNo Then
                cleanedDate = CleanDate(CStr(srcData(r, 19)))
                dateVal = 0
                On Error Resume Next
                dateVal = CDbl(CDate(cleanedDate))
                On Error GoTo 0
                codeVal = Val(CStr(srcData(r, 7)))
rec = Format(dateVal, "000000000000.00") & SEP & _
      Format(codeVal, "000000000000.00") & SEP & _
      CStr(srcData(r, 7)) & SEP & _
      cleanedDate & SEP & _
      CStr(srcData(r, 8)) & SEP & _
      CStr(srcData(r, 9)) & SEP & _
      CStr(srcData(r, 27)) & SEP & _
      CStr(srcData(r, 10)) & SEP & _
      CStr(srcData(r, 12)) & SEP & _
      CStr(srcData(r, 28)) & SEP & _
      CStr(srcData(r, 23))
                dataCol.Add rec
            End If
        Next r
        srcData = Empty

NextSheet:
        Set ws = Nothing
    Next sheetName

    If dataCol.Count = 0 Then
        MsgBox "工事番号 " & projectNo & " に一致するデータが見つかりませんでした。", vbExclamation
        Exit Sub
    End If

    ' --- Step2: 配列に転してソート---
    Dim n As Long
    n = dataCol.Count
    Dim dataArr() As String
    ReDim dataArr(1 To n)
    Dim idx As Long
    For idx = 1 To n
        dataArr(idx) = CStr(dataCol(idx))
    Next idx
    Set dataCol = Nothing

    ' クイックソート
    Call QuickSort(dataArr, 1, n)

    ' --- Step3: 既存データをDictionaryに格納---
    Dim destLastRow As Long
    destLastRow = wsDest.Cells(wsDest.Rows.Count, "C").End(xlUp).Row
    If destLastRow < DEST_START_ROW Then destLastRow = DEST_START_ROW - 1

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    If destLastRow >= DEST_START_ROW Then
        Dim destData As Variant
        destData = wsDest.Range(wsDest.Cells(DEST_START_ROW, 1), _
                                wsDest.Cells(destLastRow, 27)).Value
        Dim d As Long
        Dim existKey As String
        Dim cDateStr As String
        For d = 1 To UBound(destData, 1)
            If IsDate(destData(d, 3)) Then
                cDateStr = Format(CDate(destData(d, 3)), "yyyy/m/d")
            ElseIf destData(d, 3) <> "" Then
                cDateStr = CStr(destData(d, 3))
            Else
                cDateStr = ""
            End If
            existKey = MakeKey( _
                CStr(destData(d, 1)), _
                cDateStr, _
                CStr(destData(d, 12)), _
                CStr(destData(d, 27)))
            If existKey <> "|||" Then dict(existKey) = True
        Next d
        
        destData = Empty
    End If

    ' --- Step4: 重複チェックして書込配列に積む---
    Dim writeData() As Variant
    ReDim writeData(1 To n, 1 To 8)
    Dim writeCount As Long
    writeCount = 0
    Dim fields() As String
    Dim newKey As String
    Dim maxDateVal As Double
    maxDateVal = 0

Dim dv As Double        ' ← ループの外に移動
    For idx = 1 To n
        fields = Split(dataArr(idx), SEP)
        newKey = MakeKey(fields(2), Format(CDate(fields(3)), "yyyy/m/d"), fields(10), fields(9))
        If dict.Exists(newKey) Then
            skipCount = skipCount + 1
            Debug.Print "重複スキップ: " & fields(2) & " | " & fields(3) & " | " & fields(4) & " | " & fields(10)
        Else
            writeCount = writeCount + 1
            writeData(writeCount, 1) = fields(2)
            writeData(writeCount, 2) = fields(3)
            writeData(writeCount, 3) = fields(4)
            writeData(writeCount, 4) = fields(5)
            writeData(writeCount, 5) = fields(6)
            writeData(writeCount, 6) = fields(7)
            writeData(writeCount, 7) = fields(8)
            writeData(writeCount, 8) = fields(9)  ' AB列
            dict(newKey) = True
            addCount = addCount + 1
            dv = CDbl(Val(fields(0)))
            If dv > maxDateVal Then maxDateVal = dv
        End If
    Next idx
    Set dict = Nothing
    Erase dataArr
    If writeCount = 0 Then Exit Sub

    ' --- Step5: 列ごとに一括書込---
    Dim startRow As Long
    ' C列の最終行の次行、ただし最小値は7行目
    startRow = Application.Max(DEST_START_ROW, destLastRow + 1)
    Dim colA() As Variant, colC() As Variant, colD() As Variant
    Dim colF() As Variant, colH() As Variant, colJ() As Variant, colK() As Variant
    Dim colZ() As Variant
    ReDim colA(1 To writeCount, 1 To 1)
    ReDim colC(1 To writeCount, 1 To 1)
    ReDim colH(1 To writeCount, 1 To 1)
    ReDim colJ(1 To writeCount, 1 To 1)
    ReDim colK(1 To writeCount, 1 To 1)
    ReDim colZ(1 To writeCount, 1 To 1)

    Dim w As Long
    For w = 1 To writeCount
        colA(w, 1) = writeData(w, 1)
        On Error Resume Next
        colC(w, 1) = CDate(writeData(w, 2))
        If Err.Number <> 0 Then colC(w, 1) = writeData(w, 2): Err.Clear
        On Error GoTo 0
        colH(w, 1) = writeData(w, 5)
        colJ(w, 1) = writeData(w, 6)
        colK(w, 1) = writeData(w, 7)
        colZ(w, 1) = writeData(w, 8)  ' AB列→Z列
    Next w
    Erase writeData

    ' 行を挿入（最終データ行の下にwriteCount行分）
    wsDest.Rows(startRow & ":" & startRow + writeCount - 1).Insert Shift:=xlDown

    ' 値を書き込み
    wsDest.Range(wsDest.Cells(startRow, "A"), wsDest.Cells(startRow + writeCount - 1, "A")).Value = colA
    With wsDest.Range(wsDest.Cells(startRow, "C"), wsDest.Cells(startRow + writeCount - 1, "C"))
        .Value = colC
        .NumberFormat = "yyyy/m/d"
    End With
    wsDest.Range(wsDest.Cells(startRow, "G"), wsDest.Cells(startRow + writeCount - 1, "G")).Value = colH
    wsDest.Range(wsDest.Cells(startRow, "I"), wsDest.Cells(startRow + writeCount - 1, "I")).Value = colJ
    wsDest.Range(wsDest.Cells(startRow, "L"), wsDest.Cells(startRow + writeCount - 1, "L")).Value = colK
    wsDest.Range(wsDest.Cells(startRow, "AA"), wsDest.Cells(startRow + writeCount - 1, "AA")).Value = colZ

    ' --- 計算式・書式設定（一括処理でメモリ効率化）---
    Dim endRow As Long
    endRow = startRow + writeCount - 1
    Dim rng As Range

    ' 計算式を先頭行に入力してAutoFillで複写
    wsDest.Cells(startRow, "K").Formula = "=SUM(J" & startRow & "*I" & startRow & ")"
    wsDest.Cells(startRow, "D").Formula = "=IF(A" & startRow & "=0,"""",VLOOKUP($A" & startRow & ",単価!$A:$J,3,FALSE))"
    wsDest.Cells(startRow, "E").Formula = "=IF(A" & startRow & "=0,"""",VLOOKUP($A" & startRow & ",単価!$A:$J,4,FALSE))"
    wsDest.Cells(startRow, "F").Formula = "=IF(A" & startRow & "=0,"""",VLOOKUP($A" & startRow & ",単価!$A:$J,5,FALSE))"
    wsDest.Cells(startRow, "H").Formula = "=IF(A" & startRow & "=0,"""",VLOOKUP($A" & startRow & ",単価!$A:$J,6,FALSE))"
    wsDest.Cells(startRow, "M").Formula = "=IF($A" & startRow & "=0,"""",ROUNDDOWN(L" & startRow & "*$I" & startRow & ",0))"
    wsDest.Cells(startRow, "Q").Formula = "=IF($A" & startRow & "=0,"""",ROUNDDOWN(P" & startRow & "*$I" & startRow & ",0))"
    wsDest.Cells(startRow, "R").Formula = "=(J" & startRow & "+L" & startRow & ")-P" & startRow & ""
    wsDest.Cells(startRow, "S").Formula = "=IF($A" & startRow & "=0,"""",ROUNDDOWN(R" & startRow & "*$I" & startRow & ",0))"
    wsDest.Cells(startRow, "T").Formula = "=IF(A" & startRow & "="""","""",A" & startRow & ")"
    wsDest.Cells(startRow, "U").Formula = "=IF(A" & startRow & "="""","""",D" & startRow & ")"
    wsDest.Cells(startRow, "V").Formula = "=IF(A" & startRow & "="""","""",E" & startRow & ")"
    wsDest.Cells(startRow, "W").Formula = "=IF(A" & startRow & "="""","""",F" & startRow & ")"
    wsDest.Cells(startRow, "X").Formula = "=IF(A" & startRow & "="""","""",G" & startRow & ")"
    wsDest.Cells(startRow, "Y").Formula = "=IF(A" & startRow & "="""","""",I" & startRow & ")"
    wsDest.Cells(startRow, "Z").Formula = "=IF(A" & startRow & "="""","""",J" & startRow & ")"

    ' AutoFillで全行に複写（writeCount=1の場合はスキップ）
    If writeCount > 1 Then
        Dim fillCols As Variant
        fillCols = Array("K", "D", "E", "F", "H", "M", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z")
        Dim fc As Variant
        For Each fc In fillCols
            wsDest.Range(wsDest.Cells(startRow, CStr(fc)), wsDest.Cells(endRow, CStr(fc))).FillDown
        Next fc
    End If

    ' 書式設定（範囲一括）
    With wsDest.Range(wsDest.Cells(startRow, "A"), wsDest.Cells(endRow, "AA"))
        .Font.Size = 11
    End With
    wsDest.Range(wsDest.Cells(startRow, "A"), wsDest.Cells(endRow, "AA")).RowHeight = 15#

    ' C列・G列：文字方向0度
    wsDest.Range(wsDest.Cells(startRow, "C"), wsDest.Cells(endRow, "C")).Orientation = 0
    wsDest.Range(wsDest.Cells(startRow, "G"), wsDest.Cells(endRow, "G")).Orientation = 0
    wsDest.Range(wsDest.Cells(startRow, "G"), wsDest.Cells(endRow, "G")).ShrinkToFit = True

    ' H?S列：数値形式・右詰め
    With wsDest.Range(wsDest.Cells(startRow, "H"), wsDest.Cells(endRow, "S"))
        .NumberFormat = "#,##0"
        .HorizontalAlignment = xlRight
    End With
    wsDest.Range(wsDest.Cells(startRow, "R"), wsDest.Cells(endRow, "R")).NumberFormat = "0"

    ' D列・E列：左詰め
    wsDest.Range(wsDest.Cells(startRow, "D"), wsDest.Cells(endRow, "D")).HorizontalAlignment = xlLeft
    wsDest.Range(wsDest.Cells(startRow, "E"), wsDest.Cells(endRow, "E")).HorizontalAlignment = xlLeft

    ' 下罫線（A?AA列）
    With wsDest.Range(wsDest.Cells(startRow, "A"), wsDest.Cells(endRow, "AA")).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlHairline
    End With
    ' 各行の下罫線（xlInsideHorizontal）
    If writeCount > 1 Then
        With wsDest.Range(wsDest.Cells(startRow, "A"), wsDest.Cells(endRow, "AA")).Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlHairline
        End With
    End If

    ' J列左罫線：二重線（全行一括）
    With wsDest.Range(wsDest.Cells(startRow, "J"), wsDest.Cells(endRow, "J")).Borders(xlEdgeLeft)
        .LineStyle = xlDouble
    End With

    ' --- 「合計」行のSUM式更新 ---
    Dim sumRow As Long
    Dim sumLastRow As Long
    sumLastRow = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    Dim sumCols As Variant
    sumCols = Array("K", "M", "Q", "S")
    Dim sc As Variant
    For sumRow = 7 To sumLastRow
        If CStr(wsDest.Cells(sumRow, "A").Value) = "合計" Then
            For Each sc In sumCols
                wsDest.Cells(sumRow, CStr(sc)).Formula = _
                    "=SUM(" & sc & "7:" & sc & (sumRow - 1) & ")"
            Next sc
        End If
    Next sumRow

    Dim gotoRow As Long
    gotoRow = Application.Max(startRow, startRow + writeCount - 1 - 10)
    Application.GoTo wsDest.Cells(gotoRow, "C"), True

    ' J3に今回追記分の最新日付を設定
    If maxDateVal > 0 Then
        wsDest.Range("J3").MergeArea.Value = CDate(maxDateVal)
        wsDest.Range("J3").MergeArea.NumberFormat = "yyyy/m/d"
    End If

End Sub

'===========================================================
' クイックソート（文字列配列を昇順ソート）
'===========================================================
Private Sub QuickSort(ByRef arr() As String, ByVal lo As Long, ByVal hi As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As String
    Dim i As Long, j As Long
    Dim tmp As String
    pivot = arr((lo + hi) \ 2)
    i = lo
    j = hi
    Do
        Do While arr(i) < pivot: i = i + 1: Loop
        Do While arr(j) > pivot: j = j - 1: Loop
        If i <= j Then
            tmp = arr(i): arr(i) = arr(j): arr(j) = tmp
            i = i + 1: j = j - 1
        End If
    Loop While i <= j
    Call QuickSort(arr, lo, j)
    Call QuickSort(arr, i, hi)
End Sub
'===========================================================
' 重複チェック用キー生成
'===========================================================
Private Function MakeKey(a As String, c As String, d As String, k As String) As String
    MakeKey = a & "|" & c & "|" & d & "|" & k
End Function
'===========================================================
' 重複チェック用キー生成（5列版）
'===========================================================
Private Function MakeKey5(a As String, c As String, d As String, _
                           k As String, ab As String) As String
    MakeKey5 = a & "|" & c & "|" & d & "|" & k & "|" & ab
End Function



