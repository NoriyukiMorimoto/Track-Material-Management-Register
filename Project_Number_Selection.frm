VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Project_Number_Selection 
   Caption         =   "工事番号選択（最新順）"
   ClientHeight    =   13110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10635
   OleObjectBlob   =   "Project_Number_Selection.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "Project_Number_Selection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    ' 1. まずタイトルを確定させる
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("原図")
    Dim targetBranch As String: targetBranch = Trim(CStr(ws.Range("D1").Value))
    ' 末尾に「支店」が付いている場合は除去（ファイル名・タイトル生成で重複しないよう）
    If Right(targetBranch, 2) = "支店" Then targetBranch = Left(targetBranch, Len(targetBranch) - 2)
    Dim targetOffice As String: targetOffice = Trim(CStr(ws.Range("E1").Value))
    
    If targetOffice <> "" Then
        Me.Caption = targetBranch & "支店 " & targetOffice & " 工事選択"
    Else
        Me.Caption = targetBranch & "支店 工事選択"
    End If
    
    ' 2. ListViewの初期設定
    SetupListView
    
    ' 3. データ読込（LoadMasterDataToMemory内ではCaptionを変更しない）
    If IsEmpty(SharedMasterData) Then
        LoadMasterDataToMemory
    Else
        RefreshList ""
    End If
    
End Sub

'===========================================================
' フォーム表示完了後にTextBox1にフォーカス
'===========================================================
Private Sub UserForm_Activate()
    Me.TextBox1.SetFocus
End Sub

'===========================================================
' ListViewの初期設定
'===========================================================
Private Sub SetupListView()
    With Me.ListView1
        .View = lvwReport
        .FullRowSelect = True
        .Gridlines = True
        .HideColumnHeaders = False
        .MultiSelect = False
        
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , , "", 0                              ' 1列目：ダミー（幅0）
        .ColumnHeaders.Add , , "工事番号", 80, lvwColumnCenter   ' 2列目：中央揃え
        .ColumnHeaders.Add , , "工事件名", 450                  ' 3列目：左揃え
    End With
End Sub

'===========================================================
' データ読み込み
'===========================================================
Private Sub LoadMasterDataToMemory()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("原図")
    Dim folderPath As String: folderPath = "\\dt-ims\公開フォルダ\055_現場サポート室\★請求書一覧・材料管理格納\★各支店工事番号データ\"
    Dim targetYear As Long: targetYear = Year(Now)
    Dim targetBranch As String: targetBranch = Trim(CStr(ws.Range("D1").Value))
    ' 末尾に「支店」が付いている場合は除去（ファイル名生成で重複しないよう）
    If Right(targetBranch, 2) = "支店" Then targetBranch = Left(targetBranch, Len(targetBranch) - 2)
    Dim targetOffice As String: targetOffice = Trim(CStr(ws.Range("E1").Value))
    Dim years As Variant: years = Array(CStr(targetYear), CStr(targetYear - 1))
    Dim y As Integer, i As Long, fileName As String
    Dim tempCol As New Collection

    If targetBranch = "" Then
        MsgBox "支店名（D1）を先に入力してください。", vbExclamation
        Unload Me
        Exit Sub
    End If

    Dim cn  As Object
    Dim rs  As Object
    Dim sql As String
    Dim provider As String
    Dim filePath As String

    For y = 0 To 1
        fileName = Dir(folderPath & years(y) & "_" & targetBranch & "支店_工事現況表データ.xls*")
        If fileName <> "" Then
            filePath = folderPath & fileName
            
            ' --- ADODB接続 ---
            Set cn = CreateObject("ADODB.Connection")
            Set rs = CreateObject("ADODB.Recordset")
            
            Dim ext As String: ext = LCase(Right(filePath, 5))
            If InStr(ext, "xlsx") > 0 Or InStr(ext, "xlsm") > 0 Then
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=" & filePath & ";" & _
                           "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
            Else
                provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                           "Data Source=" & filePath & ";" & _
                           "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
            End If
            
            On Error GoTo ErrHandler
            cn.Open provider
            
            ' Sheet1から全列取得
            sql = "SELECT * FROM [Sheet1$]"
            rs.Open sql, cn, 1, 1  ' adOpenStatic, adLockReadOnly
            
            ' --- データ取得（J列=9列目、G列=6列目、AC列=28列目、AD列=29列目）---
            ' ADODBはHDR=YESのため0始まりインデックス
            Dim colG  As Long: colG = 6    ' G列
            Dim colJ  As Long: colJ = 9    ' J列
            Dim colAC As Long: colAC = 28  ' AC列
            Dim colAD As Long: colAD = 29  ' AD列
            
            ' 最終行から3行目まで逆順取得はADOBDでは困難なため
            ' 一旦全件取得して配列に格納後、逆順でCollectionに追加
            Dim tempArr() As Variant
            Dim rowCount As Long
            rowCount = 0
            
            rs.MoveFirst
            Do While Not rs.EOF
                rowCount = rowCount + 1
                rs.MoveNext
            Loop
            
            If rowCount > 0 Then
                ReDim tempArr(1 To rowCount, 1 To 4)
                rs.MoveFirst
                i = 1
                Do While Not rs.EOF
                    tempArr(i, 1) = Trim(CStr(IIf(IsNull(rs.fields(colG).Value), "", rs.fields(colG).Value)))  ' G列
                    tempArr(i, 2) = Trim(CStr(IIf(IsNull(rs.fields(colJ).Value), "", rs.fields(colJ).Value)))  ' J列
                    tempArr(i, 3) = Trim(CStr(IIf(IsNull(rs.fields(colAC).Value), "", rs.fields(colAC).Value))) ' AC列
                    tempArr(i, 4) = Trim(CStr(IIf(IsNull(rs.fields(colAD).Value), "", rs.fields(colAD).Value))) ' AD列
                    i = i + 1
                    rs.MoveNext
                Loop
                
                ' 逆順でCollectionに追加（最新順）
                For i = rowCount To 3 Step -1
                    If InStr(tempArr(i, 1), targetOffice) > 0 Or targetOffice = "" Then
                        Dim acVal As String: acVal = tempArr(i, 3)
                        Dim adVal As String: adVal = tempArr(i, 4)
                        tempCol.Add Array(tempArr(i, 2), acVal & adVal)
                    End If
                Next i
            End If
            
            rs.Close
            cn.Close
            Set rs = Nothing
            Set cn = Nothing
        End If
    Next y

    If tempCol.Count > 0 Then
        ReDim SharedMasterData(1 To tempCol.Count, 1 To 2)
        For i = 1 To tempCol.Count
            SharedMasterData(i, 1) = tempCol(i)(0)
            SharedMasterData(i, 2) = tempCol(i)(1)
        Next i
        RefreshList ""
    Else
        MsgBox "データが見つかりません。", vbExclamation
        Unload Me
    End If
    
    Exit Sub

ErrHandler:
    MsgBox "データ読込エラー: " & Err.Description, vbCritical
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not cn Is Nothing Then If cn.State = 1 Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Unload Me
End Sub

'===========================================================
' リスト表示
'===========================================================
Private Sub RefreshList(ByVal keyword As String)
    Dim i As Long
    Dim itmX As ListItem
    
    If IsEmpty(SharedMasterData) Then Exit Sub
    
    Me.ListView1.ListItems.Clear
    
    For i = 1 To UBound(SharedMasterData, 1)
        If keyword = "" Or InStr(1, SharedMasterData(i, 1) & SharedMasterData(i, 2), keyword, vbTextCompare) > 0 Then
            Set itmX = Me.ListView1.ListItems.Add(, , "")         ' 1列目：ダミー
            itmX.SubItems(1) = SharedMasterData(i, 1)            ' 2列目：工事番号
            itmX.SubItems(2) = SharedMasterData(i, 2)            ' 3列目：工事件名
        End If
    Next i
End Sub

'===========================================================
' 検索ボックス入力時
'===========================================================
Private Sub TextBox1_Change()
    RefreshList Me.TextBox1.Text
End Sub

'===========================================================
' OKボタン・ダブルクリックで確定
'===========================================================
Private Sub CommandButton1_Click()
    SetSelectedValue
End Sub

Private Sub ListView1_DblClick()
    SetSelectedValue
End Sub

'===========================================================
' 選択確定処理
'===========================================================
Private Sub SetSelectedValue()
    If Me.ListView1.SelectedItem Is Nothing Then
        MsgBox "工事番号を選択してください。", vbExclamation
        Exit Sub
    End If

    Dim wsOrigin As Worksheet
    Dim wsNew As Worksheet
    Dim projectNo As String
    Dim projectName As String
    
    projectNo = Me.ListView1.SelectedItem.SubItems(1)    ' 工事番号
    projectName = Me.ListView1.SelectedItem.SubItems(2)  ' 工事件名

    ' 「原図」シートを取得
    On Error Resume Next
    Set wsOrigin = ThisWorkbook.Worksheets("原図")
    On Error GoTo 0
    
    If wsOrigin Is Nothing Then
        MsgBox "「原図」シートが見つかりません。", vbCritical
        Exit Sub
    End If

    ' 同名シートが既に存在するか確認
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    sheetExists = False
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = projectNo Then
            sheetExists = True
            Set wsNew = ws
            Exit For
        End If
    Next ws

    If sheetExists Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("工事番号「" & projectNo & "」のシートが既に存在します。" & vbCrLf & _
                      "既存シートに切り替えますか？", vbYesNo + vbQuestion, "確認")
        If resp = vbYes Then
            Application.EnableEvents = False
            wsNew.Activate
            Application.EnableEvents = True
        End If
        Unload Me
        Exit Sub
    End If

    ' フォームを先に閉じてからシート操作
    Unload Me
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' 「原図」シートをコピー
    wsOrigin.Copy After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
    
    ' コピーされたシートを取得（必ずActiveSheetになる）
    Set wsNew = ActiveSheet
    
    ' シート名を工事番号に設定
    On Error Resume Next
    wsNew.Name = projectNo
    If Err.Number <> 0 Then
        MsgBox "シート名の変更に失敗しました。" & vbCrLf & _
               "工事番号：" & projectNo, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
    
    ' D3・E3に書き込み
    wsNew.Range("D3").Value = projectNo
    wsNew.Range("E3").Value = projectName

    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub


