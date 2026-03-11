VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMonthSelector 
   Caption         =   "処理対象月の選択"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMonthSelector.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmMonthSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' キャンセルされたかどうかのフラグ
Public IsCancelled As Boolean
' 選択されたシート名のコレクション
Public SelectedSheets As Collection

' イベント抑止フラグ（全選択時の連動ループ用）
Private isEventsDisabled As Boolean

'=========================================================
' 外部から呼び出してシート一覧をセットする
'=========================================================
Public Sub SetSheetList(ByVal srcWorkbook As Workbook)
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet
    Dim i As Long
    Dim sheetName As String
    
    ' 存在するシートのうち「受領分」を含むものを辞書に格納
    For Each ws In srcWorkbook.Worksheets
        If InStr(ws.Name, "受領分") > 0 Then
            dict(ws.Name) = True
        End If
    Next ws
    
    ' 正規順序で並べてリストに追加
    Dim orderedSheets(14) As String
    orderedSheets(0) = "1月受領分"
    orderedSheets(1) = "2月受領分"
    orderedSheets(2) = "3月受領分"
    orderedSheets(3) = "4月受領分"
    orderedSheets(4) = "5月受領分"
    orderedSheets(5) = "6月受領分"
    orderedSheets(6) = "7月受領分"
    orderedSheets(7) = "8月受領分"
    orderedSheets(8) = "9月受領分"
    orderedSheets(9) = "10月受領分"
    orderedSheets(10) = "11月受領分"
    orderedSheets(11) = "12月受領分"
    orderedSheets(12) = "(翌年)1月受領分"
    orderedSheets(13) = "(翌年)2月受領分"
    orderedSheets(14) = "(翌年)3月受領分"
    
    With Me.lstMonths
        .Clear
        ' 先頭に全選択項目を追加
        .AddItem "【全受領月シート対象】"
        
        ' 正規順序で存在するシートのみ追加
        For i = 0 To 14
            If dict.Exists(orderedSheets(i)) Then
                .AddItem orderedSheets(i)
            End If
        Next i
        
        ' 正規順序に含まれないシート名も末尾に追加
        ' （例：独自名称のシートが存在する場合の保険）
        For Each ws In srcWorkbook.Worksheets
            If InStr(ws.Name, "受領分") > 0 Then
                Dim alreadyAdded As Boolean
                alreadyAdded = False
                For i = 0 To 14
                    If ws.Name = orderedSheets(i) Then
                        alreadyAdded = True
                        Exit For
                    End If
                Next i
                If Not alreadyAdded Then
                    .AddItem ws.Name
                End If
            End If
        Next ws
        
        .ListStyle = 1   ' チェックボックス形式
        .MultiSelect = 1 ' 複数選択
    End With
    
    Set dict = Nothing
End Sub

'=========================================================
' 初期化処理
'=========================================================
Private Sub UserForm_Initialize()
    IsCancelled = True
    Set SelectedSheets = New Collection
    isEventsDisabled = False
    
    With Me.lstMonths
        .ListStyle = 1
        .MultiSelect = 1
    End With
End Sub

'=========================================================
' リストボックスの変更イベント（全選択の連動）
'=========================================================
Private Sub lstMonths_Change()
    If isEventsDisabled Then Exit Sub
    
    Dim i As Long
    
    ' 「【全受領月シート対象】」（Index 0）がクリックされた場合
    If Me.lstMonths.ListIndex = 0 Then
        isEventsDisabled = True
        
        Dim isChecked As Boolean
        isChecked = Me.lstMonths.Selected(0)
        
        ' 1番目以降を全て連動
        For i = 1 To Me.lstMonths.ListCount - 1
            Me.lstMonths.Selected(i) = isChecked
        Next i
        
        isEventsDisabled = False
    
    ' 個別の月がクリックされた場合
    Else
        ' 一つでもチェックが外れたら全選択チェックも外す
        If Me.lstMonths.Selected(Me.lstMonths.ListIndex) = False Then
            isEventsDisabled = True
            Me.lstMonths.Selected(0) = False
            isEventsDisabled = False
        End If
    End If
End Sub

'=========================================================
' 実行ボタンクリック
'=========================================================
Private Sub btnRun_Click()
    Dim i As Long
    Dim isSelected As Boolean
    
    isSelected = False
    Set SelectedSheets = New Collection
    
    With Me.lstMonths
        For i = 1 To .ListCount - 1
            If .Selected(i) Then
                SelectedSheets.Add .List(i)
                isSelected = True
            End If
        Next i
    End With
    
    If Not isSelected Then
        MsgBox "処理対象の月を選択してください。", vbExclamation
        Exit Sub
    End If
    
    IsCancelled = False
    Me.Hide
End Sub

'=========================================================
' キャンセルボタンクリック
'=========================================================
Private Sub btnCancel_Click()
    IsCancelled = True
    Me.Hide
End Sub
