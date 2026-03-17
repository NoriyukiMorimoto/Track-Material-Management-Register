VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private oldD1 As Variant
Private oldF1 As Variant
Private oldD3 As Variant
Private cancelClear As Boolean
Private isF1Processing As Boolean  ' RunF1Process二重起動防止フラグ
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim branchName As String
    Dim officeListCol As Collection
    Dim officeList As String
    Dim resp As VbMsgBoxResult

    If Target.CountLarge > 1 Then Exit Sub
    
    On Error GoTo SafeExit
    Application.EnableEvents = False

    '--------------------------------------
    ' D1（支店名）変更時：クリア確認なしで出張所リストを表示
    '--------------------------------------
    If Not Application.Intersect(Target, Me.Range("D1")) Is Nothing Then
        branchName = Trim$(CStr(Me.Range("D1").Value))
        If branchName <> "" Then
            Set officeListCol = BuildOfficeListFromExternalMaster(branchName)
            UpdateE1Validation JoinCollection(officeListCol, ",")
            ShowBranchOfficeCombo officeListCol
        Else
            Me.Range("E1").Validation.Delete
            HideOfficeCombo
        End If
    End If

    '--------------------------------------
    ' E1（出張所名）変更時：クリア確認→クリア→単価取込
    '--------------------------------------
    If Not Application.Intersect(Target, Me.Range("E1")) Is Nothing Then
        resp = MsgBox("読込済データがクリアされますがよろしいですか？", _
                      vbExclamation + vbYesNo + vbDefaultButton2, "クリアの確認")

        If resp = vbNo Then
            Application.Undo
            GoTo SafeExit
        End If

        ' データクリア実行
        mod_Newprocess.ClearTargetData Me

        ' 単価取込処理を実行（二重起動防止）
        If Not isF1Processing Then
            isF1Processing = True
            mod_Newprocess.RunF1Process Me
            ' 完了後にEnableEvents復元→Activate→メッセージ表示→再計算
            Application.EnableEvents = True
            Me.Activate
            Dim completeMsg As String
            completeMsg = mod_Newprocess.BuildCompleteMessage(Me.Parent)
            isF1Processing = False
            Application.Calculation = xlCalculationAutomatic
            Application.Calculate
            If completeMsg <> "" Then
                MsgBox completeMsg, vbInformation, "完了"
            End If
        End If
    End If
    '--------------------------------------
    ' D3（工事番号）変更時
    '--------------------------------------
    If Not Intersect(Target, Me.Range("D3")) Is Nothing Then
        If Me.Range("D3").Value <> "" Then
            resp = MsgBox("選択された工事のデータ処理を実行しますか？", _
                          vbYesNo + vbQuestion, "実行確認")
            If resp = vbYes Then
                mod_Newprocess.RunDataProcessForD3 Me
            Else
                Me.Range("D3").Value = oldD3
            End If
        End If
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
'-----------------------------------------------------------
' 選択変更時に「変更前の値」を保持（Undo用）
'-----------------------------------------------------------
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    oldD1 = Me.Range("D1").Value
    oldF1 = Me.Range("E1").Value
    oldD3 = Me.Range("D3").Value
End Sub
'===========================================================
' 外部マスタから出張所一覧を取得（ADODB高速版）
'===========================================================
Private Function BuildOfficeListFromExternalMaster(ByVal branchName As String) As Collection
    Dim col As New Collection
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim targetBranch As String
    Dim masterPath As String
    
    targetBranch = Trim(branchName)
    ' 末尾に「支店」が付いている場合は除去（マスタのB列と形式を合わせるため）
    If Right(targetBranch, 2) = "支店" Then targetBranch = Left(targetBranch, Len(targetBranch) - 2)
    If targetBranch = "" Then Exit Function
    
    masterPath = mod_Newprocess.MASTER_FILE_PATH
    If masterPath = "" Or Dir(masterPath) = "" Then
        MsgBox "マスタファイルのパスが不正です。" & vbCrLf & masterPath, vbCritical
        Set BuildOfficeListFromExternalMaster = col
        Exit Function
    End If

    ' --- ADODB接続 ---
    Dim cn  As Object: Set cn = CreateObject("ADODB.Connection")
    Dim rs  As Object: Set rs = CreateObject("ADODB.Recordset")
    Dim sql As String
    
    ' 接続文字列（.xlsx と .xlsm 両対応）
    Dim ext As String: ext = LCase(Right(masterPath, 5))
    Dim provider As String
    If InStr(ext, "xlsx") > 0 Or InStr(ext, "xlsm") > 0 Then
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & masterPath & ";" & _
                   "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Else
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & masterPath & ";" & _
                   "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
    End If
    
    On Error GoTo ErrHandler
    cn.Open provider
    
    ' B列（支店名）でフィルタ、C列（出張所名）を取得
    ' ※シート名「出張所」の$を忘れずに
    sql = "SELECT * FROM [出張所$]"
    
    rs.Open sql, cn, 1, 1  ' adOpenStatic, adLockReadOnly
    
    Dim branchVal As String
    Dim officeVal As String
    
    Do While Not rs.EOF
branchVal = Trim(CStr(IIf(IsNull(rs.fields(1).Value), "", rs.fields(1).Value)))  ' B列：支店名
officeVal = Trim(CStr(IIf(IsNull(rs.fields(2).Value), "", rs.fields(2).Value)))  ' C列：出張所名
        
        ' 支店名の一致チェック（完全一致 / 「支店」付き / 前方一致）
        If branchVal = targetBranch Or _
           branchVal = targetBranch & "支店" Or _
           branchVal Like targetBranch & "*" Then
            
            If officeVal <> "" And Not dict.Exists(officeVal) Then
                dict.Add officeVal, True
                col.Add officeVal
            End If
        End If
        rs.MoveNext
    Loop
    
    If col.Count = 0 Then
        MsgBox "マスタから「" & targetBranch & "」に関連する支店が見つかりませんでした。" & vbCrLf & _
               "マスタ側の表記（B列）を確認してください。", vbInformation
    End If
    
    GoTo Cleanup

ErrHandler:
    MsgBox "マスタ読み込みエラー: " & Err.Description, vbCritical

Cleanup:
    On Error Resume Next
    If rs.State = 1 Then rs.Close
    If cn.State = 1 Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    
    Set BuildOfficeListFromExternalMaster = col
End Function

' --- 補助関数 (既存コードのセル位置修正版) ---

Private Sub UpdateE1Validation(ByVal officeList As String)
    On Error Resume Next
    With Me.Range("E1").Validation
        .Delete
        If officeList <> "" And Len(officeList) <= 255 Then
            .Add Type:=xlValidateList, Formula1:=officeList
        End If
    End With
End Sub

Private Sub ShowBranchOfficeCombo(ByVal officeListCol As Collection)
    Dim ole As OLEObject: Set ole = Me.OLEObjects("ComboBox1")
    Dim i As Long
    ole.Object.Clear
    For i = 1 To officeListCol.Count: ole.Object.AddItem officeListCol(i): Next i
    With ole
        .Left = Me.Range("E1").Left
        .Top = Me.Range("E1").Top
        .width = Me.Range("E1").width
        .Visible = True
    End With
    ' ドロップダウンを自動で開く
    ole.Object.DropDown
End Sub

Private Sub HideOfficeCombo()
    On Error Resume Next
    Me.OLEObjects("ComboBox1").Visible = False
End Sub

Private Function JoinCollection(ByVal col As Collection, ByVal delim As String) As String
    Dim i As Long, s As String
    If col Is Nothing Then Exit Function
    For i = 1 To col.Count
        s = s & IIf(s = "", "", delim) & col(i)
    Next i
    JoinCollection = s
End Function
'===========================================================
' 3. ComboBox1 変更（選択）時イベント
'===========================================================
Private Sub ComboBox1_Change()
    Dim newVal As String
    Dim resp As VbMsgBoxResult

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    ' 選択された値を取得
    newVal = Trim$(CStr(Me.OLEObjects("ComboBox1").Object.Value))
    If newVal = "" Then GoTo ExitHandler

    ' コンボボックスを隠す
    HideOfficeCombo

    ' クリア確認
    resp = MsgBox("読込済データがクリアされますがよろしいですか？", _
                  vbExclamation + vbYesNo + vbDefaultButton2, "クリアの確認")
    If resp = vbNo Then GoTo ExitHandler

    ' E1セルに値を書き込む（EnableEvents=Falseのまま書き込みでWorksheet_Changeを発火させない）
    Me.Range("E1").Value = newVal

    ' データクリア実行
    mod_Newprocess.ClearTargetData Me

    ' 単価取込処理を実行（二重起動防止）
    If Not isF1Processing Then
        isF1Processing = True
        mod_Newprocess.RunF1Process Me
        Application.EnableEvents = True
        Me.Activate
        Dim completeMsg As String
        completeMsg = mod_Newprocess.BuildCompleteMessage(Me.Parent)
        isF1Processing = False
        Application.Calculation = xlCalculationAutomatic
        Application.Calculate
        If completeMsg <> "" Then
            MsgBox completeMsg, vbInformation, "完了"
        End If
        Exit Sub
    End If

ExitHandler:
    Application.EnableEvents = True
End Sub
''===========================================================
' D2セルの設定：UserForm2でチェックボックス選択
'===========================================================
Private Sub SetD2RoomCombinations(ByVal officeName As String)
    Dim masterPath As String
    masterPath = mod_Newprocess.MASTER_FILE_PATH
    If masterPath = "" Or Dir(masterPath) = "" Then
        MsgBox "マスタファイルのパスが不正です。", vbCritical
        Exit Sub
    End If

    Dim cn As Object: Set cn = CreateObject("ADODB.Connection")
    Dim rs As Object: Set rs = CreateObject("ADODB.Recordset")

    Dim ext As String: ext = LCase(Right(masterPath, 5))
    Dim provider As String
    If InStr(ext, "xlsx") > 0 Or InStr(ext, "xlsm") > 0 Then
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & masterPath & ";" & _
                   "Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";"
    Else
        provider = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                   "Data Source=" & masterPath & ";" & _
                   "Extended Properties=""Excel 8.0;HDR=YES;IMEX=1"";"
    End If

    On Error GoTo ErrHandler
    cn.Open provider
    rs.Open "SELECT * FROM [出張所＆JR管理室$]", cn, 1, 1

    ' --- 管理室を収集（空白行は直前の出張所名を引き継ぐ）---
    Dim rooms(100) As String
    Dim roomCount As Long
    roomCount = 0
    Dim currentOffice As String
    currentOffice = ""

    Do While Not rs.EOF
        Dim cOffice As String
        Dim cRoom As String
        cOffice = Trim(CStr(IIf(IsNull(rs.fields(2).Value), "", rs.fields(2).Value))) ' C列：出張所名
        cRoom = Trim(CStr(IIf(IsNull(rs.fields(5).Value), "", rs.fields(5).Value)))   ' F列：管理室名

        If cOffice <> "" Then currentOffice = cOffice

        If currentOffice = officeName And cRoom <> "" Then
            rooms(roomCount) = cRoom
            roomCount = roomCount + 1
        End If
        rs.MoveNext
    Loop

    If roomCount = 0 Then
        MsgBox "「" & officeName & "」に対応する管理室が見つかりませんでした。", vbInformation
        GoTo Cleanup
    End If

    ' --- UserForm2を起動してチェックボックス選択 ---
    With UserForm2
        .InitRooms rooms, roomCount
        .Show
    End With

    GoTo Cleanup

ErrHandler:
    MsgBox "管理室取得エラー: " & Err.Description, vbCritical
Cleanup:
    On Error Resume Next
    If rs.State = 1 Then rs.Close
    If cn.State = 1 Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
End Sub

'===========================================================
' D2ダブルクリック時：管理室選択フォームを表示
'===========================================================
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)

    ' D2ダブルクリック：管理室選択フォームを表示
    If Not Intersect(Target, Me.Range("D2")) Is Nothing Then
        Cancel = True
        If Me.Range("E1").Value = "" Then
            MsgBox "先に出張所名（F1）を選択してください。", vbExclamation
            Exit Sub
        End If
        Dim officeNameF1 As String
        officeNameF1 = Trim(CStr(Me.Range("E1").Value))
        Call SetD2RoomCombinations(officeNameF1)
    End If

    ' D3ダブルクリック：工事番号選択フォームを表示
    If Not Intersect(Target, Me.Range("D3")) Is Nothing Then
        Cancel = True
        If Me.Range("D1").Value = "" Then
            MsgBox "先に支店名（D1）を入力してください。", vbExclamation
            Exit Sub
        End If
        Project_Number_Selection.Show
    End If

End Sub


