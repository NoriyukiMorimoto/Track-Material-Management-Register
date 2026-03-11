VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "対応管理室選択"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private selectedRooms() As String
Private roomCount As Long

'===========================================================
' 初期化：管理室リストをチェックボックスで動的生成
'===========================================================
Public Sub InitRooms(ByRef rooms() As String, ByVal cnt As Long)
    Dim i As Long
    Dim chk As MSForms.CheckBox
    Dim ctrl As Control
    Dim existNames As Object
    
    roomCount = cnt
    ReDim selectedRooms(cnt - 1)
    
    Dim itemHeight As Single
    itemHeight = 25
    Dim topPos As Single
    topPos = 5
    
    ' --- Frame1の高さをチェックボックス数に応じて調整 ---
    ' Frame1.Top = 0 なので Frame1の下端 = frameHeight
    Dim frameHeight As Single
    frameHeight = cnt * itemHeight + 20
    Me.Frame1.Height = frameHeight
    Me.Frame1.Width = 310
    
    ' --- ボタン4つの位置をFrame1下端＋10に設定 ---
    Dim btnTop As Single
    btnTop = frameHeight + 10  ' Frame1.Top=0 なので frameHeight がそのまま下端
    
    With Me.cmdSelectAll
        .Top = btnTop
        .Left = 20
        .Width = 50
        .Height = 30
    End With
    With Me.cmdClearAll
        .Top = btnTop
        .Left = 90
        .Width = 50
        .Height = 30
    End With
    With Me.cmdOK
        .Top = btnTop
        .Left = 160
        .Width = 50
        .Height = 30
    End With
    With Me.cmdCancel
        .Top = btnTop
        .Left = 230
        .Width = 70
        .Height = 30
    End With
    
    ' --- フォーム全体の高さを調整 ---
    Me.Height = btnTop + 30 + 40  ' ボタン高さ30 ＋ 余白40

    ' --- 既存チェックボックスの名前一覧を収集 ---
    Set existNames = CreateObject("Scripting.Dictionary")
    For Each ctrl In Me.Frame1.Controls
        ctrl.Visible = False
        existNames(ctrl.Name) = True
    Next ctrl

    ' --- 必要な数だけチェックボックスを設定・表示 ---
    For i = 0 To cnt - 1
        Dim ctrlName As String
        ctrlName = "chk" & i
        
        If existNames.Exists(ctrlName) Then
            Set chk = Me.Frame1.Controls(ctrlName)
        Else
            Set chk = Me.Frame1.Controls.Add("Forms.CheckBox.1", ctrlName, True)
        End If
        
        With chk
            .Caption = rooms(i)
            .Left = 10
            .Top = topPos + i * itemHeight
            .Width = 270
            .Height = 20
            .Tag = CStr(i)
            .Value = False
            .Visible = True
        End With
        
        Set chk = Nothing
    Next i
    
    Set existNames = Nothing
End Sub
'===========================================================
' OKボタン：選択された管理室をカンマ区切りでD2に設定
'===========================================================
Private Sub cmdOK_Click()
    Dim result As String
    Dim ctrl As Control
    result = ""
    
    For Each ctrl In Me.Frame1.Controls
        If TypeName(ctrl) = "CheckBox" Then
            If ctrl.Value = True Then
                result = result & IIf(result = "", "", ",") & ctrl.Caption
            End If
        End If
    Next ctrl
    
    If result = "" Then
        MsgBox "管理室を1つ以上選択してください。", vbExclamation
        Exit Sub
    End If
    
    ' D2セルに書き込み
    Dim ws As Worksheet
    Set ws = ThisWorkbook.ActiveSheet
    Application.EnableEvents = False
    ws.Range("D2").Value = result
    Application.EnableEvents = True
    
    Unload Me
End Sub

'===========================================================
' キャンセルボタン
'===========================================================
Private Sub cmdCancel_Click()
    Unload Me
End Sub

'===========================================================
' 全選択ボタン
'===========================================================
Private Sub cmdSelectAll_Click()
    Dim ctrl As Control
    For Each ctrl In Me.Frame1.Controls
        If TypeName(ctrl) = "CheckBox" Then ctrl.Value = True
    Next ctrl
End Sub

'===========================================================
' 全解除ボタン
'===========================================================
Private Sub cmdClearAll_Click()
    Dim ctrl As Control
    For Each ctrl In Me.Frame1.Controls
        If TypeName(ctrl) = "CheckBox" Then ctrl.Value = False
    Next ctrl
End Sub
