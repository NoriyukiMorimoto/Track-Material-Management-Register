Attribute VB_Name = "mod_支払シート取込"
Option Explicit

'支払金額計算取り込み
'
'
'2025/01/22 O.Kanai
Public Sub 支払シート取込()
    '支払金額計算取り込み
    '単なる入力チェック＋フォーム呼び出し（本処理は UserForm1 側）
    If ActiveSheet.Range("D3").Text = "" Then
        MsgBox "工番・件名を入力してください", vbCritical
        Exit Sub
    End If
    UserForm1.Show
End Sub
