Attribute VB_Name = "mod_前期繰越"
Option Explicit

'残数量を前期繰越へ
'
'今期注文数量、2/4で使用をクリア
'2024/12/05 O.Kanai
Public Sub 前期繰越()
    Dim i As Long
    Dim iMax As Long
    Dim wkNum As Long
    Dim wsDest As Worksheet
    
    Set wsDest = ActiveSheet
    
    If MsgBox("残数量を前期繰越へ設定しますか？", vbYesNo + vbQuestion, "確認") = vbNo Then Exit Sub

    i = 7 '明細開始の行
    Do
        If i > 10000 Then Exit Do  '周り過ぎないように
        
        '品目コードが入力ある分のみ処理
        If wsDest.Range("A" & i) = "" Then
            Exit Do
        Else
            '残数量を取得し、前期繰越へ設定
            wkNum = Val(wsDest.Range("S" & i).Value)  '残数量(S列)
            wsDest.Range("K" & i).Value = wkNum       '前期繰越(K列)
            
            '今期注文数量、2/4で使用をクリア
            wsDest.Range("M" & i).ClearContents       '今期注文(M列)
            wsDest.Range("Q" & i).ClearContents       '2/4で使用(Q列)
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
            If Val(wsDest.Range("S" & i).Value) = 0 Then '残数量(S列)
                wsDest.Rows(i).Delete
            End If
        End If
        i = i - 1
    Loop
    
    MsgBox "前期繰越の設定が完了しました。", vbInformation
End Sub
