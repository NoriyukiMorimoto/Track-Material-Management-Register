Attribute VB_Name = "ModuleExport"
' ▼ 参照設定：Microsoft Visual Basic for Applications Extensibility（必須）
Option Explicit

' 呼び出し用マクロ（引数がないので、マクロ一覧に表示されます）
Public Sub Run_VBA_Export()
    Call Export_All_VBA_Components
End Sub

Public Sub Export_All_VBA_Components(Optional ByVal targetBook As Workbook)
    Dim wb As Workbook
    Dim vbc As VBIDE.VBComponent
    Dim destRoot As String, dest As String
    Dim timeStamp As String
    Dim ext As String
    
    On Error GoTo EH
    Application.ScreenUpdating = False

    ' 対象ブック（未指定なら ThisWorkbook）
    If targetBook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetBook
    End If
    
    ' --- 出力先の設定（ご指定のパス） ---
    timeStamp = Format(Now, "yyyymmdd_hhmmss")
    destRoot = "C:\Users\n-morimoto\OneDrive - 大鉄工業株式会社\現場サポート室(OneDrive)\軌道材料管理\帳票7VBAコード\軌道材料管理簿VBA\VBA_Export_" & timeStamp
    
    ' フォルダが存在しない場合は作成
    If Dir(destRoot, vbDirectory) = vbNullString Then
        ' 親フォルダまでのパスが正しい前提で、今回の書き出し用フォルダを作成します
        MkDir destRoot
    End If

    ' すべての VBComponent を列挙して拡張子決定 → Export
    For Each vbc In wb.VBProject.VBComponents
        Select Case vbc.Type
            Case VBIDE.vbext_ComponentType.vbext_ct_StdModule
                ext = ".bas"
            Case VBIDE.vbext_ComponentType.vbext_ct_ClassModule
                ext = ".cls"
            Case VBIDE.vbext_ComponentType.vbext_ct_MSForm
                ext = ".frm"
            Case VBIDE.vbext_ComponentType.vbext_ct_Document
                ext = ".cls"
            Case Else
                ext = vbNullString
        End Select

        If Len(ext) > 0 Then
            dest = destRoot & "\" & vbc.Name & ext
            If Dir(dest, vbNormal) <> vbNullString Then Kill dest
            vbc.Export dest
        End If
    Next

    MsgBox "エクスポート完了しました。" & vbCrLf & "保存先: " & destRoot, vbInformation
    GoTo FinallyExit

EH:
    MsgBox "エラーが発生しました。パスが正しいか、または参照設定を確認してください。" & vbCrLf & _
           "Err " & Err.Number & ": " & Err.Description, vbCritical
FinallyExit:
    Application.ScreenUpdating = True
End Sub

