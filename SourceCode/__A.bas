Attribute VB_Name = "__A"
Option Compare Database
Option Explicit


'' ======================================================
'Private Sub Form_Load()
'    Dim args() As String
'On Error Resume Next
'    Form.caption = App_アプリ名
'    Me.KeyPreview = True
''    Call subタイトル.Form.タイトル設定("出玉推移グラフ", E_Color_DeepBlue)
'
'    Call LogOpen(Me.name)
'    Call Procアプリ起動処理
'
'End Sub
'' ======================================================
'Private Sub Form_Close()
'On Error Resume Next
'    Call LogClose(Me.name)
'    Me.Undo
'
'End Sub
'' ======================================================
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'    If KeyCode = vbKeyEscape Then
'        DoCmd.Close acForm, Me.name
'        DoCmd.Close acForm, Me.Parent.name
'    End If
'
'End Sub
'' ======================================================
'Private Sub Procアプリ起動処理()
'    Const methodName As String = "Procアプリ起動処理"
'    Dim args
'On Error GoTo ErrHandler
'
'    ' メイン処理
'
'ErrHandler:
'    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
'
'End Sub




' ******************************************************
' フォームイベント（メッセージ出力あり）
' ******************************************************
Private Sub XXXXXX1_Click()
    Const methodName As String = "XXXXXX1_Click"
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ__3()
    Const methodName As String = "ZZZZZZ__3"
On Error GoTo ErrHandler

    ' メイン処理

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub





