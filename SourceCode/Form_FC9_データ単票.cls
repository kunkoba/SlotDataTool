VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC9_データ単票"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_アプリ名
    Me.KeyPreview = True
    
    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    Me.Undo

End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理
    list台番号 = Me.台番号
    text今日 = Me.日付

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub list台番号_AfterUpdate()
On Error Resume Next
    Call Procフィルタリング処理

End Sub

' ======================================================
Private Sub btn今日_Click()
On Error Resume Next
    text今日 = Date
    Call Procフィルタリング処理

End Sub
' ======================================================
Private Sub btn指定日_前日_Click()
On Error Resume Next
    日付加減算 text今日, -1

End Sub
' ======================================================
Private Sub btn指定日_翌日_Click()
On Error Resume Next
    日付加減算 text今日, 1

End Sub
' ******************************************************
' 日付移動
' ******************************************************
Private Sub 日付加減算(ByRef target As Control, ByVal offset As Long)
    Const methodName As String = "日付加減算"
    ' 宣言部
On Error GoTo ErrHandler
    
    ' メイン処理
    If IsDate(target.value) Then
        target.value = DateAdd("d", offset, target.value)
        Call Procフィルタリング処理
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub



' ******************************************************
' 日付移動
' ******************************************************
Private Sub Procフィルタリング処理()
    Const methodName As String = "Procフィルタリング処理"
    Dim where As String
On Error Resume Next
    
    where = where & " 台番号 = " & Convert条件値文字列(list台番号)
    where = where & " and 日付 = " & Convert条件値文字列(text今日)
    
    Me.Form.filter = where
    Me.Form.FilterOn = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub
