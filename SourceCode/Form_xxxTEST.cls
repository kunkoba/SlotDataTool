VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_xxxTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit



' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Call LogOpen(Me.name)
    Form.caption = App_アプリ名
    Me.KeyPreview = True
    
    Call Proc画面読み込み時

End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)

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
Private Sub Proc画面読み込み時()
On Error Resume Next
    Call subタイトル.Form.Procタイトル更新("データ取り込み", E_Color_Blue)

End Sub



' ******************************************************
Private Sub コマンド83_Click()
    Const methodName As String = "コマンド83_Click"
On Error Resume Next

    ' メイン処理
    DoCmd.OpenForm Me.name & "2"

ErrHandler:
    ' エラー
    Call ErrorSave(methodName)
End Sub




' ======================================================
' ======================================================
' ======================================================
Private Sub btnデータ取り込み_Click()
    Const methodName As String = "btnデータ取り込み_Click"
    ' 宣言部
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    Sleep 500
'    Error 13   '型が一致しません
    Call ZZZZZZ_1
    
    Call ShowToast("完了しました！", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭

End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ_1()
    Const methodName As String = "ZZZZZZ_1"
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Sleep 500
'    Error 91
    Call ZZZZZZ__2

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後
    
End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ__2()
    Const methodName As String = "ZZZZZZ__2"
    Dim db As Database
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Sleep 500
'    Error 13   '型が一致しません
'    Error 91   '型が一致しません
    Call ZZZZZZ__3

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後
    
End Sub



' ******************************************************
Private Sub コマンド84_Click()
    Const methodName As String = "コマンド84_Click"
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
'    Error 94
    Call ZZZZZZ__3
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭
    
End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ__3()
    Const methodName As String = "ZZZZZZ__3"
    Dim db As Database
On Error GoTo ErrHandler

    ' メイン処理
    Sleep 500
    Error 13   '型が一致しません
    Error 91   '型が一致しません
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



