VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_xxxTEST2"
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



' ======================================================
' ======================================================
' ======================================================
Private Sub btnデータ取り込み_Click()
    Const methodName As String = "btnデータ取り込み_Click"
    ' 宣言部
On Error GoTo ErrHandler
    Call LogStart("btnデータ取り込み_Click")
    
    ' メイン処理
    Sleep 1000
'    Error 13   '型が一致しません
    Call ZZZZZZ1
    
    Call ShowToast("完了しました！", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ1()
    Const methodName As String = "ZZZZZZ_1"
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Sleep 1000
'    Error 13   '型が一致しません
    Call ZZZZZZ__2

ErrHandler:
    ' エラー
End Sub


' ******************************************************
' 中間処理（メッセージ出力なし）
' ******************************************************
Private Sub ZZZZZZ__2()
    Const methodName As String = "ZZZZZZ_2"
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Sleep 1000
'    Error 13   '型が一致しません

ErrHandler:
    ' エラー
End Sub
