VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F02_アプリ情報_入力"
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
    Call subタイトル.Form.Procタイトル更新("アプリ情報", E_Color_Blue)
        
    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
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
'        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnバージョン更新_Click()
On Error Resume Next
    textバージョン = "v" & Format(Date, "dd_mm_yyyy")

End Sub

' ======================================================
Private Sub btnリリース日_Click()
On Error Resume Next
    textリリース日 = Date

End Sub

' ======================================================
Private Sub btn制限解除クリア_Click()
On Error Resume Next
    text解除コード = Null
    text解除日 = Null

End Sub





' ======================================================
Private Sub btnアプリ設定初期化_Click()
On Error Resume Next

    If ShowConfirm("初期化前確認", "設定を初期化してもよろしいですか？") <> vbYes Then Exit Sub
    
    Call Clearアプリ起動時設定
    Me.Requery
    
End Sub


' ======================================================
'■ S_アプリ設定の指定項目をすべてNULLにする
' ======================================================
Public Sub Clearアプリ起動時設定()
    Const methodName As String = "Clearアプリ起動時設定"
    ' 宣言部
    Dim strSQL As String
On Error GoTo ErrHandler
        
    ' SQL作成
    strSQL = "UPDATE S_アプリ設定 " & _
             "SET " & _
             "アプリ起動日 = NULL, " & _
             "アプリ有効期限 = NULL, " & _
             "秘密コード = NULL, " & _
             "暗号キー = NULL, " & _
             "解除コード = NULL, " & _
             "アプリ解除日 = NULL, " & _
             "MACアドレス1 = NULL, " & _
             "MACアドレス2 = NULL, " & _
             "MACアドレス3 = NULL;"
    
    ' SQL実行
    CurrentDb.Execute strSQL, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub






