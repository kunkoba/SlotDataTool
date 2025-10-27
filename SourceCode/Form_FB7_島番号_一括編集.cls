VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB7_島番号_一括編集"
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
    Call subタイトル.Form.Procタイトル更新("島番号一括編集", E_Color_Green)
    
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

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub



' ======================================================
Private Sub btn台番号編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB3_MstMachineNum
    DoCmd.Close acForm, Me.name

End Sub


' ======================================================
Private Sub btn更新_Click()
    Const methodName As String = "btn更新_Click"
    Dim nums As Variant
    Dim lists As String
    Dim sql As String
On Error GoTo ErrHandler

    Call LogStart(methodName)

    nums = Proc選択された要素を取得配列(list台番号)
    lists = Join(nums, ", ")
    If lists = "" Then Exit Sub
    
    ' SQL文字列を組み立て
    sql = "UPDATE M_台番号マスタ SET "
    sql = sql & "  島番号 = " & Convert条件値文字列(text島番号)
    sql = sql & " WHERE 台番号 in (" & lists & ")"
    
    ' 実行
    CurrentDb.Execute sql, dbFailOnError
    
    list台番号.Requery
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub


