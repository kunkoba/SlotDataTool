VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FA3_データ削除"
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
    Call subタイトル.Form.Procタイトル更新("積み上げProcデータ削除", E_Color_Red)
    
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
Private Sub list機種ID_AfterUpdate()
On Error Resume Next
    list開始日.Requery
    list終了日.Requery
    
    text機種名 = list機種ID.Column(1)
    
End Sub

' ======================================================
Private Sub list開始日_AfterUpdate()
On Error Resume Next
    list終了日 = list開始日
    list終了日.Requery
    
    text開始日 = list開始日
    text終了日 = list終了日

End Sub

' ======================================================
Private Sub list終了日_AfterUpdate()
On Error Resume Next
    text終了日 = list終了日

End Sub



' ======================================================
Private Sub btn削除_Click()
    Const methodName As String = "btn削除_Click"
    ' 宣言部
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    If Nz(Me.list機種ID, "") = "" Then Exit Sub

    If ShowConfirm("削除前確認", "選択されている機種のデータを削除します。　よろしいですか？") <> vbYes Then Exit Sub
    If ShowConfirm("削除前確認", "本当によろしいですか？") <> vbYes Then Exit Sub

    ' Procデータ削除
    Call Procデータ削除
'    Call クエリ実行_単一("Q99_集計区分_差分削除")
    CurrentDb.QueryDefs("Q99_集計区分_差分削除").Execute dbFailOnError

    ' 再表示
    Call ShowToast("正常にデータは削除されました。", 1)
        
    list機種ID = Null
    list開始日 = Null
    list終了日 = Null
    list機種ID.Requery
    list開始日.Requery
    list終了日.Requery

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' *******************************************************
Private Sub Procデータ削除()
    Const methodName As String = "Procデータ削除"
    ' 宣言部
    Dim sql As String
    Dim cond As String
    Dim val1 As Long
    Dim val3 As String
    Dim val5 As String
    Dim 開始日 As String
    Dim 終了日 As String
On Error GoTo ErrHandler
    
    ' メイン処理
    val1 = Nz(Me.list機種ID, 0)
    開始日 = Nz(Me.list開始日, "")
    終了日 = Nz(Me.list終了日, "")
    
    cond = "機種ID = " & val1 & " "
    
    ' ▼ 範囲指定条件
    If 開始日 <> "" Then
        cond = cond & "AND 日付 >= #" & Format(開始日, "yyyy/mm/dd") & "# "
    End If
    
    If 終了日 <> "" Then
        cond = cond & "AND 日付 <= #" & Format(終了日, "yyyy/mm/dd") & "# "
    End If
    
    sql = "DELETE FROM T_SLOTデータ履歴 WHERE " & cond & ";"
    
    ' 実行
    CurrentDb.Execute sql, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub



