VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC3_データ日付推移"
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
    Call subタイトル.Form.Procタイトル更新("データ日付推移", E_Color_DeepBlue)

    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    Me.Undo
    Call Update集計元Query(Null, Null, Null, Null)

End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
    Select Case KeyCode
        Case vbKeyRight
            Call btn指定日_翌日_Click
        Case vbKeyLeft
            Call btn指定日_前日_Click
    End Select

End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理
    Me.subForm1.SourceObject = Me.name & "sub"

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btn今日_Click()
On Error Resume Next
    text今日 = Date
    Call Me.subForm1.Form.日付列セット(text今日)

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
' ======================================================
Private Sub btn指定日_10日前_Click()
On Error Resume Next
    日付加減算 text今日, -10

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
        Call Me.subForm1.Form.日付列セット(text今日)
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



' ======================================================
Private Sub btn切替1_Click()
On Error Resume Next
    Procサブフォーム切替 "sub", text今日
    
End Sub

' ======================================================
Private Sub btn切替2_Click()
On Error Resume Next
    Procサブフォーム切替 "sub2", text今日
    
End Sub


' ======================================================
Private Sub Procサブフォーム切替(ByVal subForm名 As String, ByVal 今日 As Date)
    Const methodName As String = "Procサブフォーム切替"
    Dim where As String
    Dim order As String
On Error GoTo ErrHandler

    DoCmd.Echo False

    With Me.subForm1.Form
        ' 現在のフィルタ・並び順を退避（クエリ名部分削除）
        where = Replace(.filter, "[D_日付推移クエリ].", "")
        where = Replace(where, "[D_日付推移クエリ2].", "")
        order = Replace(.OrderBy, "[D_日付推移クエリ].", "")
        order = Replace(order, "[D_日付推移クエリ2].", "")
    End With
    
    ' サブフォームを切り替え
    Me.subForm1.SourceObject = Me.name & subForm名
    
    ' 日付列のセット（※呼び出し先フォームのメソッド）
    Call Me.subForm1.Form.日付列セット(今日)
    
    ' フィルタと並び順を復元
    With Me.subForm1.Form
        .filter = where
        .OrderBy = order
        .FilterOn = True
        .OrderByOn = True
    End With
    
ErrHandler:
    DoCmd.Echo True
    Call ErrorSave(methodName, True) '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



