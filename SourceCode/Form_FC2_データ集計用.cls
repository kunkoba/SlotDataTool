VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC2_データ集計用"
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
    Call subタイトル.Form.Procタイトル更新("データ集計", E_Color_DeepBlue)
    
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
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理
    Me.subForm1.SourceObject = Me.name & "sub"
    Call listグループ項目1_AfterUpdate
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnクリア1_Click()
On Error Resume Next
    listグループ項目1 = Null
    Call Proc動的集計クエリ反映処理
    
End Sub

' ======================================================
Private Sub btnクリア2_Click()
On Error Resume Next
    listグループ項目2 = Null
    Call Proc動的集計クエリ反映処理
    
End Sub

' ======================================================
Private Sub btnクリア3_Click()
On Error Resume Next
    listグループ項目3 = Null
    Call Proc動的集計クエリ反映処理
    
End Sub



' ======================================================
Private Sub btnクリア日付_Click()
On Error Resume Next
    comb開始日 = Null
    Call Proc動的集計クエリ反映処理

End Sub

' ======================================================
Private Sub comb開始日_AfterUpdate()
On Error Resume Next
    Call Proc動的集計クエリ反映処理

End Sub

' ======================================================
Private Sub btn開始日_前日_Click()
On Error Resume Next
    日付加減算 comb開始日, -1
    
End Sub

' ======================================================
Private Sub btn開始日_翌日_Click()
On Error Resume Next
    日付加減算 comb開始日, 1
    
End Sub

' ******************************************************
' 日付移動
' ******************************************************
Private Sub 日付加減算(ByRef target As Control, ByVal offset As Long)
On Error Resume Next
    If IsDate(target.value) Then
        target.value = DateAdd("d", offset, target.value)
        Call Proc動的集計クエリ反映処理
    End If
    
End Sub



' ======================================================
Private Sub listグループ項目1_AfterUpdate()
On Error Resume Next
    Call Proc動的集計クエリ反映処理
    
End Sub

' ======================================================
Private Sub listグループ項目2_AfterUpdate()
On Error Resume Next
    Call Proc動的集計クエリ反映処理

End Sub

' ======================================================
Private Sub listグループ項目3_AfterUpdate()
On Error Resume Next
    Call Proc動的集計クエリ反映処理

End Sub

' ======================================================
Private Sub Proc動的集計クエリ反映処理()
    Const methodName As String = "Proc動的集計クエリ反映処理"
    ' 宣言部
    Dim sql As String
    Dim where As String
On Error GoTo ErrHandler
    
    ' メイン処理
    If Nz(comb開始日, "") <> "" Then
        If where <> "" Then where = where & " AND "
        where = where & "日付 >= " & Convert条件値文字列(comb開始日)
    End If
    
    sql = Proc動的集計クエリSQL生成(listグループ項目1, listグループ項目2, listグループ項目3, where)
    Call Procクエリ内SQLを書き換え("D_SLOT集計クエリ", sql)
    
    Call Procフィルタ情報を保持したままサブフォーム更新(Me, Me.subForm1.name)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub




