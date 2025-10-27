VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD3_グラフ_カスタムsub"
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
    
    Call Procチャート用SQL更新

End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    Me.Undo
    list項目1.RowSource = ""
    list項目2.RowSource = ""

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
Private Sub btn行クリア1_Click()
On Error Resume Next
    list項目1 = Null
    Call list項目1_AfterUpdate

End Sub

Private Sub btn行クリア2_Click()
On Error Resume Next
    list項目2 = Null
    Call list項目2_AfterUpdate

End Sub



' ======================================================
Private Sub list項目1_AfterUpdate()
On Error Resume Next
    
    Call Procチャート用SQL更新

End Sub

' ======================================================
Private Sub list項目2_AfterUpdate()
On Error Resume Next
    
    Call Procチャート用SQL更新

End Sub

Private Sub list値_AfterUpdate()
On Error Resume Next

    If Nz(list値, "") = "" Then list値 = "差枚数"
    Call Procチャート用SQL更新

End Sub



' ******************************************************
' 汎用：固定名のコンボボックスとチャートを使用して処理する
' ******************************************************
Public Sub Procチャート用SQL更新()
    Const methodName As String = "Procチャート用SQL更新"
    ' 宣言部
    Dim field1 As String
    Dim field2 As String
    Dim sumField As String
    Dim valueOrder As String
    Dim sql As String
On Error GoTo ErrHandler

    ' コンボボックスの値を取得
    field1 = Nz(Me.list項目1, "")
    field2 = Nz(Me.list項目2, "")
    sumField = Nz(Me.list値, "")

    ' 必須項目チェック
    If sumField = "" Then Exit Sub

    ' SQL生成
    sql = Proc動的チャート用SQL生成(field1, field2, sumField, "avg", "", "")

    ' データ設定＆再描画
    chartグラフ.RowSource = sql
    chartグラフ.Requery

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub

