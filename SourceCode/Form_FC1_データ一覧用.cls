VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC1_データ一覧用"
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
    Call subタイトル.Form.Procタイトル更新("データ参照（一覧表示）", E_Color_DeepBlue)

    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
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
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    Me.Undo
    Call Update集計元Query(Null, Null, Null, Null)

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
Private Sub btnクリア機種_Click()
On Error Resume Next

    list機種マスタ = Null
    list台番号マスタ = Null
    
    Call Proc集計元フィルタリング処理
    
End Sub


' ======================================================
Private Sub list機種マスタ_AfterUpdate()
On Error Resume Next
    
    list台番号マスタ = Null
    Call Proc集計元フィルタリング処理

    Call Update台番号Query(list機種マスタ)
    list台番号マスタ.RowSource = list台番号マスタ.RowSource
    
End Sub

' ======================================================
Private Sub list台番号マスタ_AfterUpdate()
On Error Resume Next

    list日付 = Null
    Call Proc集計元フィルタリング処理

End Sub

' ======================================================
Private Sub list日付_AfterUpdate()
On Error Resume Next

    list台番号マスタ = Null
    Call Proc集計元フィルタリング処理

End Sub




' ******************************************************
' ******************************************************
Private Sub Proc集計元フィルタリング処理()
    Const methodName As String = "Proc集計元フィルタリング処理"
    ' 宣言部
On Error GoTo ErrHandler
    
    ' メイン処理
    Call Update集計元Query(list機種マスタ, list台番号マスタ, list日付, list日付)
    
    Call Procフィルタ情報を保持したままサブフォーム更新(Me, Me.subForm1.name)
    
    list日付.Requery

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub



