VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD3_グラフ_カスタム"
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
    Call subタイトル.Form.Procタイトル更新("グラフ（カスタム）", E_Color_DeepBlue)

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
    ' メイン処理
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
    
    list機種マスタ = Null
    Call Update台番号Query(0)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnクリア機種_Click()
On Error Resume Next
    list機種マスタ = Null
    
    Call Proc集計元フィルタリング処理

End Sub

' ======================================================
Private Sub btnクリア日付_Click()
On Error Resume Next
    comb開始日 = Null
    
    Call Proc集計元フィルタリング処理

End Sub

' ======================================================
Private Sub btn行クリア1_Click()
On Error Resume Next
    list項目1 = Null
    listフィルタ値1.RowSource = ""

    Call Proc集計元フィルタリング処理

End Sub



' ======================================================
Private Sub comb開始日_AfterUpdate()
On Error Resume Next
    Call Proc集計元フィルタリング処理

End Sub
' ======================================================
Private Sub btn開始日_前日_Click()
On Error Resume Next
    日付加減算 comb開始日, -1

End Sub

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
        Call Proc集計元フィルタリング処理
    End If
    
End Sub




' ======================================================
Private Sub list機種マスタ_AfterUpdate()
On Error Resume Next
    Call Proc集計元フィルタリング処理
    
End Sub



' ======================================================
Private Sub listフィルタ値1_AfterUpdate()
On Error Resume Next
    Call Proc集計元フィルタリング処理

End Sub


' ======================================================
Private Sub list項目1_AfterUpdate()
On Error Resume Next
    Call Procコンボボックス設定フィルタ(list項目1, listフィルタ値1)
    
    listフィルタ値1 = listフィルタ値1.ItemData(0)
    
    Call Proc集計元フィルタリング処理
    
End Sub

' ======================================================
' コンボの行ソースをセットする共通処理
' ======================================================
Private Sub Procコンボボックス設定フィルタ(Source As ListBox, target As ListBox)
    Const methodName As String = "Procコンボボックス設定フィルタ"
On Error Resume Next
    target = Null
    If Nz(Source, "") = "" Then
        target = Null
        target.RowSource = ""
    Else
        target.RowSource = Procリスト用SQL生成(Nz(Source, ""))
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub




' ******************************************************
Private Sub Proc集計元フィルタリング処理()
    Const methodName As String = "Proc集計元フィルタリング処理"
    ' 宣言部
    Dim where As String
On Error GoTo ErrHandler

    ' メイン処理
    If Nz(list項目1, "") <> "" Then
        where = list項目1 & " = " & Convert条件値文字列(listフィルタ値1)
    End If
    
    Call Update集計元Query(list機種マスタ, Null, comb開始日, Null, where)
    
    Me.subForm1.Form.Recalc

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub






