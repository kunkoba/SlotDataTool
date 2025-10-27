VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD2_グラフ_比較"
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
    Call subタイトル.Form.Procタイトル更新("グラフ（比較）", E_Color_DeepBlue)
    
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
    
    Call Procグラフ差し替え処理

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub





' ======================================================
Private Sub btnクリア機種_Click()
On Error Resume Next
    list機種マスタ = Null
    Call Update台番号Query(0)
    
    Call Procグラフ差し替え処理
    Call Proc集計元フィルタリング処理

End Sub

' ======================================================
Private Sub btn集計項目_Click()
On Error Resume Next
    list集計項目 = Null
    listフィルタ値.RowSource = ""
    listフィルタ値 = Null
    
    Call Procグラフ差し替え処理
    Call Proc集計元フィルタリング処理

End Sub

' ======================================================
Private Sub btnクリア日付_Click()
On Error Resume Next
    comb開始日 = Null
    
    Call Proc集計元フィルタリング処理
    Call Procグラフ差し替え処理
    
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
        Call Proc集計元フィルタリング処理
        Call Procグラフ差し替え処理
    End If

End Sub




' ======================================================
Private Sub list機種マスタ_AfterUpdate()
On Error Resume Next

    Call Procグラフ差し替え処理
    Call Proc集計元フィルタリング処理
    
End Sub

' ======================================================
Private Sub list集計項目_AfterUpdate()
    ' 宣言部
    Dim フィルタ値 As String
On Error Resume Next

    Call フィルタ再設定(list集計項目, listフィルタ値)
'
'    フィルタ値 = 値変換(subForm2.Form.Get日付, list集計項目)
'
'    If フィルタ値 = "-1" Then
'        listフィルタ値 = listフィルタ値.ItemData(0)
'    Else
'        listフィルタ値 = フィルタ値
'    End If

    Call Procグラフ差し替え処理

End Sub

' ======================================================
Private Sub listフィルタ値_AfterUpdate()
On Error Resume Next
    Call Procグラフ差し替え処理

End Sub

' ******************************************************
Private Sub Proc集計元フィルタリング処理()
    Const methodName As String = "Proc集計元フィルタリング処理"
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Call Update集計元Query(list機種マスタ, Null, comb開始日, Null)
    
    Me.subForm1.Form.Recalc

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub

' ******************************************************
Private Sub Procグラフ差し替え処理()
    Const methodName As String = "グラフ差し替え処理"
    ' 宣言部
On Error GoTo ErrHandler

    ' メイン処理
    Call Me.subForm1.Form.Procグラフデータ差し替え処理_比較(IIf(Nz(list機種マスタ, "") = "", 1, 2), list集計項目, listフィルタ値)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub

' ******************************************************
Private Sub フィルタ再設定(base As ListBox, target As ListBox)
On Error Resume Next

    target = Null
    If Nz(base, "") = "" Then
        target = Null
        target.RowSource = ""
    Else
        target.RowSource = Procリスト用SQL生成(Nz(base, ""))
    End If
    
End Sub


