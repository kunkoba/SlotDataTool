VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB1_機種マスタ"
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
    Call subタイトル.Form.Procタイトル更新("機種マスタ", 3)
    
    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)

    If flg最適化 Then
        DoCmd.OpenForm F01_OptimizeForm, , , , , , F11_MainMenu
    End If
    
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
    Me.list機種 = Me.list機種.ItemData(0)
    Call list機種_AfterUpdate

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub



' ======================================================
Private Sub Form_Dirty(Cancel As Integer)
On Error Resume Next
'    btn再計算.Visible = True
    btn元に戻す.Visible = True
    btn保存する.Visible = True
    
End Sub
' ======================================================
Private Sub Form_Undo(Cancel As Integer)
On Error Resume Next
'    btn再計算.Visible = False
    btn元に戻す.Visible = False
    btn保存する.Visible = False
    
End Sub



' ======================================================
Private Sub btn元に戻す_Click()
On Error Resume Next
    Me.Undo
    text機種名正式.SetFocus
'    btn再計算.Visible = False
    btn元に戻す.Visible = False
    btn保存する.Visible = False
    
End Sub

' ======================================================
Private Sub btn保存する_Click()
On Error Resume Next
    Me.Dirty = False
    text機種名正式.SetFocus
'    btn再計算.Visible = False
    btn元に戻す.Visible = False
    btn保存する.Visible = False

End Sub


' ======================================================
Private Sub btn再計算_Click()
    Const methodName As String = "btn再計算_Click"
    ' 宣言部
On Error GoTo ErrHandler
    
    Call LogStart(methodName)
    
    'チェック
    If Me.NewRecord Then GoTo ErrHandler

    If ShowConfirm("再計算前確認", "現在入力されている値を保存し、積み上げデータを再計算します。" & vbNewLine & "よろしいですか？") <> vbYes Then Exit Sub
    
'    Call btn再計算_Click
    Me.Dirty = False
    text機種名正式.SetFocus
'    btn再計算.Visible = False
    btn元に戻す.Visible = False
    btn保存する.Visible = False
    
    DoCmd.Hourglass True   ' ← くるくる開始（代替としても可）
    Screen.MousePointer = 11 ' ← くるくるカーソル開始
    
    '判別完了フラグ クリア
    Call Procクリア判別完了フラグ
    
    'クエリ実行3
    Call Procクエリリスト一括実行(Getクエリ一覧("Q3*"))
    
    '最適化
    Call Procファイル最適化
    
    Call ShowToast("再計算は正常に実行しました。", E_Color_Blue)
    
    flg最適化 = True

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    DoCmd.Hourglass False
    Screen.MousePointer = 0
        
End Sub


' ******************************************************
Public Sub Procクリア判別完了フラグ()
    Const methodName As String = "Procクリア判別完了フラグ"
    ' 宣言部
    Dim sql As String
On Error GoTo ErrHandler
    
    ' SQL文字列を組み立て
    sql = "UPDATE T_SLOT集計区分 S INNER JOIN T_SLOTデータ履歴 T ON  S.台番号 = T.台番号 AND S.日付 = T.日付 SET 判別完了フラグ = NULL WHERE T.機種ID = " & list機種
        
    ' 実行
    CurrentDb.Execute sql, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub





' ======================================================
Private Sub btn新規登録_Click()
On Error Resume Next
    list機種.value = Null
    Me.AllowAdditions = True
    DoCmd.GoToRecord , , acNewRec
    text機種名正式.SetFocus
    Me.AllowAdditions = False
    
End Sub


' ======================================================
Private Sub list機種_AfterUpdate()
On Error Resume Next
    Me.Undo
    If IsNull(Me.list機種) Then
        Me.FilterOn = False
    Else
        Me.filter = "[機種ID] = " & Me.list機種
        Me.FilterOn = True
    End If
    
End Sub

' ======================================================
Private Sub list機種_Click()
On Error Resume Next
    list機種.Requery
    
End Sub


' ======================================================
Private Sub btn調整値算出_Click()
On Error Resume Next
    DoCmd.OpenForm FB2_MstSamaisu
    Call Form_Dirty(False)
    
End Sub


' ======================================================
Private Sub btnシミュレーション_Click()
    Const methodName As String = "btnシミュレーション_Click"
    ' 宣言部
    Dim a As Double
    Dim b As Double
    Dim c As Double
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    text想定差枚数 = 想定差枚数(textゲーム数, textBB数, textRB数, textBB獲得枚数, textRB獲得枚数, text差枚数係数)
    
    a = 汎用設定判別(textゲーム数, textBB数, text設定1_BB, text設定2_BB, text設定3_BB, text設定4_BB, text設定5_BB, text設定6_BB)
    b = 汎用設定判別(textゲーム数, textRB数, text設定1_RB, text設定2_RB, text設定3_RB, text設定4_RB, text設定5_RB, text設定6_RB)
    c = 汎用設定判別(textゲーム数, textBB数 + textRB数, text設定1_合算, text設定2_合算, text設定3_合算, text設定4_合算, text設定5_合算, text設定6_合算)
    
    text設定判別 = 設定総合評価(a, b, c, textBB_重み, textRB_重み, text合算_重み, textゲーム数, text回転数_しきい値, text設定調整値)
    text機械割差枚数 = 機械割差枚数(textゲーム数, text設定判別, text設定1_機械割, text設定2_機械割, text設定3_機械割, text設定4_機械割, text設定5_機械割, text設定6_機械割)

    textBB率 = 当選率算出(textゲーム数, textBB数)
    textRB率 = 当選率算出(textゲーム数, textRB数)
    text合成率 = 当選率算出(textゲーム数, textBB数 + textRB数)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' ======================================================
Private Sub textBB数_AfterUpdate()
On Error Resume Next
    text想定差枚数 = ""
    text設定判別 = ""
    text機械割差枚数 = ""
    textBB率 = ""
    textRB率 = ""
    text合成率 = ""
    
End Sub
' ======================================================
Private Sub textRB数_AfterUpdate()
On Error Resume Next
    text想定差枚数 = ""
    text設定判別 = ""
    text機械割差枚数 = ""
    textBB率 = ""
    textRB率 = ""
    text合成率 = ""

End Sub
' ======================================================
Private Sub textゲーム数_AfterUpdate()
On Error Resume Next
    text想定差枚数 = ""
    text設定判別 = ""
    text機械割差枚数 = ""
    textBB率 = ""
    textRB率 = ""
    text合成率 = ""

End Sub




' ======================================================
Private Sub btn下げ1_Click()
On Error Resume Next
    Call Procテキストボックス値増減(textBB_重み, -1)
    
End Sub
' ======================================================
Private Sub btn下げ2_Click()
On Error Resume Next
    Call Procテキストボックス値増減(textRB_重み, -1)

End Sub
' ======================================================
Private Sub btn下げ3_Click()
On Error Resume Next
    Call Procテキストボックス値増減(text合算_重み, -1)

End Sub
' ======================================================
Private Sub btn上げ1_Click()
On Error Resume Next
    Call Procテキストボックス値増減(textBB_重み, 1)

End Sub
' ======================================================
Private Sub btn上げ2_Click()
On Error Resume Next
    Call Procテキストボックス値増減(textRB_重み, 1)

End Sub
' ======================================================
Private Sub btn上げ3_Click()
On Error Resume Next
    Call Procテキストボックス値増減(text合算_重み, 1)

End Sub

' ********************************************************************
Public Sub Procテキストボックス値増減(ctrl As Control, addValue As Double)
    Const methodName As String = "Procテキストボックス値増減"
    Dim currentValue As Double
On Error GoTo ErrHandler
    
'    ctrl.value = 0
    
    currentValue = CDbl(Nz(ctrl.value, 0))
    currentValue = currentValue + addValue
    If currentValue < 1 Then currentValue = 1
    ctrl.value = currentValue
    
    Call Form_Dirty(False)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub

' ======================================================
Private Sub btn下げ4_Click()
On Error Resume Next
    Call Procテキストボックス値増減(text設定調整値, -1)

End Sub

' ======================================================
Private Sub btn上げ4_Click()
On Error Resume Next
    Call Procテキストボックス値増減(text設定調整値, 1)
    If text設定調整値 > 5 Then text設定調整値 = 5

End Sub




