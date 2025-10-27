VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FA1_取り込み用"
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
    Call subタイトル.Form.Procタイトル更新("データ取り込み", E_Color_Blue)
    
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
    comb機種ID.value = 0
    text日付.value = Date - 1
    
    list形式 = list形式.ItemData(0)
    Call list形式_AfterUpdate
    
    flg最適化 = False

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub





' ======================================================
Private Sub btnデータ削除_Click()
On Error Resume Next
    DoCmd.OpenForm FA3_DataDelete
    
End Sub

' ======================================================
Private Sub btnデータクリア_Click()
    Dim queryList
On Error Resume Next

    queryList = Getクエリ一覧("Q0*")
    Call Procクエリリスト一括実行(queryList)
    subForm.Requery

    labファイル名.caption = ""
    
End Sub

' ======================================================
Private Sub btn貼り付け_Click()
    Const methodName As String = "btn貼り付け_Click"
    ' 宣言部
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    If subForm.SourceObject = "" Then Exit Sub
    
    Call btnデータクリア_Click
    
    subForm.SetFocus
    RunCommand acCmdSelectRecord    ' レコード全体選択
    
    DoCmd.SetWarnings False
    DoCmd.RunCommand acCmdPaste     ' ペースト

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    DoCmd.SetWarnings True
    
End Sub


' ======================================================
Private Sub btn前日_Click()
On Error Resume Next
    text日付.value = text日付.value - 1
    
    If IsNull(text日付) Then text日付 = Date
    
End Sub
' ======================================================
Private Sub btn翌日_Click()
On Error Resume Next
    text日付.value = text日付.value + 1
    
    If IsNull(text日付) Then text日付 = Date
    
End Sub
' ======================================================
Private Sub btn最小_Click()
On Error Resume Next
    
    If IsNull(text日付) Then text日付 = Date
    
    If Nz(comb機種ID, 0) = 0 Then Exit Sub
    Dim dat日付 As Date
    dat日付 = Nz(DMin("日付", "Check_積み上げ日_機種別", "機種ID = " & comb機種ID), Date)
    text日付 = dat日付 - 1

End Sub
' ======================================================
Private Sub btn最大_Click()
On Error Resume Next
    
    If IsNull(text日付) Then text日付 = Date
    
    If Nz(comb機種ID, 0) = 0 Then Exit Sub
    Dim dat日付 As Date
    dat日付 = Nz(DMax("日付", "Check_積み上げ日_機種別", "機種ID = " & comb機種ID), Date)
    text日付 = dat日付 + 1

End Sub


' ======================================================
Private Sub list形式_AfterUpdate()
On Error Resume Next
    Call Proclist形式_AfterUpdate

End Sub

' ======================================================
Public Sub Proclist形式_AfterUpdate()
On Error Resume Next
    Call btnデータクリア_Click
    
    If Nz(list形式, "") <> "" Then subForm.SourceObject = "テーブル." & list形式
    
End Sub


' ======================================================
Private Sub comb機種ID_AfterUpdate()
On Error Resume Next
    Call ProcComb機種ID_AfterUpdate

End Sub

' ======================================================
Public Sub ProcComb機種ID_AfterUpdate()
On Error Resume Next
    list積み上げ履歴.RowSource = Generate積み上げ履歴SQL(comb機種ID, "日付 DESC")

End Sub


' ======================================================
Private Sub btnCSV取り込み_Click()
    Const methodName As String = "btnCSV取り込み_Click"
    ' 宣言部
    Dim fileName As String
    Dim ret As Boolean
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    ' メイン処理
    fileName = Dialogファイル選択("", "CSVデータ", "*.csv")
    If fileName = "" Then Exit Sub
    
    ret = Proc取り込みCSV(fileName)
    
    If ret Then
'        Call ShowToast("正常終了", E_Color_Blue)
    Else
        Call ShowToast("CSVファイルの取り込みに失敗しました。", E_Color_Red)
    End If
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' ******************************************************
' CSV取り込み処理
' ******************************************************
Public Function Proc取り込みCSV(filePath As String) As Boolean
    Const methodName As String = "Proc取り込みCSV"
    ' 宣言部
    Dim tblName As String
On Error GoTo ErrHandler

    Proc取り込みCSV = False

    ' メイン処理
    tblName = Nz(list形式, "")
    If tblName = "" Or tblName = "1" Then Exit Function

    Call btnデータクリア_Click
    labファイル名.caption = GetFileName(filePath)

    ' CSVファイル取り込み
    If dataインポートCSV(filePath, tblName) Then
        subForm.Requery
        text日付 = Convert文字列から日付(Left(GetFileName(filePath), 8))
        
        Proc取り込みCSV = True
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function



' ======================================================
Private Sub btnCSV一括取り込み_Click()
On Error Resume Next
    comb機種ID.SetFocus
    DoCmd.OpenForm FA2_ImportDataCSV

End Sub




' ======================================================
Private Sub btn積み上げ_Click()
    Const methodName As String = "btn積み上げ_Click"
    Dim cnt As Long
On Error Resume Next
    
    Call LogStart(methodName)
    
    ' メイン処理
    If ShowConfirm("終了前確認", "積み上げ処理を実行しますか？") <> vbYes Then Exit Sub
        
    If Proc積み上げ処理 Then
        cnt = DCount("*", "T_取り込み用")
        Call ShowToast(cnt & " 件が取り込まれました。", E_Color_Blue)
    Else
'        Call ShowToast("データの積み上げは異常終了しました。", E_Color_Red)
    End If

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' ======================================================
Public Function Proc積み上げ処理() As Boolean
    Const methodName As String = "Proc積み上げ処理"
    ' 宣言部
    Dim msg As String
    Dim queryList As Variant
    Dim cnt As Integer
On Error GoTo ErrHandler

    Proc積み上げ処理 = False
   
    comb機種ID.SetFocus
    
    'チェック
    msg = Check積み上げ処理前
    If msg <> "" Then
'        MsgBox msg, vbExclamation
        Call ShowToast(msg, E_Color_Red)
        GoTo ErrHandler
    End If
    
    'クエリ実行1
    Call Procクエリリスト一括実行(Getクエリ一覧("Q1*"))
    
    Call Update積み上げ用
    
    '正常値チェック
    If Not Checkデータ正常値判定 Then
        Call ShowToast("積み上げ予定データが異常値です。　積み上げ処理を中断します。", E_Color_Red)
        GoTo ErrHandler
    End If
    
    'クエリ実行2
    Call Procクエリリスト一括実行(Getクエリ一覧("Q2*"))
    
    'クエリ実行3
    Call Procクエリリスト一括実行(Getクエリ一覧("Q3*"))
    
    'クエリ実行4
    Call Procクエリリスト一括実行(Getクエリ一覧("Q4*"))
    
    'クエリ実行0（クリア）
    Call Procクエリリスト一括実行(Getクエリ一覧("Q0*"))
    
    subForm.Requery
    
    list積み上げ履歴.Requery
    
    Proc積み上げ処理 = True
    
    flg最適化 = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function

' ******************************************************
Public Sub Update積み上げ用()
    Const methodName As String = "Update積み上げ用"
    Dim sql As String
On Error Resume Next
    
    ' SQL文字列を組み立て
    sql = "UPDATE T_積み上げ用 SET 機種ID = " & comb機種ID.value & _
          ", 日付 = #" & Format(text日付.value, "yyyy/mm/dd") & "#;"

    ' 実行
    CurrentDb.Execute sql, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ******************************************************
Public Function Check積み上げ処理前() As String
    Const methodName As String = "Check積み上げ処理前"
    Dim msg As String
On Error Resume Next
    
    'チェック
    If Nz(comb機種ID, 0) = 0 Then msg = "機種が選択されていません！"
    If Nz(text日付, "") = "" Then msg = "日付が入力されていません！"
'    If DCount("*", list形式.value) = 0 Then msg = "積み上げ用データが用意されていません！"
    
    Check積み上げ処理前 = msg
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



' ******************************************************
Public Function Checkデータ正常値判定() As Boolean
    Const methodName As String = "Checkデータ正常値判定"
    ' 宣言部
    Dim rs As DAO.Recordset
    Dim isValid As Boolean
On Error GoTo ErrHandler
    
    ' メイン処理
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM Check_積み上げ前_正常値", dbOpenSnapshot)

    If rs.EOF Then
        GoTo ErrHandler
    End If

    isValid = True

    ' ▼▼ 適正範囲チェック ▼▼
    If rs!ゲーム数平均 < 500 Then isValid = False
    If rs!BB数平均 < 0 Or rs!BB数平均 > 200 Then isValid = False
    If rs!RB数平均 < 0 Or rs!RB数平均 > 200 Then isValid = False
    ' ↑ 他のフィールドもあれば追加してください

    ' ▼▼ 結果メッセージ ▼▼
    If isValid Then
        Checkデータ正常値判定 = True
    Else
        Checkデータ正常値判定 = False
        Call ShowToast("正常値範囲外のデータがあります。", 2)
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(rs)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function




' ======================================================
Private Sub btn降順データ日付_Click()
On Error Resume Next
    list積み上げ履歴.RowSource = Generate積み上げ履歴SQL(comb機種ID, "日付 DESC")
    
End Sub
' ======================================================
Private Sub btn昇順データ日付_Click()
On Error Resume Next
    list積み上げ履歴.RowSource = Generate積み上げ履歴SQL(comb機種ID, "日付")

End Sub
' ======================================================
Public Sub btn降順データ積み上げ日_Click()
On Error Resume Next
    list積み上げ履歴.RowSource = Generate積み上げ履歴SQL(comb機種ID, "積み上げ日時 DESC")

End Sub
' ======================================================
Private Sub btn昇順データ積み上げ日_Click()
On Error Resume Next
    list積み上げ履歴.RowSource = Generate積み上げ履歴SQL(comb機種ID, "積み上げ日時")

End Sub




' ======================================================
Private Sub btn解説_Click()
    Dim msg As String
On Error Resume Next

    msg = ""
    msg = JoinText(msg, "　① 「データ形式リスト」からデータ形式を選択します。", True)
    msg = JoinText(msg, "　　　形式によって列が違います。★の列は切り捨てられます。", True)
    msg = JoinText(msg, "　② 取り込んだデータの内容を確認します。", True)
    msg = JoinText(msg, "　③ 「機種リスト」から機種を選択します。", True)
    msg = JoinText(msg, "　④ 取り込みデータの日付（いつのデータか）を選択します。", True)
    msg = JoinText(msg, "　⑤ すべて問題なければ「データ積み上げボタン」を押します。", True)
    msg = JoinText(msg, "　⑥ 問題なく積み上げられたら、「積み上げ履歴」に履歴が残ります。", True)
    msg = JoinText(msg, "　　　間違えて積み上げてしまった場合は、「データ削除ボタン」から削除できます。", True)

    Call ShowInfomation("【データ積み上げ手順】", msg)

End Sub




