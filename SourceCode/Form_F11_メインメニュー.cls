VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F11_メインメニュー"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_アプリ名 & " - " & App_バージョン
    Me.KeyPreview = True
    Call subタイトル.Form.Procタイトル更新("メインメニュー", E_Color_Black)

    Call LogOpen(Me.name)
    Call Procアプリ起動処理
    
    DoCmd.Close acForm, "F_初期フォーム"

End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
        Call btnアプリ終了_Click
        
    End If
    
End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)

    DoCmd.SetWarnings True
    DoCmd.Echo True
    Screen.MousePointer = 0    ' 0 = 通常の矢印
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    DoCmd.SetWarnings False
    
    text店舗名 = Get店舗名
    labライセンスマーク.caption = App_ライセンスマーク

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btn編集_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FB0_MstStore)

End Sub

' ======================================================
Private Sub text店舗名_Click()
On Error Resume Next
    text店舗名 = Get店舗名

End Sub

Private Sub 詳細_Click()
On Error Resume Next
    text店舗名 = Get店舗名
    labライセンスマーク.caption = App_ライセンスマーク
    Call subタイトル.Form.Procアプリ状態更新

End Sub


' ======================================================
Private Sub btnアプリ情報_Click()
On Error Resume Next
    DoCmd.OpenForm F02_AppInfo_Display

End Sub

' ======================================================
Private Sub btnアプリ終了_Click()
On Error Resume Next

    If ShowConfirm("終了前確認", "アプリケーションを閉じますか？") <> vbYes Then Exit Sub
    DoCmd.Quit

End Sub



' ======================================================
Private Sub btnA1_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FA1_ImportData)

End Sub
' ======================================================
Private Sub btnA2_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(F16_LinkManager)

End Sub
' ======================================================
Private Sub btnA3_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FC3_DataTimeline)

End Sub
' ======================================================
Private Sub btnB1_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FB1_MstMachine)

End Sub
' ======================================================
Private Sub btnB2_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FB3_MstMachineNum)

End Sub
' ======================================================
Private Sub btnB3_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FB5_MstEvent)

End Sub
' ======================================================
Private Sub btnC1_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FC1_DataView)

End Sub
' ======================================================
Private Sub btnC2_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FC2_DataSummary)

End Sub
' ======================================================
Private Sub btnC3_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FC4_DataClossSummary)

End Sub
' ======================================================
Private Sub btnD1_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FD1_GraphTrends)

End Sub
' ======================================================
Private Sub btnD2_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FD2_GraphCompare)

End Sub
' ======================================================
Private Sub btnD3_Click()
On Error Resume Next
    Call Proc画面遷移前認証チェック(FD3_GraphCustom)

End Sub



' **************************************************
' ライセンスチェックしてフォームを開く（汎用関数）
' **************************************************
Public Sub Proc画面遷移前認証チェック(formName As String)
    Const methodName As String = "Proc画面遷移前認証チェック"
    ' 宣言部
    Dim errMsg As String
On Error Resume Next

    ' メイン処理
    errMsg = Authアプリ利用チェック
    
    If errMsg = "" Then
        DoCmd.OpenForm formName
    Else
        Call ShowToast(errMsg, E_Color_Red)
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub




