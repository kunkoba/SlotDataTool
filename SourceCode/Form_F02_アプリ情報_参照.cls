VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F02_アプリ情報_参照"
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
    Call subタイトル.Form.Procタイトル更新("アプリ情報", E_Color_Blue)
    
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

    ' 暗号キーセット
    If IsNull(App_アプリ解除日) Then
        textアプリ解除日.Visible = False
        text有効期限.Visible = True
    Else
        textアプリ解除日.Visible = True
        text有効期限.Visible = False
    End If

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub text公開サイト_Click()
On Error Resume Next
    If Not IsNull(Me.text公開サイト) Then
        Application.FollowHyperlink Me.text公開サイト
    End If

End Sub

' ======================================================
Private Sub text問合せ先_Click()
On Error Resume Next
    If Not IsNull(Me.text問合せ先) Then
        Application.FollowHyperlink Me.text問合せ先
    End If

End Sub




' ======================================================
Private Sub btn解除依頼文_Click()
    Dim strWord As String
On Error Resume Next

    ' --- 開発者に送信する匿名依頼メール本文を組み立て ---
    strWord = JoinText(strWord, "◆◆◆　" & App_アプリ名 & " 制限解除依頼　◆◆◆", True)
    strWord = JoinText(strWord, "", True)
    strWord = JoinText(strWord, App_アプリ名 & "の制限解除をお願いします！", True)
    strWord = JoinText(strWord, "", True)
    strWord = JoinText(strWord, "==============================================", True)
    strWord = JoinText(strWord, "暗号キー： " & App_暗号キー, True)
    strWord = JoinText(strWord, "==============================================", True)
    strWord = JoinText(strWord, "", True)
    strWord = JoinText(strWord, "【ユーザ情報】", True)
    strWord = JoinText(strWord, "　メールアドレス：○○○○○○○○○", True)
    strWord = JoinText(strWord, "　お客様情報　　：○○○○○○○○○", True)
    strWord = JoinText(strWord, "", True)
'    strWord = JoinText(strWord, "【システム情報】", True)
'    strWord = JoinText(strWord, "　パソコン名：" & Getパソコン情報()(0), True)
'    strWord = JoinText(strWord, "　ユーザ名　：" & Getパソコン情報()(1), True)
'    strWord = JoinText(strWord, "　IPアドレス：" & Getパソコン情報()(2), True)
'    strWord = JoinText(strWord, "", True)
    strWord = JoinText(strWord, "よろしくお願いいたします。", True)
    strWord = JoinText(strWord, "", True)

    Call ShowInfomation("アプリ制限解除依頼文", strWord)
    
End Sub


' ======================================================
Private Sub btnアプリ解除_Click()
    Const methodName As String = "btnアプリ解除_Click"
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    If Trim(Nz(text解除コード, "")) = "" Then Exit Sub
    
    If Lic解除コードチェック(Me.暗号キー, text解除コード) Then
    
        Call Setアプリ解除処理(text解除コード)
        
        textアプリ解除日.Visible = True
        text有効期限.Visible = False
        textアプリ解除日.SetFocus
        
        Call ShowToast("正しい解除コードが入力されました！", E_Color_Blue)
        
        Call subタイトル.Form.Procアプリ状態更新
        
    Else
        Call ShowToast("正しい解除コードを入力してください！", E_Color_Red)
        
    End If

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub



' ======================================================
'■ S_アプリ設定の指定項目を設定する
' ======================================================
Public Function Setアプリ解除処理(str解除コード As String) As Boolean
    Const methodName As String = "Setアプリ解除処理"
    ' 宣言部
    Dim strSQL As String
    Dim key As String
On Error GoTo ErrHandler

    Setアプリ解除処理 = False
        
    ' SQL作成
    strSQL = "UPDATE S_アプリ設定 " & _
             "SET " & _
             " アプリ解除日 = #" & Date & "#" & _
             ",解除コード = '" & str解除コード & "'"
    
    ' SQL実行
    CurrentDb.Execute strSQL, dbFailOnError
    
    Setアプリ解除処理 = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


