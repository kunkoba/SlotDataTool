VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_簡易操作ォーム"
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
Private Sub btnすべて閉じる_Click()
On Error Resume Next
    If ShowConfirm("システム", "すべての画面を閉じますか？", vbYesNo) <> vbYes Then Exit Sub
    Call Procすべてのフォームを閉じる(Me.name)

End Sub
' ======================================================
Private Sub btnオブジェクト開く_Click()
On Error Resume Next
    DoCmd.OpenForm "F_オブジェクト開く"

End Sub
' ======================================================
Private Sub btnポップアップ_オン_Click()
On Error Resume Next
    If ShowConfirm("システム", "ポップアップをオンにしますか？", vbYesNo) <> vbYes Then Exit Sub
    Call Procすべてのフォームを閉じる(Me.name)
    Call utlフォーム仕様統一(Me.name, True)

End Sub
' ======================================================
Private Sub btnポップアップ_オフ_Click()
On Error Resume Next
    If ShowConfirm("システム", "ポップアップをオフにしますか？", vbYesNo) <> vbYes Then Exit Sub
    Call Procすべてのフォームを閉じる(Me.name)
    Call utlフォーム仕様統一(Me.name, False)
    
End Sub
' ======================================================
Private Sub btn初期フォーム_Click()
On Error Resume Next
    Call Procアプリ起動初期処理(False)

End Sub
' ======================================================
Private Sub btnソース出力_Click()
On Error Resume Next
    If ShowConfirm("処理前確認", "ソースを出力します。　よろしいですか？") <> vbYes Then Exit Sub
    Call utlソースコード出力
    
End Sub
' ======================================================
Private Sub btnアプリ情報入力_Click()
On Error Resume Next
    DoCmd.OpenForm F02_AppInfo_Input

End Sub
' ======================================================
Private Sub btn解除コード生成_Click()
On Error Resume Next
    DoCmd.OpenForm "F_解除コード生成"

End Sub
' ======================================================
Private Sub btnパッケージ_Click()
On Error Resume Next
'    If ShowConfirm("システム", "新バージョンパッケージを生成しますか？", vbYesNo) <> vbYes Then Exit Sub
'    Call Pacパッケージ作成
    DoCmd.OpenForm "F_パッケージ生成"

End Sub



