VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F12_伝達ダイアログ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' グローバル変数
Public 選択結果 As Integer



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
Private Sub Form_Open(Cancel As Integer)
On Error Resume Next
    Dim args() As String
    
    If Not IsNull(Me.OpenArgs) Then
        args = Split(Me.OpenArgs, vbTab)
        Me.lblタイトル.caption = "■" & args(0)
        Me.textメッセージ = args(1)
    End If
    
End Sub



' ======================================================
Private Sub btnコピー_Click()
    Const methodName As String = "btnコピー_Click"
On Error Resume Next
    
    Call Procクリップボードへ転送(textメッセージ)
    MsgBox "クリップボードへコピーしました。", vbInformation

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub

' ======================================================
Private Sub btnはい_Click()
On Error Resume Next
    選択結果 = 1
    Me.Visible = False
    
End Sub
' ======================================================
Private Sub btnいいえ_Click()
On Error Resume Next
    選択結果 = 2
    Me.Visible = False

End Sub

