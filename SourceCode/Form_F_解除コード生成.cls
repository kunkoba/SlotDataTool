VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_解除コード生成"
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
    
    text暗号キー = App_暗号キー
    
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
Private Sub btnコード生成_Click()
On Error Resume Next
    text暗号キー = Trim(text暗号キー)
    If Nz(text暗号キー, "") = "" Then Exit Sub
    
    text解除コード = Lic暗号キーから解除コード(text暗号キー)
    
End Sub


' ======================================================
Private Sub btnコピー_Click()
On Error Resume Next

    Call Procクリップボードへ転送(text解除コード)
    Call ShowToast("クリップボードへ転送されました。", E_Color_Blue)
    
End Sub


' ======================================================
Private Sub btnチェック_Click()
    Dim ret As Boolean
On Error Resume Next
    ret = Lic解除コードチェック(text暗号キー, text解除コード)
    
    lab判定結果.Visible = True
    lab判定結果.caption = IIf(ret, "チェックＯＫ", "チェックＮＧ")

End Sub





