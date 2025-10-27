VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F13_確認ダイアログ"
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
Private Sub Form_Open(Cancel As Integer)
On Error Resume Next
    Dim args() As String
    Dim mode As Integer
    
    If Not IsNull(Me.OpenArgs) Then
        args = Split(Me.OpenArgs, vbTab)
        Me.lblタイトル.caption = "■" & args(0)
        Me.lblメッセージ.caption = args(1)
        mode = CInt(args(2))
    End If
    
    Select Case mode
        Case vbYes, vbOKOnly
            '「はい」のみ
            btnいいえ.Visible = False
        Case Else
            '「はい、いいえ」
    End Select
End Sub



' ======================================================
Private Sub btnはい_Click()
On Error Resume Next
    選択結果 = vbYes
    Me.Visible = False
    
End Sub
' ======================================================
Private Sub btnいいえ_Click()
On Error Resume Next
    選択結果 = vbNo
    Me.Visible = False
    
End Sub

