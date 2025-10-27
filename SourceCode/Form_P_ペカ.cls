VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_P_ペカ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Dim ボタン押下数 As Integer


' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_アプリ名
    Me.KeyPreview = True
    
    Me.labあたり.Visible = False
'    Me.labはずれ.Visible = True

    ボタン押下数 = 3
    
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
Private Sub btn1_Click()
On Error Resume Next
    ボタン押下数 = ボタン押下数 - 1
    btn枠.SetFocus
    btn1.Visible = False
    Call 抽選開始
    
End Sub

' ======================================================
Private Sub btn2_Click()
On Error Resume Next
    ボタン押下数 = ボタン押下数 - 1
    btn枠.SetFocus
    btn2.Visible = False
    Call 抽選開始

End Sub

' ======================================================
Private Sub btn3_Click()
On Error Resume Next
    ボタン押下数 = ボタン押下数 - 1
    btn枠.SetFocus
    btn3.Visible = False
    Call 抽選開始

End Sub


' *******************************************************
Private Sub 抽選開始()
    Const あたり確率 As Integer = 10
    Dim 結果 As Integer
    Dim 表示文字 As String
On Error Resume Next

    If ボタン押下数 > 0 Then Exit Sub
    
    結果 = Int(Rnd() * あたり確率)
        
    If 結果 = 0 Then
        Me.labあたり.Visible = True
        
        表示文字 = Procおみくじ抽選
        labおみくじ結果.caption = 表示文字
        Debug.Print Now, "あたり", 表示文字
        
        DoEvents
        Sleep 5000
        
        Me.labあたり.Visible = False
        labおみくじ結果.caption = ""
        
    Else
'        Debug.Print Now, "はずれ", 結果
        
    End If
    
    ボタン押下数 = 3
    btn1.Visible = True
    btn2.Visible = True
    btn3.Visible = True
    
End Sub


' *******************************************************
Private Function Procおみくじ抽選() As String
    Dim messages(0 To 9) As String
    Dim index As Integer
On Error Resume Next

    ' メッセージをセット
    messages(0) = "今日はいいことあるかも！"
    messages(1) = "運気上昇中！"
    messages(2) = "目の前のチャンスを逃すな！"
    messages(3) = "信じる者は救われる！"
    messages(4) = "ひと休みも大事だよ"
    messages(5) = "大胆に行動してみよう！"
    messages(6) = "努力が実を結びます"
    messages(7) = "ラッキーアイテムは青！"
    messages(8) = "何か始めるのに最適な日！"
    messages(9) = "笑顔を忘れずにね！"

    ' ランダムインデックスを選ぶ（0～9）
    Randomize ' ← 忘れずに初期化
    index = Int(Rnd() * 10)

    ' メッセージ返却
    Procおみくじ抽選 = messages(index)
    
End Function




