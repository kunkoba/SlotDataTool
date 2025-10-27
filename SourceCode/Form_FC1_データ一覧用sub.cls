VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC1_データ一覧用sub"
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
    
    Call btn解除_Click
    
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
Private Sub btn解除_Click()
On Error Resume Next
    With Me
        .filter = ""
        .FilterOn = False
        .OrderBy = ""
        .OrderByOn = False
    End With

End Sub

' ======================================================
Private Sub btn更新_Click()
    Me.Requery
    
End Sub




'' ======================================================
'Private Sub btn出玉傾向_Click()
'On Error Resume Next
'    Dim しきい値 As String
'    しきい値 = InputBox("差枚数と機械割との差分を表記しています。" & vbNewLine & "上振れ・下振れを検知するためのしきい値を入力してください。" & vbNewLine & "（デフォルト：1000）", "しきい値の入力")
'    If しきい値 = "" Then Exit Sub
'
'    text出玉傾向.ControlSource = "=出玉傾向([差枚数], [機械割差枚数], " & しきい値 & ")"
'
'End Sub


' ======================================================
Private Sub btn日情報_Click()
    Dim where As String
On Error Resume Next
    where = "台番号 = " & Convert条件値文字列(Me.台番号)
    where = where & " AND 日付 = " & Convert条件値文字列(Me.日付)
    
    DoCmd.OpenForm FC9_DataDateInfo, , , where
    
End Sub


