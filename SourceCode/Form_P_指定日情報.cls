VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_P_指定日情報"
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
    
    Call btn今日_Click
    
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



' ******************************************************
' ******************************************************
Public Sub Set日付(sDate As String)
On Error Resume Next

    If IsDate(sDate) Then
        comb開始日 = CDate(sDate)
        Call comb開始日_AfterUpdate
    End If
    
End Sub

' ******************************************************
Public Function Get日付() As String
On Error Resume Next

    Get日付 = comb開始日.value
    
End Function




' ======================================================
Private Sub comb開始日_AfterUpdate()

    lab月旬.caption = 値変換(comb開始日, "月旬")
    lab月旬6.caption = 値変換(comb開始日, "月旬6")
    lab曜日.caption = 値変換(comb開始日, "曜日")
    lab日付末尾.caption = 値変換(comb開始日, "末尾")
    lab日.caption = 値変換(comb開始日, "日")
    
End Sub

' ======================================================
Private Sub btn開始日_前日_Click()
    日付加減算 comb開始日, -1
    Call comb開始日_AfterUpdate
    
End Sub

Private Sub btn開始日_翌日_Click()
    日付加減算 comb開始日, 1
    Call comb開始日_AfterUpdate
    
End Sub

' ******************************************************
' 日付移動
' ******************************************************
Private Sub 日付加減算(ByRef target As Control, ByVal offset As Long)
    If IsDate(target.value) Then
        target.value = DateAdd("d", offset, target.value)
'        Call Proc集計元フィルタリング処理
    End If
End Sub


' ======================================================
Private Sub btn今日_Click()
    comb開始日 = Date
    Call comb開始日_AfterUpdate

End Sub



