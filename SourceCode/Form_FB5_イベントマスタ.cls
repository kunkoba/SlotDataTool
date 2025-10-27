VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB5_イベントマスタ"
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
    Call subタイトル.Form.Procタイトル更新("日付マスタ編集", E_Color_Green)
    
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
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If

End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理
    Me!listイベント.Selected(1) = True
    Call listイベント_AfterUpdate

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnリスト編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB6_MstEventType

End Sub


' ======================================================
Private Sub listイベント_AfterUpdate()
On Error Resume Next
'    Me.Undo

    If IsNull(Me.listイベント) Then
        Me.FilterOn = False
    Else
        Me.filter = "[日付] = #" & Me.listイベント & "#"
        Me.FilterOn = True
    End If
    
End Sub

' ======================================================
Private Sub listイベント_Click()
On Error Resume Next
    listイベント.Requery
    
End Sub


' ======================================================
Private Sub listイベントタイプ_AfterUpdate()
On Error Resume Next
    Me.Dirty = False
    listイベント.Requery

End Sub

' ======================================================
Private Sub listイベントタイプ_Click()
On Error Resume Next
    listイベントタイプ.Requery

End Sub



' ======================================================
Private Sub btn一括編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB5_MstEvent_Bulk
    DoCmd.Close acForm, Me.name

End Sub



