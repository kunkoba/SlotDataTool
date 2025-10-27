VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB5_イベントマスタ_一括編集"
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
    Call subタイトル.Form.Procタイトル更新("日付マスタ編集（一括）", E_Color_Green)
    
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
'    Me!listイベント.Selected(1) = True
'    Call listイベント_AfterUpdate

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnリスト編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB6_MstEventType

End Sub

'' ======================================================
'Private Sub listイベント_AfterUpdate()
'On Error Resume Next
''    Me.Undo
'
'    If IsNull(Me.listイベント) Then
'        Me.FilterOn = False
'    Else
'        Me.filter = "[日付] = #" & Me.listイベント & "#"
'        Me.FilterOn = True
'    End If
'
'End Sub

' ======================================================
Private Sub listイベント_Click()
On Error Resume Next
    listイベント.Requery
    
    If IsNull(Me.listイベント) Then
        Me.FilterOn = False
    Else
        Me.filter = "[日付] = #" & Me.listイベント & "#"
        Me.FilterOn = True
    End If
    
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
    DoCmd.OpenForm FB5_MstEvent
    DoCmd.Close acForm, Me.name

End Sub



' ======================================================
Private Sub btn更新_Click()
    Const methodName As String = "btn更新_Click"
    Dim nums As Variant
    Dim lists As String
    Dim sql As String
On Error GoTo ErrHandler

    Call LogStart(methodName)

    nums = Proc選択された要素を取得配列(listイベント, True)
    lists = Join(nums, ", ")
    If lists = "" Then Exit Sub
    
    ' SQL文字列を組み立て
    sql = "UPDATE M_イベントマスタ SET "
    sql = sql & "  イベントタイプID = " & Convert条件値文字列(listイベントタイプ)
    sql = sql & " ,イベント詳細 = " & Convert条件値文字列(textメモ)
    sql = sql & " WHERE 日付 in (" & lists & ")"
    
    ' 実行
    CurrentDb.Execute sql, dbFailOnError
        
    Call Procリストボックス複数選択解除(listイベント)
    
    listイベント.Requery
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）

End Sub



