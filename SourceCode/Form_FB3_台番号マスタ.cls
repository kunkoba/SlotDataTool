VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB3_台番号マスタ"
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
    Call subタイトル.Form.Procタイトル更新("台番号マスタ編集", E_Color_Green)
    
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
    Me!list台番号.Selected(1) = True
    Call list台番号_AfterUpdate
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub list台番号_AfterUpdate()
On Error Resume Next
'    Me.Undo
    
    If IsNull(Me.list台番号) Then
        Me.FilterOn = False
    Else
        Me.filter = "[台番号] = " & Me.list台番号 & ""
        Me.FilterOn = True
    End If
    
End Sub


' ======================================================
Private Sub btnリスト編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB4_MstMachineNumType

End Sub

' ======================================================
Private Sub btn島番号編集_Click()
On Error Resume Next
    DoCmd.OpenForm FB7_IslandNum
    DoCmd.Close acForm, Me.name

End Sub

' ======================================================
Private Sub list台番号_Click()
On Error Resume Next
    list台番号.Requery

End Sub



' ======================================================
Private Sub list台タイプ_AfterUpdate()
On Error Resume Next
    Me.Dirty = False
    list台番号.Requery
    
End Sub

' ======================================================
Private Sub list台タイプ_Click()
On Error Resume Next
    list台タイプ.Requery

End Sub


