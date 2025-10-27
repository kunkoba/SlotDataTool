VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_パッケージ生成"
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
'    Call subタイトル.Form.Procタイトル更新("アプリ情報", E_Color_Blue)
        
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
'        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub
' ======================================================
Private Sub Procアプリ起動処理()
    Const methodName As String = "Procアプリ起動処理"
On Error Resume Next

    ' メイン処理

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub




' ======================================================
Private Sub btnバージョン更新_Click()
On Error Resume Next
    textバージョン = "v" & Format(Date, "dd_mm_yyyy")
    Me.Dirty = False

End Sub


' ======================================================
Private Sub btnパッケージ生成_Click()
    Const methodName As String = "btnパッケージ生成_Click"
On Error GoTo ErrHandler

    Call LogStart(methodName)
    
    If ShowConfirm("システム", "新バージョンパッケージを生成しますか？", vbYesNo) <> vbYes Then Exit Sub
    
    Call Procすべてのフォームを閉じる(Me.name)
    Call utlフォーム仕様統一(Me.name, True)
    Call Pacパッケージ作成
    
    Call utlフォーム仕様統一(Me.name, False)
    
    Call ShowToast("処理は正常に完了しました。", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub


