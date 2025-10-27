VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_オブジェクト開く"
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
'    Call subタイトル.Form.Procタイトル更新("XXXXXXX", 0)

    Call リストボックス取得
    
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
Sub リストボックス取得()
On Error Resume Next
    listテーブル.RowSource = utlオブジェクト名一覧取得(1, ",")
    listクエリ.RowSource = utlオブジェクト名一覧取得(2, ",")
    listフォーム.RowSource = utlオブジェクト名一覧取得(3, ",")
    
    Call Procリストボックス並べ替え(listテーブル)
    Call Procリストボックス並べ替え(listクエリ)
    Call Procリストボックス並べ替え(listフォーム)
    
End Sub


' ======================================================
Private Sub listテーブル_DblClick(Cancel As Integer)
On Error Resume Next
    DoCmd.OpenTable listテーブル

End Sub

' ======================================================
Private Sub listクエリ_DblClick(Cancel As Integer)
On Error Resume Next
    DoCmd.OpenQuery listクエリ, acViewDesign
    
End Sub

' ======================================================
Private Sub listフォーム_DblClick(Cancel As Integer)
On Error Resume Next
'    DoCmd.OpenForm listフォーム, acViewDesign
    DoCmd.OpenForm listフォーム

End Sub




