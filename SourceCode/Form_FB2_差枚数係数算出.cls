VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FB2_差枚数係数算出"
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
        Me.Undo
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
Private Sub btn算出_Click()
    Const methodName As String = "btn算出_Click"
    ' 宣言部
On Error GoTo ErrHandler

    text差枚数係数.value = text獲得リプレイ.value / textリプレイ率.value
    text差枚数係数.value = text差枚数係数.value + text獲得チェリー.value / textチェリー率.value
    text差枚数係数.value = text差枚数係数.value + text獲得ぶどう.value / textぶどう率.value
    text差枚数係数.value = text差枚数係数.value + text獲得ピエロ.value / textピエロ率.value
    text差枚数係数.value = text差枚数係数.value + text獲得ベル.value / textベル率.value
    text差枚数係数.value = Round(text差枚数係数.value, 3)
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（メッセージ出力）
    
End Sub

' ======================================================
Private Sub btn保存_Click()
On Error Resume Next

    If Not Nz(text差枚数係数, "") = "" Then
        Forms(FB1_MstMachine)!text差枚数係数 = text差枚数係数
    End If
    
    DoCmd.Close acForm, Me.name

End Sub



