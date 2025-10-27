VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD2_グラフ_比較sub"
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
    
End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    ' メイン処理
    If KeyCode = vbKeyEscape Then
        Me.Undo
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub




' ======================================================
Public Sub Procグラフデータ差し替え処理_比較(集計区分 As Integer, 集計項目 As Variant, フィルタ値 As Variant)
    Const methodName As String = "Procグラフデータ差し替え処理_比較"
    ' 宣言部
On Error GoTo ErrHandler

    ' 集計区分（1…機種、2…台番号）
    Dim 項目列 As String
    項目列 = IIf(集計区分 = 1, "機種名", "台番号")
    集計項目 = Nz(集計項目, "")
    フィルタ値 = Nz(フィルタ値, "")

    グラフ比較_差枚数.RowSource = GenerateGraphSQL_B(項目列, CStr(集計項目), "差枚数", フィルタ値)
    グラフ比較_設定判別.RowSource = GenerateGraphSQL_B(項目列, CStr(集計項目), "設定判別", フィルタ値)
    グラフ比較_設定4.RowSource = GenerateGraphSQL_B(項目列, CStr(集計項目), "設定4以上", フィルタ値)
    グラフ比較_設定5.RowSource = GenerateGraphSQL_B(項目列, CStr(集計項目), "設定5以上", フィルタ値)
    グラフ比較_設定6.RowSource = GenerateGraphSQL_B(項目列, CStr(集計項目), "設定6", フィルタ値)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub



