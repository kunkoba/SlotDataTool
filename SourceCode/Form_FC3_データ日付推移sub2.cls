VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC3_データ日付推移sub2"
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
    
    Call 日付列セット(Date)
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


' ******************************************************
' 動的カラム設定
' ******************************************************
Public Sub 日付列セット(targetDate As Date)
    Const methodName As String = "日付列セット"
    Dim sql As String
    Dim rs As DAO.Recordset
    Dim i As Integer
    Dim d As Date
On Error GoTo ErrHandler
    
    ' クエリから直近15件の日付を取得（新しい順）
    sql = "SELECT DISTINCT TOP 30 日付 FROM L_イベントマスタ WHERE 日付 <= #" & targetDate & "# ORDER BY 日付 DESC;"
    Set rs = CurrentDb.OpenRecordset(sql)
    
    i = 1
    Do Until rs.EOF Or i > 15
        d = rs!日付
        
        ' text列にセット
        Me("text列" & i).ControlSource = Format(d, "yymmdd")
        
        ' lab列にキャプションをセット
        Me("lab列" & i).caption = Format(d, "m/d" & vbNewLine & "(aaa)")
        
        rs.MoveNext
        i = i + 1
    Loop

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(rs)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
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
' 各 Text 列クリックイベント
' ======================================================
Private Sub text列1_Click(): HandleTextClick 1: End Sub
Private Sub text列2_Click(): HandleTextClick 2: End Sub
Private Sub text列3_Click(): HandleTextClick 3: End Sub
Private Sub text列4_Click(): HandleTextClick 4: End Sub
Private Sub text列5_Click(): HandleTextClick 5: End Sub
Private Sub text列6_Click(): HandleTextClick 6: End Sub
Private Sub text列7_Click(): HandleTextClick 7: End Sub
Private Sub text列8_Click(): HandleTextClick 8: End Sub
Private Sub text列9_Click(): HandleTextClick 9: End Sub
Private Sub text列10_Click(): HandleTextClick 10: End Sub
Private Sub text列11_Click(): HandleTextClick 11: End Sub
Private Sub text列12_Click(): HandleTextClick 12: End Sub
Private Sub text列13_Click(): HandleTextClick 13: End Sub
Private Sub text列14_Click(): HandleTextClick 14: End Sub
Private Sub text列15_Click(): HandleTextClick 15: End Sub



' ======================================================
' 共通クリック処理
' ======================================================
Private Sub HandleTextClick(ByVal index As Integer)
    Const methodName As String = "HandleTextClick"
    Dim where As String
    Dim ctl As Control
On Error Resume Next
    
    ' text列X を参照
    Set ctl = Me("text列" & index)
    
    where = "台番号 = " & Convert条件値文字列(Me.台番号)
    where = where & " AND 日付 = " & Convert条件値文字列(Convert文字列から日付(ctl.ControlSource))
    
    DoCmd.OpenForm FC9_DataDateInfo, , , where

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub




' ======================================================
Private Sub btnグラフ1_Click()
On Error Resume Next
    DoCmd.OpenForm FD9_GraphCoinTrend, , , , , , Convert文字列から日付(text列1.ControlSource) & vbTab & Me.台番号

End Sub

' ======================================================
Private Sub btnグラフ2_Click()
On Error Resume Next
    DoCmd.OpenForm FD9_GraphCoinTrend, , , , , , Convert文字列から日付(text列1.ControlSource) & vbTab & Me.台番号

End Sub

