VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD1_グラフ_傾向sub"
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
Public Sub グラフデータ差し替え処理_傾向(集計値 As String)
    Const methodName As String = "グラフデータ差し替え処理_傾向"
    ' 宣言部
On Error GoTo ErrHandler

    

    ' メイン処理
    Call グラフデータ再設定処理("月旬", 集計値)
    Call グラフデータ再設定処理("月旬6", 集計値)
    Call グラフデータ再設定処理("曜日", 集計値)
    Call グラフデータ再設定処理("日付末尾", 集計値)
    Call グラフデータ再設定処理("日", 集計値)
    
    Call グラフデータ再設定処理("イベントタイプ", 集計値)
    Call グラフデータ再設定処理("台タイプ", 集計値)
'    Call グラフデータ再設定処理("注目", 集計値)
    Call グラフデータ再設定処理("台番号末尾", 集計値)
    
    ' 設定判別用
    Call グラフ最小値設定_設定判別用(集計値)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub

' ======================================================
Private Sub グラフデータ再設定処理(項目列 As String, 集計列 As String)
    Const methodName As String = "グラフデータ再設定処理"
    ' 宣言部
On Error GoTo ErrHandler
    
    Select Case 項目列
        Case "月旬"
            グラフ月旬.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "月旬6"
            グラフ月旬6.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "曜日"
            グラフ曜日.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "日付末尾"
            グラフ日付末尾.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "日"
            グラフ日.RowSource = GenerateGraphSQL_A(項目列, 集計列)
            
        Case "イベントタイプ"
            グラフイベントタイプ.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "台タイプ"
            グラフ台タイプ.RowSource = GenerateGraphSQL_A(項目列, 集計列)
'        Case "注目"
'            グラフ注目.RowSource = GenerateGraphSQL_A(項目列, 集計列)
        Case "台番号末尾"
            グラフ台番号末尾.RowSource = GenerateGraphSQL_A(項目列, 集計列)
    End Select

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub


' ======================================================
Private Sub グラフ最小値設定_設定判別用(集計値 As String)
    Const methodName As String = "グラフ最小値設定_設定判別用"
    ' 宣言部
On Error GoTo ErrHandler

    Const minVal As Integer = 2
    Dim isOn As Boolean
    isOn = 集計値 = "設定判別"
    
    Call SetGraphMinScale(グラフ月旬, isOn, minVal)
    Call SetGraphMinScale(グラフ月旬6, isOn, minVal)
    Call SetGraphMinScale(グラフ曜日, isOn, minVal)
    Call SetGraphMinScale(グラフ日付末尾, isOn, minVal)
    Call SetGraphMinScale(グラフ日, isOn, minVal)
    
    Call SetGraphMinScale(グラフイベントタイプ, isOn, minVal)
    Call SetGraphMinScale(グラフ台タイプ, isOn, minVal)
'    Call SetGraphMinScale(グラフ注目, isOn, minVal)
    Call SetGraphMinScale(グラフ台番号末尾, isOn, minVal)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub

' ***********************************************
' グラフの目盛りを条件付きで設定
' ***********************************************
Public Sub SetGraphMinScale(ctl As Object, ByVal UseFixedMin As Boolean, Optional ByVal FixedMin As Double = 2)
    Const methodName As String = "SetGraphMinScale"
    ' 宣言部
On Error GoTo ErrHandler

    With ctl.Object
        If UseFixedMin Then
            .Axes(2).MinimumScale = FixedMin
            .Axes(2).MinimumScaleIsAuto = False
        Else
            .Axes(2).MinimumScaleIsAuto = True
        End If
    End With

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub

