VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_~TMPCLP173701"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'Option Explicit
'
'
'' ======================================================
'Private Sub Form_Load()
'On Error Resume Next
'    Form.caption = App_アプリ名
'    Me.KeyPreview = True
''    Call subタイトル.Form.Procタイトル更新("データ取り込み", 1)
'
'
'
'
'    Me.subForm1.SourceObject = Me.name & "sub"
'
'    list機種マスタ = list機種マスタ.ItemData(1)
'    Call list機種マスタ_AfterUpdate
'    list台番号マスタ = list台番号マスタ.ItemData(1)
'    Call list台番号マスタ_AfterUpdate
'
'
'End Sub
'' ======================================================
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'On Error Resume Next
'    If KeyCode = vbKeyEscape Then
'        Me.Undo
'        DoCmd.Close acForm, Me.name
'        DoCmd.Close acForm, Me.Parent.name
'    End If
'
'    Select Case KeyCode
'        Case vbKeyRight
'            Call btn指定日_翌日_Click
'        Case vbKeyLeft
'            Call btn指定日_前日_Click
'    End Select
'
'
'End Sub
'' ======================================================
'Private Sub Form_Close()
'On Error Resume Next
'    Call Update集計元Query(Null, Null, Null, Null)
'
'
'End Sub
'
'
'
'' ======================================================
'Private Sub btn閉じる_Click()
'On Error Resume Next
'    DoCmd.Close acForm, Me.name
'
'
'End Sub
'
'
'' ======================================================
'Private Sub btnクリア機種_Click()
'On Error Resume Next
'    list機種マスタ = Null
'    Call Update台番号Query(0)
'    list台番号マスタ.RowSource = list台番号マスタ.RowSource
'
'    Call Proc集計元フィルタリング処理
'
'
'End Sub
'
'
'
'
'' ======================================================
'Private Sub btn今日_Click()
'On Error Resume Next
'    text今日 = Date
'    Call Proc集計元フィルタリング処理
'
'
'End Sub
'' ======================================================
'Private Sub btn指定日_前日_Click()
'On Error Resume Next
'    日付加減算 text今日, -1
'
'
'End Sub
'' ======================================================
'Private Sub btn指定日_翌日_Click()
'On Error Resume Next
'    日付加減算 text今日, 1
'
'
'End Sub
'' ======================================================
'Private Sub btn指定日_10日前_Click()
'On Error Resume Next
'    日付加減算 text今日, -10
'    Call Proc集計元フィルタリング処理
'
'
'End Sub
'' ******************************************************
'' 日付移動
'' ******************************************************
'Private Sub 日付加減算(ByRef target As Control, ByVal offset As Long)
'    ' 宣言部
'On Error GoTo ErrHandler
'
'    ' メイン処理
'    If IsDate(target.value) Then
'        target.value = DateAdd("d", offset, target.value)
'        Call Proc集計元フィルタリング処理
'    End If
'
'ErrHandler:
'    ' エラー
'
'End Sub
'
'
'
'
'' ======================================================
'Private Sub list機種マスタ_AfterUpdate()
'    ' 宣言部
'On Error GoTo ErrHandler
'
'    ' メイン処理
'    Call Update台番号Query(list機種マスタ)
'    list台番号マスタ.RowSource = list台番号マスタ.RowSource
'    list台番号マスタ = Null
'
'    Call Proc集計元フィルタリング処理
'
'ErrHandler:
'    ' エラー
'
'End Sub
'
'' ======================================================
'Private Sub list台番号マスタ_AfterUpdate()
'On Error Resume Next
'
'    ' メイン処理
'    Call Proc集計元フィルタリング処理
'
'
'End Sub
'
'
'
'
'' ******************************************************
'Private Sub Proc集計元フィルタリング処理()
'    ' 宣言部
'On Error GoTo ErrHandler
'
'    ' メイン処理
'    Call Update集計元Query(list機種マスタ, list台番号マスタ, Null, text今日)
'
'    Me.subForm1.SourceObject = Me.subForm1.SourceObject
'
'ErrHandler:
'    ' エラー
'
'End Sub
'
