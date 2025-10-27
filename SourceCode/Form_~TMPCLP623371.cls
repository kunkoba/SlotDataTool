VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_~TMPCLP623371"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Compare Database
'Option Explicit
'
'
'
'' ======================================================
'Private Sub Form_Current()
''    Me.subForm1.SourceObject = "_F_グラフA_傾向sub"
'
'End Sub
'
'' ======================================================
'Private Sub Form_Load()
'
'    ' メイン処理
'    Form.caption = App_アプリ名
'    Me.KeyPreview = True
'    Call subタイトル.Form.Procタイトル更新("XXXXXXX", 0)
'
'    list機種マスタ = Null
''    Call Update台番号Query(0)
''    list台番号マスタ.RowSource = list台番号マスタ.RowSource
'
'ErrHandler:
'
'End Sub
'' ======================================================
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    ' 宣言部
'On Error Resume Next
'
'    ' メイン処理
'    If KeyCode = vbKeyEscape Then
'        Me.Undo
'        DoCmd.Close acForm, Me.name
'        DoCmd.Close acForm, Me.Parent.name
'    End If
'
'End Sub
'' ======================================================
'Private Sub Form_Close()
'    Call Update集計元Query(Null, Null, Null, Null)
'
'End Sub
'
'
'
'
'' ======================================================
'Private Sub btn閉じる_Click()
'    DoCmd.Close acForm, Me.name
'
'End Sub
'
'
'' ======================================================
'Private Sub btnクリア機種_Click()
'    list機種マスタ = Null
'    Call Update台番号Query(0)
''    list台番号マスタ.RowSource = list台番号マスタ.RowSource
'
'    Call Proc集計元フィルタリング処理
'
'End Sub
'
'' ======================================================
'Private Sub btnクリア日付_Click()
'    comb開始日 = Null
'
'    Call Proc集計元フィルタリング処理
'
'End Sub
'
'
'' ======================================================
'Private Sub comb開始日_AfterUpdate()
'    Call Proc集計元フィルタリング処理
'
'End Sub
'' ======================================================
'Private Sub btn開始日_前日_Click()
'    日付加減算 comb開始日, -1
'End Sub
'
'Private Sub btn開始日_翌日_Click()
'    日付加減算 comb開始日, 1
'End Sub
'
'' ******************************************************
'' 日付移動
'' ******************************************************
'Private Sub 日付加減算(ByRef target As Control, ByVal offset As Long)
'    If IsDate(target.value) Then
'        target.value = DateAdd("d", offset, target.value)
'        Call Proc集計元フィルタリング処理
'    End If
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
''    Call Update台番号Query(list機種マスタ)
''    list台番号マスタ.RowSource = list台番号マスタ.RowSource
''    list台番号マスタ = Null
'
'    Call Proc集計元フィルタリング処理
'
'ErrHandler:
'
'End Sub
'
''' ======================================================
''Private Sub list台番号マスタ_AfterUpdate()
''    Call Proc集計元フィルタリング処理
''
''End Sub
'
'' ======================================================
'Private Sub list値リスト_AfterUpdate()
'    Call グラフ差し替え処理
'
'End Sub
'
'
'
'' ******************************************************
'Private Sub Proc集計元フィルタリング処理()
'
'    ' 宣言部
'On Error GoTo ErrHandler
'
'    ' メイン処理
'    Call Update集計元Query(list機種マスタ, Null, comb開始日, Null)
'
'    DoCmd.Echo False
'    Me.subForm1.Form.Recalc
'
'ErrHandler:
'
'End Sub
'
'
'' ******************************************************
'Private Sub Procグラフ差し替え処理()
'    ' 宣言部
'On Error GoTo ErrHandler
'
'    ' メイン処理
''    Call Me.subForm1.Form.グラフデータ差し替え処理_傾向(list値リスト)
'
'ErrHandler:
'    ' エラー
'
'End Sub
'
