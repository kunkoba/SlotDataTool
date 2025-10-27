VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_~TMPCLP82581"
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
'
'    Call チャート反映
'    Call comb機種ID_AfterUpdate
'    Call comb一括設定
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
'End Sub
'
'
'' ******************************************************
'Private Sub チャート反映(Optional 項目1 As String = "")
'    Call Procグラフ用クエリ動的変更(comb機種ID, comb台番号, text開始日, text終了日)
'
'    With subForm1
'        Call 設定動的チャート(.Form, 項目1, "曜日", comb値, comb演算, Nz(comb曜日, ""))
'        .Form.Procチャート用SQL更新
'    End With
'
'    With subForm2
'        Call 設定動的チャート(.Form, 項目1, "月旬", comb値, comb演算, Nz(comb月旬, ""))
'        .Form.Procチャート用SQL更新
'    End With
'
'    With subForm3
'        Call 設定動的チャート(.Form, 項目1, "月旬6", comb値, comb演算, Nz(comb月旬6, ""))
'        .Form.Procチャート用SQL更新
'    End With
'
'End Sub
'
'' ******************************************************
'' 汎用：サブフォーム内のコンボボックスに値をセットする
'' ******************************************************
'Public Sub 設定動的チャート(sf As Form, field1 As String, field2 As String, sumField As String, aggFunc As String, filter2 As String)
'    ' サブフォーム内のコントロールに値をセット
'    With sf
'        .comb項目1.value = field1
'        .comb項目2.value = field2
'        .comb値.value = sumField
'        .comb演算.value = aggFunc
'        .combフィルタ2.RowSource = ""
'        .combフィルタ2.value = 調整(filter2)
'    End With
'
'End Sub
'
'
'' ======================================================
'Private Sub btn次へ_Click()
'    If Nz(comb台番号, "") = "" Then
'        comb台番号.value = comb台番号.ItemData(1)
'    Else
'        Call changeCombbox(comb台番号, 1)
'    End If
'
'End Sub
'
'Private Sub btn前へ_Click()
'    If Nz(comb台番号, "") = "" Then
'        comb台番号.value = comb台番号.ItemData(1)
'    Else
'        Call changeCombbox(comb台番号, -1)
'    End If
'
'End Sub
'
'' ======================================================
'Private Sub comb演算_AfterUpdate()
'    If Nz(comb演算, "") = "" Then comb演算 = "avg"
'
'End Sub
'
'Private Sub comb値_AfterUpdate()
'    If Nz(comb値, "") = "" Then comb値 = "差枚数"
'
'End Sub
'
'
'' ======================================================
'Private Sub comb期間_AfterUpdate()
'    text開始日 = comb期間.Column(1)
'    text終了日 = comb期間.Column(2)
'
'End Sub
'
'Private Sub comb機種ID_AfterUpdate()
'    comb台番号 = Null
'    comb台番号.Requery
'
'End Sub
'
'
'' ======================================================
'Private Sub btn期間クリア_Click()
'    comb期間 = Null
'    text開始日 = Null
'    text終了日 = Null
'
'End Sub
'
'Private Sub btn機種_Click()
'    comb機種ID = Null
'    comb台番号 = Null
'
'End Sub
'
'Private Sub btn日付クリア_Click()
'    text日付.SetFocus
'    text日付.text = ""
'    Call text日付_Change
'
'End Sub
'
'
'' ======================================================
'Private Sub text日付_Change()
'    Dim strDate As String
'    strDate = text日付.text
'
'    comb曜日 = 調整(値変換(strDate, "曜日"))
'    comb月旬 = 調整(値変換(strDate, "月旬"))
'    comb月旬6 = 調整(値変換(strDate, "月旬6"))
'    comb日 = 調整(値変換(strDate, "日"))
'    comb日付末尾 = 調整(値変換(strDate, "末尾"))
'
'    Call comb一括設定
'
'End Sub
'
'' ******************************************************
'Private Sub comb一括設定()
'    comb曜日.RowSource = Procリスト用SQL生成("曜日")
'    comb月旬.RowSource = Procリスト用SQL生成("月旬")
'    comb月旬6.RowSource = Procリスト用SQL生成("月旬6")
'    comb日.RowSource = Procリスト用SQL生成("日")
'    comb日付末尾.RowSource = Procリスト用SQL生成("日付末尾")
'
'End Sub
'
'' ******************************************************
'Private Function 調整(val As String)
'    If val = "-1" Then
'        調整 = ""
'        Exit Function
'    End If
'    調整 = val
'
'End Function
'
'
'
'' ======================================================
'Private Sub btn通常_Click()
''    Call btn日付クリア_Click
'    Call チャート反映
'
'End Sub
'
'' ======================================================
'Private Sub btn時系列_Click()
'    Call チャート反映("年月")
'
'End Sub
