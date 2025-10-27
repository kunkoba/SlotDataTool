VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FD9_グラフ_出玉推移"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim しきい日 As Date

' ======================================================
Private Sub Form_Load()
    Dim args() As String
On Error Resume Next
    Form.caption = App_アプリ名
    Me.KeyPreview = True
'    Call subタイトル.Form.Procタイトル更新("出玉推移グラフ", E_Color_DeepBlue)

    Call LogOpen(Me.name)
    Call Procアプリ起動処理

End Sub
' ======================================================
Private Sub Form_Close()
On Error Resume Next

    Call ErrorSave(Me.name, True) '閉じる前にエラーリセット
    Call LogClose(Me.name)
    Me.Undo

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
    Dim args
On Error GoTo ErrHandler

    ' メイン処理
    If Len(Nz(Me.OpenArgs, "")) > 0 Then
        ' vbTab で分割
        args = Split(Me.OpenArgs, vbTab)
        
        しきい日 = args(0)
        list台番号 = args(1)
    Else
        しきい日 = Date
        list台番号 = list台番号.ItemData(0)
    End If
    
    グラフ差枚数.RowSource = グラフ用SQL生成("差枚数", しきい日, list台番号)
    グラフ設定判別.RowSource = グラフ用SQL生成("設定判別", しきい日, list台番号)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭
    
End Sub



' ======================================================
Private Sub list台番号_AfterUpdate()
On Error Resume Next
    グラフ差枚数.RowSource = グラフ用SQL生成("差枚数", しきい日, list台番号)
    グラフ設定判別.RowSource = グラフ用SQL生成("設定判別", しきい日, list台番号)

End Sub


' ******************************************************
' グラフ用SQLを自動生成（日付・台番号指定あり）
' ******************************************************
Public Function グラフ用SQL生成(ByVal fieldName As String, _
                                ByVal targetDate As Date, _
                                ByVal num As Long) As String
    Const TARGET_TABLE As String = "D_集計元データ_最新_フィルタ"
    Dim sql As String
On Error Resume Next

    sql = " SELECT TOP 15 Format([日付],""mm/dd aaa"") AS 日付1, Avg(" & fieldName & ") AS 集計値 " & _
          " FROM " & TARGET_TABLE & _
          " WHERE 日付 <= #" & Format(targetDate, "yyyy/mm/dd") & "# " & _
          " AND 台番号 = " & num & _
          " GROUP BY Format([日付],""mm/dd aaa"") " & _
          " ORDER BY Format([日付],""mm/dd aaa"") DESC;"

    グラフ用SQL生成 = sql

End Function




