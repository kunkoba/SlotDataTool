Attribute VB_Name = "SQL2"
Option Compare Database
Option Explicit


' ***********************************************
' 動的クエリ生成（集計クエリ）
' ***********************************************
Public Function Proc動的集計クエリSQL生成(行1 As Variant, 行2 As Variant, 行3 As Variant, where句 As String) As String
    Const methodName As String = "Proc動的集計クエリSQL生成"
    ' 宣言部
    Dim sql As String
    Dim groupFields As String
    Dim selectFields As String
    Dim f1 As String, f2 As String, f3 As String
On Error GoTo ErrHandler
    
    ' メイン処理
    ' --- Null/空白チェック ---
    If IsNull(行1) Or Trim(行1 & "") = "" Then
        f1 = """"""
    Else
        f1 = 行1
    End If
    
    If IsNull(行2) Or Trim(行2 & "") = "" Then
        f2 = """"""
    Else
        f2 = 行2
    End If
    
    If IsNull(行3) Or Trim(行3 & "") = "" Then
        f3 = """"""
    Else
        f3 = 行3
    End If
    
    ' --- SELECT句 ---
    selectFields = f1 & " AS 項目1, " & _
                   f2 & " AS 項目2, " & _
                   f3 & " AS 項目3, " & _
                   "Count(機種ID) AS データ件数, " & _
                   "Sum(収支) AS 収支の合計, " & _
                   "Sum(ゲーム数) AS ゲーム数の合計, " & _
                   "Sum(BB数) AS BB数の合計, " & _
                   "Sum(RB数) AS RB数の合計, " & _
                   "Sum(差枚数) AS 差枚数の合計, " & _
                   "Avg(差枚数) AS 差枚数の平均, " & _
                   "Avg(設定判別) AS 設定判別の平均, " & _
                   "Avg(設定4以上) AS 設定4投入率, " & _
                   "Avg(設定5以上) AS 設定5投入率, " & _
                   "Avg(設定6) AS 設定6投入率"
    
    ' --- GROUP BY句 ---
    groupFields = f1 & ", " & f2 & ", " & f3

    ' --- SQL組み立て ---
    sql = " SELECT " & selectFields & _
          " FROM D_集計元データ"

    If Trim(where句 & "") <> "" Then
        sql = sql & " WHERE " & where句 & vbCrLf
    End If

    sql = sql & " GROUP BY " & groupFields & _
                " ORDER BY " & groupFields

    ' ログ出力
    Call LogDebug("Proc動的集計クエリSQL生成", sql)
    Proc動的集計クエリSQL生成 = sql

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' 動的クエリ生成（クロス集計クエリ）
' ***********************************************
Public Function Proc動的クロス集計SQL生成( _
                    行1 As String, 行2 As String, _
                    列 As String, 値 As String, where句 As String) As String
    Const methodName As String = "Proc動的クロス集計SQL生成"
    ' 宣言部
    Dim sql As String
    Dim groupClause As String
    Dim selectClause As String
    Dim whereClause As String
On Error GoTo ErrHandler
        
    ' SELECTとGROUP BY
    If 行1 <> "" Then
        selectClause = 行1
        groupClause = " GROUP BY " & 行1
    End If

    If 行2 <> "" Then
        If selectClause <> "" Then
            selectClause = selectClause & ", " & 行2
            groupClause = groupClause & ", " & 行2
        Else
            selectClause = 行2
            groupClause = " GROUP BY " & 行2
        End If
    End If

    ' WHERE句
    If where句 <> "" Then
        whereClause = " WHERE " & where句
    Else
        whereClause = ""
    End If

    ' SQL構築
sql = " TRANSFORM Round(Avg(" & 値 & "), 2) AS 式1 " & _
      " SELECT " & selectClause & "," & _
      " Count(" & 列 & ") AS データ数," & _
      " Round(Avg(" & 値 & "), 2) AS 全体 " & _
      " FROM D_集計元データ" & _
      whereClause & _
      groupClause & _
      " ORDER BY " & selectClause & _
      " PIVOT " & 列 & ";"
    
    Proc動的クロス集計SQL生成 = sql

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' 動的生成クエリ（チャート用集計クエリ + フィルタ対応）
' ***********************************************
Public Function Proc動的チャート用SQL生成( _
                field1 As String, field2 As String, _
                sumField As String, aggFunc As String, _
                filterValue1 As String, filterValue2 As String) As String
    Const methodName As String = "Proc動的チャート用SQL生成"
    ' 宣言部
    Dim sql As String
    Dim validAgg As String
    Dim selectField1 As String
    Dim selectField2 As String
    Dim groupField1 As String
    Dim groupField2 As String
    Dim whereClause As String
On Error GoTo ErrHandler
    
    ' 集計関数のバリデーション
    Select Case LCase(aggFunc)
        Case "sum": validAgg = "Sum"
        Case "avg": validAgg = "Avg"
        Case Else:  validAgg = "Avg"
    End Select

    ' field1, field2 が空白なら '' に置き換え
    selectField1 = IIf(Trim(field1) = "", "''", field1)
    groupField1 = selectField1
    selectField2 = IIf(Trim(field2) = "", "''", field2)
    groupField2 = selectField2

    ' HAVING句生成
    If Trim(field1) <> "" And Trim(filterValue1) <> "" Then
        whereClause = field1 & " = " & Convert条件値文字列(filterValue1)
    End If

    If Trim(field2) <> "" And Trim(filterValue2) <> "" Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        whereClause = whereClause & field2 & " = " & Convert条件値文字列(filterValue2)
    End If

    ' SQL 組み立て
    sql = "SELECT TOP 50 " & _
          selectField1 & " AS 項目１, " & _
          selectField2 & " AS 項目２, " & _
          validAgg & "(" & sumField & ") AS 値 " & _
          "FROM D_集計元データ_フィルタ "
    
    ' WHERE句追加
    If whereClause <> "" Then
        sql = sql & "WHERE " & whereClause
    End If
    
    sql = sql & _
          " GROUP BY " & _
          groupField1 & ", " & _
          groupField2 & _
          " ORDER BY " & groupField1 & " DESC, " & groupField2 & " DESC"
        
    Proc動的チャート用SQL生成 = sql

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



' ******************************************************
' 固定テーブル：T_SLOT集計区分 のリスト取得用SQL生成
' ******************************************************
Public Function Procリスト用SQL生成(fieldName As String) As String
    Const methodName As String = "Procリスト用SQL生成"
    ' 宣言部
    Dim sql As String
On Error GoTo ErrHandler

    ' メイン処理
    sql = " SELECT DISTINCT " & fieldName & _
          " FROM D_集計元データ" & _
          " ORDER BY " & fieldName & ";"
    
    Procリスト用SQL生成 = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' D_集計元データ_フィルタ（動的変更）
' ***********************************************
Public Sub Procグラフ用クエリ動的変更( _
                Optional 機種ID As Variant, _
                Optional 台番号 As Variant, _
                Optional 開始日 As Variant, _
                Optional 終了日 As Variant)
    Const methodName As String = "Procグラフ用クエリ動的変更"
    ' 宣言部
    Const クエリ名 As String = "D_集計元データ_フィルタ"
    Const テーブル名 As String = "D_集計元データ"
    Dim 条件 As String
    Dim 新SQL As String
On Error GoTo ErrHandler
    
    ' 条件作成
    If Not IsNull(機種ID) And Trim(機種ID & "") <> "" Then
        条件 = 条件 & " AND [機種ID] = " & 機種ID
    End If

    If Not IsNull(台番号) And Trim(台番号 & "") <> "" Then
        条件 = 条件 & " AND [台番号] = " & 台番号
    End If

    If Not IsNull(開始日) And Trim(開始日 & "") <> "" Then
        条件 = 条件 & " AND [日付] >= #" & Format(開始日, "yyyy/mm/dd") & "#"
    End If

    If Not IsNull(終了日) And Trim(終了日 & "") <> "" Then
        条件 = 条件 & " AND [日付] <= #" & Format(終了日, "yyyy/mm/dd") & "#"
    End If

    ' WHERE句作成
    If 条件 <> "" Then
        条件 = " WHERE " & Mid(条件, 6)  ' 先頭の AND を削除して WHERE に
    End If

    ' SQL作成
    新SQL = "SELECT * FROM " & テーブル名 & 条件

    ' SQL 置換
    Call Procクエリ内SQLを書き換え(クエリ名, 新SQL)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


