Attribute VB_Name = "SQL"
Option Compare Database
Option Explicit


' ***********************************************
' 動的にSQLをセットする
' ***********************************************
Public Sub Procクエリ内SQLを書き換え(queryName As String, newSQL As String)
    Const methodName As String = "Procクエリ内SQLを書き換え"
    Dim qdf As DAO.QueryDef
On Error GoTo ErrHandler

    ' メイン処理
    Set qdf = CurrentDb.QueryDefs(queryName)
    qdf.sql = newSQL
    
    Call ProcNothing(qdf)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



' ***********************************************
' クエリ D_集計元データ のSQLを動的に変更する
' ***********************************************
Public Sub Update集計元Query(Optional 機種ID As Variant = Null, _
                             Optional 台番号ID As Variant = Null, _
                             Optional 開始日 As Variant = Null, _
                             Optional 終了日 As Variant = Null, _
                             Optional その他 As Variant = Null)
    Const methodName As String = "Update集計元Query"
    Const queryName As String = "D_集計元データ_フィルタ"
    Dim sql As String
    Dim 条件 As String
On Error GoTo ErrHandler

    ' メイン処理
    条件 = ""

    If Not IsNull(機種ID) And 機種ID <> 0 And 機種ID <> "" Then
        条件 = JoinText(条件, " AND 機種ID = " & 機種ID, True)
    End If

    If Not IsNull(台番号ID) And 台番号ID <> 0 And 台番号ID <> "" Then
        条件 = JoinText(条件, " AND 台番号 = " & 台番号ID, True)
    End If

    If Not IsNull(開始日) And 開始日 <> "" Then
        条件 = JoinText(条件, " AND 日付 >= #" & Format(開始日, "yyyy/mm/dd") & "#", True)
    End If

    If Not IsNull(終了日) And 終了日 <> "" Then
        条件 = JoinText(条件, " AND 日付 <= #" & Format(終了日, "yyyy/mm/dd") & "#", True)
    End If
    
    If Not IsNull(その他) And その他 <> "" Then
        条件 = JoinText(条件, " AND " & その他, True)
    End If

    If 条件 <> "" Then
        条件 = " WHERE " & Mid(条件, 6)  ' 先頭の AND を削除して WHERE に
    End If

    sql = "SELECT * " & _
          "FROM D_集計元データ " & 条件 & ";"
    
    ' 既存クエリを取得してSQLを書き換え
    Call Procクエリ内SQLを書き換え(queryName, sql)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub

' ***********************************************
' クエリ D_集計元データ のSQLを動的に変更する
' ***********************************************
Public Sub Update集計元Query_最新(Optional 機種ID As Variant = Null, _
                                Optional 台番号ID As Variant = Null, _
                                Optional 開始日 As Variant = Null, _
                                Optional 終了日 As Variant = Null, _
                                Optional その他 As Variant = Null)
    Const methodName As String = "Update集計元Query_最新"
    ' 宣言部
    Const queryName As String = "D_集計元データ_最新_フィルタ"
    Dim sql As String
    Dim 条件 As String
On Error GoTo ErrHandler
    
    ' メイン処理
    条件 = ""

    If Not IsNull(機種ID) And 機種ID <> 0 And 機種ID <> "" Then
        条件 = JoinText(条件, " AND 機種ID = " & 機種ID, True)
    End If

    If Not IsNull(台番号ID) And 台番号ID <> 0 And 台番号ID <> "" Then
        条件 = JoinText(条件, " AND 台番号 = " & 台番号ID, True)
    End If

    If Not IsNull(開始日) And 開始日 <> "" Then
        条件 = JoinText(条件, " AND 日付 >= #" & Format(開始日, "yyyy/mm/dd") & "#", True)
    End If

    If Not IsNull(終了日) And 終了日 <> "" Then
        条件 = JoinText(条件, " AND 日付 <= #" & Format(終了日, "yyyy/mm/dd") & "#", True)
    End If
    
    If Not IsNull(その他) And その他 <> "" Then
        条件 = JoinText(条件, " AND " & その他, True)
    End If

    If 条件 <> "" Then
        条件 = " WHERE " & Mid(条件, 6)  ' 先頭の AND を削除して WHERE に
    End If

    sql = "SELECT * " & _
          "FROM D_集計元データ_最新 " & 条件 & ";"
    
    ' 既存クエリを取得してSQLを書き換え
    Call Procクエリ内SQLを書き換え(queryName, sql)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub



' ***********************************************
' グループ化フィールドと集計対象フィールドからSQL文字列を生成して返す
' ***********************************************
Public Function GenerateGraphSQL_A(項目列 As String, 集計列 As String) As String
    Const methodName As String = "GenerateGraphSQL_A"
    ' 宣言部
    Dim sql As String
On Error GoTo ErrHandler

    ' メイン処理
    sql = "SELECT " & 項目列 & ", Avg(" & 集計列 & ") AS 集計値 " & _
          "FROM D_集計元データ_フィルタ " & _
          "GROUP BY " & 項目列

'    Call LogDebug("GenerateGraphSQL_A", sql)
    
    GenerateGraphSQL_A = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function

' ***********************************************
' グループ化フィールドと集計対象列からSQL文字列を生成して返す
' 月旬の条件を動的に付与可能
' ***********************************************
Public Function GenerateGraphSQL_B(項目列1 As String, 項目列2 As String, 集計列 As String, _
                                   Optional フィルタ値 As Variant) As String
    Const methodName As String = "GenerateGraphSQL_B"
    ' 宣言部
    Dim sql As String
    Dim whereStr As String
On Error GoTo ErrHandler

    ' メイン処理
    If Not IsMissing(フィルタ値) Then
        If Not IsNull(フィルタ値) And Trim(フィルタ値 & "") <> "" Then
            whereStr = " WHERE CStr([" & 項目列2 & "]) = """ & フィルタ値 & """"
        End If
    End If
    
    ' SQL組み立て
    sql = "SELECT " & 項目列1 & ", Avg(" & 集計列 & ") AS 集計値 " & _
      "FROM D_集計元データ_フィルタ" & whereStr & _
      " GROUP BY " & 項目列1 & ";"
    
    GenerateGraphSQL_B = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



' ***********************************************
' クエリ L_台番号リスト のSQLを動的に変更する
' ***********************************************
Public Sub Update台番号Query(Optional 機種ID As Variant)
    Const methodName As String = "Update台番号Query"
    ' 宣言部
    Dim sql As String
    Const queryName As String = "L_台番号リスト_フィルタ"
On Error GoTo ErrHandler

    ' メイン処理
    sql = "SELECT * FROM L_台番号リスト"
    
    ' 機種IDが指定されていればWHERE句を付与
    If Not IsNull(機種ID) And 機種ID <> "" Then
        sql = sql & " WHERE 機種ID = " & 機種ID
    Else
        sql = sql & " WHERE 1 = 0"
    End If
    
    ' 既存クエリを取得してSQLを書き換え
    Call Procクエリ内SQLを書き換え(queryName, sql)
    
    Debug.Print sql
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ***********************************************
' クエリ L_絞り込みリスト のSQLを動的に変更する
' ***********************************************
Public Sub Update絞り込みQuery(ByVal fieldName As String)
    Const methodName As String = "Update絞り込みQuery"
    ' 宣言部
    Dim sql As String
    Const queryName As String = "L_絞り込みリスト"
On Error GoTo ErrHandler

    ' フィールド名が空なら処理しない
    If Nz(fieldName, "") = "" Then Exit Sub

    ' メイン処理
    sql = "SELECT DISTINCT " & fieldName & " AS 絞り込み " & _
          "FROM T_SLOT集計区分 " & _
          "ORDER BY " & fieldName & ";"

    ' 既存クエリを取得してSQLを書き換え
    Call Procクエリ内SQLを書き換え(queryName, sql)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ***********************************************
' 引数からSQL文字列を生成して返す
' ***********************************************
Public Function Generate積み上げ履歴SQL(機種ID As String, 並び順 As String) As String
    Const methodName As String = "Generate積み上げ履歴SQL"
    ' 宣言部
    Dim sql As String
On Error GoTo ErrHandler

    ' メイン処理
    sql = "SELECT 機種名, 日付, 件数, 積み上げ日時 " & _
          "FROM Check_積み上げ日_機種別 " & _
          "WHERE 機種ID = " & 機種ID & " " & _
          "ORDER BY " & 並び順 & ";"
    
    Generate積み上げ履歴SQL = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



