Attribute VB_Name = "Data"
Option Compare Database
Option Explicit




' **************************************************
' 店舗名を1件取得（DLookup版）
' **************************************************
Public Function Get店舗名() As String
    Const methodName As String = "Get店舗名"
On Error Resume Next
    Get店舗名 = Nz(DLookup("店舗名", "M_店舗マスタ"), "")

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function

' ***********************************************
' 店舗名を設定する（M_店舗マスタは1件前提）
' ***********************************************
Public Sub Set店舗名(newName As String)
    Const methodName As String = "Set店舗名"
On Error GoTo ErrHandler

    CurrentDb.Execute "UPDATE M_店舗マスタ SET 店舗名 = '" & Replace(newName, "'", "''") & "';", dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Sub

' ***********************************************
' 店舗挿入クエリを実行
' ***********************************************
Public Sub Add店舗名()
    Const methodName As String = "Add店舗名"
On Error GoTo ErrHandler

    CurrentDb.Execute "DELETE FROM M_店舗マスタ;", dbFailOnError
    CurrentDb.Execute "INSERT INTO M_店舗マスタ (店舗名) VALUES ('店舗名を入力してください');", dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Sub



' **************************************************
' CSVインポート
' **************************************************
Function dataインポートCSV(csvPath As String, targetTable As String) As Boolean
    Const methodName As String = "dataインポートCSV"
    ' 宣言部
    Dim fso As Object, ts As Object
    Dim db As DAO.Database
    Dim fields As Collection
    Dim sql As String
    Dim colCount As Long
    Dim i As Long
On Error GoTo ErrHandler

    dataインポートCSV = 0

    ' メイン処理
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(csvPath, 1)
    Set db = CurrentDb
    
    ' 列数確認
    colCount = db.TableDefs(targetTable).fields.count
    
    Do While Not ts.AtEndOfStream
        Set fields = Parse単行読み込み(ts.ReadLine)
        
        If colCount <> fields.count Then
            ' 自作エラー
            Err.Raise ERR_BIZ, , "取り込みファイルのデータ形式が一致していません。"
        End If
        
        If fields.count = colCount Then
            sql = "INSERT INTO " & targetTable & " VALUES ('"
            For i = 1 To fields.count
                sql = sql & Replace(fields(i), "'", "''")
                If i < fields.count Then sql = sql & "','"
            Next
            sql = sql & "')"
            db.Execute sql
        End If
    Loop

    ' 結果セット
    dataインポートCSV = Not DCount("*", targetTable) = 0
        
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function

' **************************************************
Function Parse単行読み込み(line As String) As Collection
    Const methodName As String = "Parse単行読み込み"
    ' 宣言部
    Dim result As New Collection
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String
    Dim field As String
On Error GoTo ErrHandler
    
    ' メイン処理
    inQuotes = False
    field = ""
    
    For i = 1 To Len(line)
        ch = Mid(line, i, 1)
        
        If ch = """" Then
            ' ダブルクォートの場合
            If inQuotes And i < Len(line) And Mid(line, i + 1, 1) = """" Then
                ' 連続するダブルクォート → " を1つ追加
                field = field & """"
                i = i + 1
            Else
                ' クォートの開閉
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            ' カンマ区切り（クォート外のみ）
            result.Add field
            field = ""
        Else
            field = field & ch
        End If
    Next
    
    ' 最後のフィールド追加
    result.Add field
    
    Set Parse単行読み込み = result

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ***********************************************
' テーブルのレコードを条件付きで削除
' ***********************************************
Public Function dataテーブルクリア(tableName As String, Optional whereCondition As String = "") As Boolean
    Const methodName As String = "dataテーブルクリア"
    ' 宣言部
    Dim sql As String
On Error GoTo ErrHandler

    dataテーブルクリア = False
    
    ' 基本のDELETE文
    sql = "DELETE FROM [" & tableName & "]"

    ' 条件が渡されたらWHERE句を付ける
    If Trim(whereCondition) <> "" Then
        sql = sql & " WHERE " & whereCondition
    End If

    ' 実行
    CurrentDb.Execute sql, dbFailOnError
    dataテーブルクリア = True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


'' ***********************************************
'' 単一クエリを実行する
'' ***********************************************
'Function クエリ実行_単一(queryName As String) As Boolean
'    Const methodName As String = "クエリ実行_単一"
'On Error Resume Next
'    クエリ実行_単一 = False
'
'    ' メイン処理
'    CurrentDb.QueryDefs(queryName).Execute dbFailOnError
'
'    クエリ実行_単一 = True
'
'ErrHandler:
'    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
'
'End Function


' ***********************************************
' 配列内のクエリをすべて実行する
' ***********************************************
Function Procクエリリスト一括実行(queryNames As Variant) As Boolean
    Const methodName As String = "Procクエリリスト一括実行"
    ' 宣言部
    Dim i As Long
    Dim db As DAO.Database
    Dim queryName As String
On Error GoTo ErrHandler
    
    ' メイン処理
    Procクエリリスト一括実行 = False
    
    ' メイン処理
    Set db = CurrentDb
    
    For i = LBound(queryNames) To UBound(queryNames)
        queryName = queryNames(i)
        db.QueryDefs(queryName).Execute dbFailOnError
    Next i
    
    Procクエリリスト一括実行 = True

ErrHandler:
    Call ErrorSave(methodName, False, queryName) '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function

