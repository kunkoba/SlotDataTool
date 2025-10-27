Attribute VB_Name = "_Utility"
Option Compare Database
Option Explicit


' **************************************************
' すべてのフォームを統一仕様にしよう1
' **************************************************
Public Sub utlフォーム仕様統一(除外フォーム As String, isPopUp As Boolean)
    Const methodName As String = "utlフォーム仕様統一"
    ' 宣言部
    Dim accObj As AccessObject
    Dim frm As Form
    Dim ctl As Control
On Error GoTo ErrHandler
    
    ' メイン処理
    For Each accObj In CurrentProject.AllForms
    
        If accObj.name <> 除外フォーム And Left(accObj.name, 2) <> "F_" Then
        
            ' フォームをデザインビューかつ非表示で開く
            DoCmd.OpenForm accObj.name, acDesign, , , , acHidden
            Set frm = Forms(accObj.name)
    
            ' フォーム全体のプロパティ変更
            With frm
                .PopUp = isPopUp                      ' ポップアップにする
                .Modal = isPopUp                      ' モーダルにはしない（作業ウィンドウに固定）
                .ShortcutMenu = Not isPopUp       ' ★右クリック無効化（これを追加！）
    '            .RecordSelectors = False           ' レコードセレクタ非表示
    '            .AllowDesignChanges = True         ' デザイン変更許可（念のため）
    '            .AllowDatasheetView = False        ' データシート表示は無効
    '            .BorderStyle = 1
    '            .MinMaxButtons = 0
    '            .ScrollBars = 0                    ' スクロールバーなし（必要なら調整）
    '            .AutoCenter = True
    '            .AllowLayoutView = False
    '            .NavigationButtons = False
    '            .Moveable = True
            End With
    
            ' 保存して閉じる
            DoCmd.Close acForm, accObj.name, acSaveYes
        
        End If
        
    Next accObj

    Call ShowToast("すべてのフォームの仕様を統一しました。", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（発生したエラーを確実にキャッチするため）
    
End Sub



' **************************************************
' すべての関連オブジェクトを取得する（文字列返却版）
' 引数 delimiter : 区切り文字（例 vbCrLf, vbTab, "," など）
' **************************************************
Public Function utl依存オブジェクト抽出文字列(除外フォーム As String, ByVal 検索ワード As String, Optional ByVal delimiter As String = ",") As String
    Const methodName As String = "utl依存オブジェクト抽出文字列"
    ' 宣言部
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim doc As AccessObject
    Dim frm As Form
    Dim rpt As Report
    Dim objName As String
    Dim ctl As Control
    Dim result As String
On Error GoTo ErrHandler
    
    ' メイン処理
    Set db = CurrentDb
    Application.Echo False

    ' === クエリ =============================
    For Each qdf In db.QueryDefs
        If InStr(1, qdf.sql, 検索ワード, vbTextCompare) > 0 Then
            If Left(qdf.name, 1) <> "~" Then
                result = result & "Query:" & qdf.name & delimiter
            End If
        End If
    Next

    ' === フォーム ===========================
    For Each doc In Application.CurrentProject.AllForms
        objName = doc.name
        
        If Left(objName, 1) <> "~" And Left(objName, 2) <> "F_" Then
                
            If objName <> 除外フォーム Then
                
                DoCmd.OpenForm objName, acDesign, , , , acHidden
                Set frm = Forms(objName)
        
                If InStr(1, frm.RecordSource, 検索ワード, vbTextCompare) > 0 Then
                    result = result & "Form:" & objName & "（RecordSource）" & delimiter
                End If
        
                For Each ctl In frm.Controls
                    If ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
                        If InStr(1, ctl.RowSource, 検索ワード, vbTextCompare) > 0 Then
                            result = result & "Form:" & objName & "（RowSource: " & ctl.name & "）" & delimiter
                        End If
                    End If
                Next ctl
        
                DoCmd.Close acForm, objName, acSaveNo
            
            End If
            
        End If
    Next

    ' === レポート ===========================
    For Each doc In Application.CurrentProject.AllReports
        objName = doc.name
        DoCmd.OpenReport objName, acDesign, , , acHidden
        Set rpt = Reports(objName)

        If InStr(1, rpt.RecordSource, 検索ワード, vbTextCompare) > 0 Then
            If Left(objName, 1) <> "~" Then
                result = result & "Report:" & objName & "（RecordSource）" & delimiter
            End If
        End If

        DoCmd.Close acReport, objName, acSaveNo
    Next

    ' 戻り値に設定
    utl依存オブジェクト抽出文字列 = result
    
    Call ShowToast("関連オブジェクトの抽出は完了しました。", E_Color_Blue)

ErrHandler:
    Application.Echo True
    Call ErrorSave(methodName, False, objName) '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



' **************************************************
' すべてのオブジェクト名を文字列で返す
' 引数 mode
'   1 = テーブル一覧
'   2 = クエリ一覧
'   3 = フォーム一覧
' 引数 delim 区切り文字
'   例: "," や vbCrLf や vbTab
' **************************************************
Public Function utlオブジェクト名一覧取得(ByVal mode As Integer, Optional ByVal delim As String = ",") As String
    Const methodName As String = "utlオブジェクト名一覧取得"
    Dim obj As AccessObject
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim result As String
On Error GoTo ErrHandler

    result = ""

    Set db = CurrentDb
    Select Case mode
        Case 1 ' テーブル
            Set db = CurrentDb
            For Each tdf In db.TableDefs
                ' システムテーブル（MSys～）は除外
                If Left(tdf.name, 4) <> "MSys" Then
                    If result <> "" Then result = result & delim
                    result = result & tdf.name
                End If
            Next
        
        Case 2 ' クエリ
            For Each qdf In db.QueryDefs
                ' 一時クエリ（~...）は除外
                If Left(qdf.name, 1) <> "~" Then
                    If result <> "" Then result = result & delim
                    result = result & qdf.name
                End If
            Next
        
        Case 3 ' フォーム
            For Each obj In CurrentProject.AllForms
                If result <> "" Then result = result & delim
                result = result & obj.name
            Next
        
        Case Else
            result = "modeは 1=テーブル / 2=クエリ / 3=フォーム を指定してください"
    End Select

    utlオブジェクト名一覧取得 = result
    
'    MsgBox "オブジェクト名一覧取得が完了しました。", vbInformation
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function




' ***********************************************
' テーブル定義生成
' ***********************************************
Sub utlテーブル定義生成(tableName As String)
    Const methodName As String = "utlテーブル定義生成"
    ' 宣言部
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.field
    Dim idx As DAO.index
    Dim pkFields As Collection
On Error GoTo ErrHandler
    
    ' メイン処理
    Set pkFields = New Collection
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    ' 主キーのフィールドを収集
    For Each idx In tdf.Indexes
        If idx.Primary Then
            Dim f As DAO.field
            For Each f In idx.fields
                pkFields.Add f.name
            Next f
        End If
    Next idx
    
    ' CREATE TABLE 文作成
    Dim sql As String
    sql = "CREATE TABLE [" & tableName & "] (" & vbCrLf
    
    Dim fieldSQL As String
    For Each fld In tdf.fields
        fieldSQL = "  [" & fld.name & "] " & Convertフィールドタイプ(fld.Type, fld.size)
        If fld.Required Then fieldSQL = fieldSQL & " NOT NULL"
        fieldSQL = fieldSQL & ","
        sql = sql & fieldSQL & vbCrLf
    Next fld
    
    ' 主キー設定
    If pkFields.count > 0 Then
        Dim pkSQL As String
        pkSQL = "  PRIMARY KEY ("
        Dim i As Long
        For i = 1 To pkFields.count
            pkSQL = pkSQL & "[" & pkFields(i) & "]"
            If i < pkFields.count Then pkSQL = pkSQL & ", "
        Next i
        pkSQL = pkSQL & ")"
        sql = sql & pkSQL & vbCrLf
    Else
        ' 最後のカンマを削除
        sql = Left(sql, Len(sql) - 3) & vbCrLf
    End If
    
    sql = sql & ");"
    
    Call LogDebug("utlテーブル定義生成", sql)

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(db)
    Call ProcNothing(tdf)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Sub


' ***********************************************
' XXX
' ***********************************************
Function Convertフィールドタイプ(type_ As Long, size As Long) As String
On Error Resume Next
    Select Case type_
        Case dbText
            Convertフィールドタイプ = "TEXT(" & size & ")"
        Case dbMemo
            Convertフィールドタイプ = "MEMO"
        Case dbByte
            Convertフィールドタイプ = "BYTE"
        Case dbInteger
            Convertフィールドタイプ = "INTEGER"
        Case dbLong
            Convertフィールドタイプ = "LONG"
        Case dbCurrency
            Convertフィールドタイプ = "CURRENCY"
        Case dbSingle
            Convertフィールドタイプ = "SINGLE"
        Case dbDouble
            Convertフィールドタイプ = "DOUBLE"
        Case dbDate
            Convertフィールドタイプ = "DATETIME"
        Case dbBoolean
            Convertフィールドタイプ = "YESNO"
        Case Else
            Convertフィールドタイプ = "TEXT"
    End Select
    
End Function



' ************************************************
' ソースコード出力
' ************************************************
Sub utlソースコード出力()
    Const methodName As String = "utlソースコード出力"
    Const Path_Source As String = "SourceCode\"
    ' 宣言部
    Dim vbcmp As Object
    Dim strFileName As String
    Dim strExt As String
    Dim batPath As String
    Dim strCmd As String
    Dim outPath As String
On Error GoTo ErrHandler
    
    ' 出力フォルダ設定
    outPath = CurrentProject.path & "\" & Path_Source
    Call Procフォルダが存在しない場合は作成(outPath)
    
    ' メイン処理
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            'ファイル名までを設定
            strFileName = outPath & .name
            '拡張子を設定
            Select Case .Type
                Case 1        '標準モジュール
                    strExt = ".bas"
                Case 2        'クラスモジュール
                    strExt = ".cls"
                Case 100      'フォーム/レポートのモジュール
                    strExt = ".cls"
                Case Else
                    strExt = ".txt"
            End Select
            ' モジュールをエクスポート
            .Export strFileName & strExt
        End With
    Next vbcmp

    Sleep 1000
    
    batPath = CurrentProject.path & "\ExportedObjects\__Convert_UTF8.bat"
    strCmd = "cmd /c """ & batPath & """"
    shell strCmd, vbHide

    Call ShowToast("エクスポートが完了しました。", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True, strFileName) '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(vbcmp)
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub




