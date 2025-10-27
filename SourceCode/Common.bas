Attribute VB_Name = "Common"
Option Compare Database
Option Explicit



' ***********************************************
' とにかくクリア
' ***********************************************
Public Sub ProcNothing(ByRef obj)
On Error Resume Next
    obj.Close
    Set obj = Nothing

'    Debug.Print Now, "ProcNothing", Err.Number, Err.Description
    Err.Clear
End Sub

'' ***********************************************
'' 共通開始処理
'' ***********************************************
'Public Sub ProcInitSetting()
'On Error Resume Next
'    DoCmd.SetWarnings False
'
'End Sub

'' ***********************************************
'' 共通終了処理
'' ***********************************************
'Public Sub ProcFinally()
'On Error Resume Next
'    DoCmd.SetWarnings True
'    DoCmd.Echo True
'    Screen.MousePointer = 0    ' 0 = 通常の矢印
'
'End Sub


' ******************************************************
' フォームが開いているかどうか
' ******************************************************
Function IsFormOpen(formName As String) As Boolean
On Error Resume Next
    If (SysCmd(acSysCmdGetObjectState, acForm, formName) And acObjStateOpen) <> 0 Then
        If Forms(formName).CurrentView = 1 Then
            IsFormOpen = True
        End If
    End If
   
End Function

' ***********************************************
' 文字列結合関数
' ***********************************************
Function JoinText(word1 As String, word2 As String, Optional addNewLine As Boolean = False) As String
On Error Resume Next

    If addNewLine Then
        JoinText = word1 & word2 & vbCrLf
    Else
        JoinText = word1 & word2
    End If
    
End Function



' ***********************************************
' センチ変換
' ***********************************************
Public Function CmToTwips(cm As Double) As Long
On Error Resume Next
    CmToTwips = cm * 567
    
End Function



' ***********************************************
' オブジェクトを取得する（クエリ）
' ***********************************************
Public Function Getクエリ一覧(Optional filter As String = "") As Variant
    Const methodName As String = "Getクエリ一覧"
    ' 宣言部
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim names() As String
    Dim count As Long
On Error GoTo ErrHandler
    
    ' メイン処理
    Set db = CurrentDb
    count = 0

    For Each qdf In db.QueryDefs
        If Left(qdf.name, 1) <> "~" Then ' テンポラリクエリ除外
            If filter = "" Or qdf.name Like filter Then
                ReDim Preserve names(count)
                names(count) = qdf.name
                count = count + 1
            End If
        End If
    Next

    If count = 0 Then
        Getクエリ一覧 = Array() ' 空配列
    Else
        Getクエリ一覧 = names
    End If

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(qdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
        
End Function






' ***********************************************
' 色コード変換
' ***********************************************
Public Function ConvertHexColorToRGB(hexColor As String) As Long
    Dim r As Integer
    Dim g As Integer
    Dim b As Integer
On Error Resume Next
    ' 例: hexColor = "#CDDCAF"

    ' "#" を取り除く
    If Left(hexColor, 1) = "#" Then
        hexColor = Mid(hexColor, 2)
    End If

    ' R, G, B を16進→10進に変換
    r = CInt("&H" & Mid(hexColor, 1, 2))
    g = CInt("&H" & Mid(hexColor, 3, 2))
    b = CInt("&H" & Mid(hexColor, 5, 2))

    ' RGB関数で色を返す（Long型）
    ConvertHexColorToRGB = RGB(r, g, b)

End Function


' ***********************************************
' クエリをCSVで出力する関数
' 引数:
'   queryName  … クエリ名（例: "Q_出力データ"）
'   folderPath … 出力フォルダ（末尾に \ は不要でも可）
'   fileName   … 出力ファイル名（.csv 拡張子は自動付与）
' ***********************************************
Public Sub クエリ出力ToCSV(queryName As String, folderPath As String, fileName As String)
    Const methodName As String = "クエリ出力ToCSV"
    ' 宣言部
    Dim fullPath As String
On Error GoTo ErrHandler
    
    ' フォルダ末尾にバックスラッシュを付ける
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If

    ' .csv 拡張子がなければ追加
    If LCase(Right(fileName, 4)) <> ".csv" Then
        fileName = fileName & ".csv"
    End If

    ' フルパス作成
    fullPath = folderPath & fileName

    ' クエリをCSVでエクスポート
    DoCmd.TransferText _
        TransferType:=acExportDelim, _
        tableName:=queryName, _
        fileName:=fullPath, _
        HasFieldNames:=True

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub





' ******************************************************
' 日付文字型を日付へ変換（yyyymmdd / yymmdd 両対応）
' yymmdd の場合は先頭に 20 を付けて変換
' ******************************************************
Function Convert文字列から日付(strDate As Variant) As Variant
    Const methodName As String = "Convert文字列から日付"
    Dim s As String
    Dim Y As Long, m As Long, d As Long
    Dim yyyy As Long
On Error GoTo ErrHandler
    
    ' Nullチェック
    If IsNull(strDate) Then
        Convert文字列から日付 = Null
        Exit Function
    End If

    s = Trim(CStr(strDate))
    
    ' yyyymmdd
    If Len(s) = 8 And s Like "########" Then
        Y = CLng(Left(s, 4))
        m = CLng(Mid(s, 5, 2))
        d = CLng(Right(s, 2))
        Convert文字列から日付 = DateSerial(Y, m, d)
        Exit Function
    End If
    
    ' yymmdd → 20yy mm dd
    If Len(s) = 6 And s Like "######" Then
        yyyy = 2000 + CLng(Left(s, 2))
        m = CLng(Mid(s, 3, 2))
        d = CLng(Right(s, 2))
        
        Convert文字列から日付 = DateSerial(yyyy, m, d)
        Exit Function
    End If
    
    ' それ以外は Null
    Convert文字列から日付 = Null
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function




' ***********************************************
' 数値ならそのまま、文字列ならシングルクォーテーションで囲む
' 日付型なら # で囲む
' ***********************************************
Public Function Convert条件値文字列(value As Variant) As String
    Const methodName As String = "Convert条件値文字列"
On Error GoTo ErrHandler

    If IsNull(value) Then
        Convert条件値文字列 = "Null"
    ElseIf IsDate(value) Then
        ' 日付は # で囲んでSQL用に整形
        Convert条件値文字列 = "#" & Format(value, "yyyy/mm/dd") & "#"
    ElseIf IsNumeric(value) Then
        Convert条件値文字列 = CStr(value)
    Else
        ' 文字列 → シングルクォートで囲み、シングルクォートをエスケープ
        Convert条件値文字列 = "'" & Replace(value, "'", "''") & "'"
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



' ======================================================
' PC情報取得関数（配列で返す）
' ======================================================
Public Function Getパソコン情報() As Variant
    Const methodName As String = "Getパソコン情報"
    Dim infoArr(2) As String
    Dim pcName As String
    Dim userName As String
    Dim ipAddr As String
    Dim wmi As Object, colItems As Object, objItem As Object
On Error GoTo ErrHandler
    
    ' パソコン名
    pcName = Environ("COMPUTERNAME")
    
    ' ユーザー名
    userName = Environ("USERNAME")
    
    ' IPアドレス取得（アクティブな接続1件）
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = wmi.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    ipAddr = ""
    For Each objItem In colItems
        If Not IsNull(objItem.IPAddress) Then
            ipAddr = objItem.IPAddress(0)
            Exit For
        End If
    Next
    
    ' 配列に格納
    infoArr(0) = pcName
    infoArr(1) = userName
    infoArr(2) = ipAddr
    
    Getパソコン情報 = infoArr
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function


' ============================================================
' 関数名 : 配列要素存在チェック
' 機　能 : 指定した文字列が配列のいずれかの要素に含まれているかを判定する
' 引　数 : targetArray - 検索対象の配列（Variant型推奨）
'　　　 : searchValue  - 探したい文字列
' 戻り値 : Boolean（含まれていれば True）
' ============================================================
Public Function 配列要素存在チェック(targetArray As Variant, searchValue As String) As Boolean
    Const methodName As String = "配列要素存在チェック"
    Dim v As Variant
On Error GoTo ErrHandler
    配列要素存在チェック = False  ' 初期値
    
    ' 空配列チェック
    If IsEmpty(targetArray) Then Exit Function
    If IsNull(searchValue) Or searchValue = "" Then Exit Function
    
    ' 配列の各要素をループして一致を確認
    For Each v In targetArray
        If v = searchValue Then
            配列要素存在チェック = True
            Exit Function
        End If
    Next v
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



