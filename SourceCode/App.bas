Attribute VB_Name = "App"
Option Compare Database
Option Explicit


Public Const T_AppSetting As String = "S_アプリ設定"


' **************************************************
Public Function App_接続チェック()
    Const methodName As String = "App_接続チェック"
    Dim newPath As String
'    Dim flg As Boolean
On Error Resume Next

    ' 規定のデータファイルに接続する（配下のDataを優先して接続する）
    newPath = CurrentProject.path & "\" & PATH_DATA & "\" & App_データファイル名 & SYS拡張子
    If Procファイル存在確認(newPath) Then
        ' 所定の場所にデータファイルがあれば、リンク先を更新する（無ければ更新しない）
        Call Procリンクテーブル一括更新(newPath)
    
    End If

    App_接続チェック = DCount("*", "M_店舗マスタ") > 0
    If Not App_接続チェック Then Call ShowConfirm("システム", "データファイルが見つかりません。　再度、設定をしてください。", vbYes)
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function

' **************************************************
Public Function App_アプリ名() As String
On Error Resume Next
    App_アプリ名 = DLookup("アプリ名", T_AppSetting)
        
End Function
' **************************************************
Public Function App_ライセンスマーク() As String
On Error Resume Next
    App_ライセンスマーク = DLookup("ライセンスマーク", T_AppSetting)
        
End Function
' **************************************************
Public Function App_バージョン() As String
On Error Resume Next
    App_バージョン = DLookup("バージョン", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MACアドレス1() As String
On Error Resume Next
    App_MACアドレス1 = DLookup("MACアドレス1", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MACアドレス2() As String
On Error Resume Next
    App_MACアドレス2 = DLookup("MACアドレス2", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MACアドレス3() As String
On Error Resume Next
    App_MACアドレス3 = DLookup("MACアドレス3", T_AppSetting)
        
End Function
' **************************************************
Public Function App_アプリ有効期限()
On Error Resume Next
    App_アプリ有効期限 = DLookup("アプリ有効期限", T_AppSetting)
    
End Function
' **************************************************
Public Function App_秘密コード() As String
On Error Resume Next
    App_秘密コード = DLookup("秘密コード", T_AppSetting)
    
End Function
' **************************************************
Public Function App_暗号キー() As String
On Error Resume Next
    App_暗号キー = DLookup("暗号キー", T_AppSetting)
    
End Function
' **************************************************
Public Function App_アプリ解除日()
On Error Resume Next
    App_アプリ解除日 = DLookup("アプリ解除日", T_AppSetting)
    
End Function
' **************************************************
Public Function App_解除コード() As String
On Error Resume Next
    App_解除コード = DLookup("解除コード", T_AppSetting)
    
End Function
' **************************************************
Public Function App_リリース日()
On Error Resume Next
    App_リリース日 = DLookup("リリース日", T_AppSetting)
    
End Function
' **************************************************
Public Function App_データファイル名() As String
On Error Resume Next
    App_データファイル名 = DLookup("データファイル名", T_AppSetting)
    
End Function



' **************************************************
' Appデータ初期化
' **************************************************
Public Sub Appデータ初期化()
    Const methodName As String = "Appデータ初期化"
    ' 宣言部
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tablesToDelete As Collection
    Dim tblName As Variant
On Error GoTo ErrHandler
    
    ' メイン処理
    Set db = CurrentDb
    Set tablesToDelete = New Collection

    For Each tdf In db.TableDefs
        If Left(tdf.name, 1) = "T" Or Left(tdf.name, 1) = "M" Then
            If Left(tdf.name, 4) <> "MSys" Then
                tablesToDelete.Add tdf.name
            End If
        End If
    Next
    
On Error Resume Next
    For Each tblName In tablesToDelete
        db.Execute "DELETE FROM [" & tblName & "]", dbFailOnError
    Next

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ***********************************************
' MACアドレス更新
' ***********************************************
Public Sub MACアドレス更新()
    Const methodName As String = "MACアドレス更新"
    ' 宣言部
    Dim macs As Variant
    Dim i As Long
    Dim sql As String
    Dim vals(1 To 3) As String
On Error GoTo ErrHandler
    
    ' メイン処理
    macs = 端末MACアドレス取得配列()

    ' --- 最大3件分を空文字で初期化 ---
    For i = 1 To 3
        If IsArray(macs) And UBound(macs) >= i Then
            vals(i) = macs(i)
        Else
            vals(i) = ""
        End If
    Next

    ' --- UPDATE文作成 ---
    sql = "UPDATE S_アプリ設定 SET " & _
          "[MACアドレス1]='" & Replace(vals(1), "'", "''") & "', " & _
          "[MACアドレス2]='" & Replace(vals(2), "'", "''") & "', " & _
          "[MACアドレス3]='" & Replace(vals(3), "'", "''") & "';"

    CurrentDb.Execute sql, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Sub


' ***********************************************
' 端末MACアドレス取得配列（制限なし）
' ***********************************************
Public Function 端末MACアドレス取得配列() As Variant
    Const methodName As String = "端末MACアドレス取得配列"
    ' 宣言部
    Dim objWMIService As Object
    Dim colAdapters As Object
    Dim objAdapter As Object
    Dim tmpMacs() As String
    Dim mac As String
    Dim name As String
    Dim category As Integer
    Dim count As Long
    Dim i As Long, j As Long
    Dim tmpCat As Integer, tmpVal As String
    Dim categories() As Integer
On Error GoTo ErrHandler
    
    ' メイン処理
    count = 0
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL AND PhysicalAdapter=True")
    
    ' 一時格納
    For Each objAdapter In colAdapters
        mac = Trim(objAdapter.MACAddress)
        name = LCase(Trim(objAdapter.name & " " & objAdapter.NetConnectionID))
        If mac <> "" Then
            ' --- 種別を分類（小さい数字が優先） ---
            If InStr(name, "ethernet") > 0 Or InStr(name, "lan") > 0 Then
                category = 1   ' 有線LAN
            ElseIf InStr(name, "wireless") > 0 Or InStr(name, "wi-fi") > 0 Or InStr(name, "wifi") > 0 Then
                category = 2   ' 無線LAN
            Else
                category = 3   ' その他（仮想など）
            End If
            
            ' --- 重複チェック ---
            For j = 1 To count
                If tmpMacs(j) = mac Then GoTo SkipAdd
            Next
            
            count = count + 1
            ReDim Preserve tmpMacs(1 To count)
            ReDim Preserve categories(1 To count)
            tmpMacs(count) = mac
            categories(count) = category
        End If
SkipAdd:
    Next
    
    ' --- 優先順にソート（category順） ---
    For i = 1 To count - 1
        For j = i + 1 To count
            If categories(j) < categories(i) Then
                tmpCat = categories(i)
                categories(i) = categories(j)
                categories(j) = tmpCat
                
                tmpVal = tmpMacs(i)
                tmpMacs(i) = tmpMacs(j)
                tmpMacs(j) = tmpVal
            End If
        Next
    Next
    
    ' --- 結果を返す ---
    端末MACアドレス取得配列 = tmpMacs

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function



' ******************************************************
' アプリ利用チェック
' ******************************************************
Public Function Authアプリ利用チェック() As String
    Const methodName As String = "Authアプリ利用チェック"
    Dim errMsg As String
On Error GoTo ErrHandler

    If Not AuthMACアドレス認証チェック Then
        Authアプリ利用チェック = "許可された端末以外での利用は認めていません。"
        Exit Function
    End If
    
    If Not Auth有効期限認証チェック Then
        Authアプリ利用チェック = "アプリ利用の有効期限が切れました。"
        Exit Function
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


' ******************************************************
' 有効期限認証チェック（アプリ有効期限、アプリ解除日）
' ******************************************************
Function Auth有効期限認証チェック() As Boolean
    Const methodName As String = "Auth有効期限認証チェック"
On Error GoTo ErrHandler

    Auth有効期限認証チェック = True
    
    If Not IsNull(App_アプリ解除日) Then Exit Function
    If Date <= App_アプリ有効期限 Then Exit Function

    Auth有効期限認証チェック = False
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


' ******************************************************
' MACアドレス認証チェック（アプリ有効期限、アプリ解除日）
' ******************************************************
Function AuthMACアドレス認証チェック() As Boolean
    Const methodName As String = "Auth有効期限認証チェック"
    Dim ary1 As Variant
    Dim result As Boolean
On Error GoTo ErrHandler
    
    ary1 = 端末MACアドレス取得配列
    
    result = 配列要素存在チェック(ary1, Nz(App_MACアドレス1))
    AuthMACアドレス認証チェック = result
    If result Then Exit Function
    
    result = 配列要素存在チェック(ary1, Nz(App_MACアドレス2))
    AuthMACアドレス認証チェック = result
    If result Then Exit Function
    
    result = 配列要素存在チェック(ary1, Nz(App_MACアドレス3))
    AuthMACアドレス認証チェック = result
    If result Then Exit Function
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


