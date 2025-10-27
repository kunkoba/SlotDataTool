Attribute VB_Name = "_License"
Option Compare Database
Option Explicit



Const KEY_LEN As Integer = 15  ' 生成する文字数
Const CODE_LEN_1 As Integer = 5 ' 数字の長さ
Const CODE_LEN_2 As Integer = 10 ' 数字の長さ
    
    
' **************************************************************
' アプリ制限解除
' **************************************************************
Public Sub TESTアプリ制限解除()
    Dim key1 As String
'    Dim key2 As String
'    Dim key_tmp As String
    Dim code1 As String
    Dim code2 As String

    Debug.Print Now
    
    Debug.Print "----　秘密コード生成　----"
    key1 = Lic秘密コード生成

    Debug.Print "----　暗号キー生成　----"
    code1 = Lic秘密コードから暗号キー(key1)

    Debug.Print "----　解除コード生成　----"
    code2 = Lic暗号キーから解除コード(code1)
    code2 = str指定文字数ごとに指定文字を挟む(code2, "-", 5)
    

    Debug.Print "----　結果　----"
    Debug.Print Lic解除コードチェック(code1, code2)
    
    Debug.Print

End Sub



' **************************************************************
' 解除コードチェック（暗号キーと比較）
' **************************************************************
Public Function Lic解除コードチェック(暗号キー As String, 解除コード As String) As String
    Const methodName As String = "Lic解除コードチェック"
    Dim key1 As String, key2 As String
On Error GoTo ErrHandler

    Lic解除コードチェック = False

    Debug.Print "Lic解除コードチェック1 >> ", 暗号キー, 解除コード
    
    解除コード = Replace(解除コード, "-", "")   'ハイフン除外
    
    If Len(解除コード) <> KEY_LEN + CODE_LEN_2 Then GoTo ErrHandler
    
    key1 = Lic指定文字列から秘密コード(暗号キー)
    key2 = Lic指定文字列から秘密コード(解除コード)
    
    Debug.Print "Lic解除コードチェック2-> ", key1, key2
    Lic解除コードチェック = (str文字列を昇順ソート(key1) = str文字列を昇順ソート(key2))
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function



' **************************************************************
' **************************************************************
' **************************************************************
' **************************************************************
' **************************************************************
' 秘密コード(10)　→　暗号キー(20)　→　秘密コード(10)　→　解除コード(25)
' **************************************************************
' ランダム英字生成
' **************************************************************
Public Function Lic秘密コード生成() As String
    Const methodName As String = "Lic秘密コード生成"
    Dim code1 As String, code2 As String
    Dim i As Integer
    Dim result As String
    Dim rndChar As Integer
On Error GoTo ErrHandler

    Randomize ' 乱数初期化
    
    result = ""
    For i = 1 To KEY_LEN
        ' A～Z (ASCIIコード 65～90)
        rndChar = Int((26 * Rnd) + 65)
        result = result & Chr(rndChar)
    Next i
    
    Lic秘密コード生成 = result
    Debug.Print "Lic秘密コード生成 >> " & result
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


' **************************************************************
Public Function Lic秘密コードから暗号キー(ByVal key As String) As String
On Error Resume Next
    Lic秘密コードから暗号キー = Lic秘密コードから指定文字分生成(key, CODE_LEN_1)
    
End Function


' **************************************************************
Public Function Lic暗号キーから解除コード(ByVal code As String) As String
    Dim temp As String
On Error Resume Next
    temp = Lic指定文字列から秘密コード(code)
    temp = Lic秘密コードから指定文字分生成(temp, CODE_LEN_2)
    Lic暗号キーから解除コード = str指定文字数ごとに指定文字を挟む(temp, "-", 5)
    
End Function



' **************************************************************
' 英字をカモフラージュする
' **************************************************************
Private Function Lic秘密コードから指定文字分生成(ByVal src As String, codeNum As Integer) As String
    Const methodName As String = "Lic秘密コードから指定文字分生成"
    Dim i As Integer
    Dim ch As String
    Dim result As String
    Dim rndNum As String
    Dim mixStr As String
    Dim shuffled As String
On Error GoTo ErrHandler

    Randomize
    
    '--- 英字をランダムで小文字化 ---
    result = ""
    For i = 1 To Len(src)
        ch = Mid(src, i, 1)
        If Rnd < 0.5 Then
            result = result & LCase(ch)
        Else
            result = result & UCase(ch)
        End If
    Next i
    
    '--- ランダム数値を生成 ---
    rndNum = ""
    For i = 1 To codeNum
        rndNum = rndNum & CStr(Int((10 * Rnd)))
    Next i
    
    '--- 英字＋数字を結合 ---
    mixStr = result & rndNum
    
    '--- ランダム並び替え ---
    shuffled = str文字列をバラバラソート(mixStr)
    
    Lic秘密コードから指定文字分生成 = shuffled
    
    Debug.Print "Lic秘密コードから指定文字分生成 >> " & shuffled
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function

' **************************************************************
' ランダム生成された文字をもとに戻す
' **************************************************************
Private Function Lic指定文字列から秘密コード(ByVal camo As String) As String
    Const methodName As String = "Lic指定文字列から秘密コード"
    Dim i As Integer
    Dim ch As String
    Dim result As String
On Error GoTo ErrHandler
    
    Lic指定文字列から秘密コード = "X"
    
    ' 1文字ずつチェック
    For i = 1 To Len(camo)
        ch = Mid(camo, i, 1)
        
        ' 数字を除外して英字だけを取る
        If ch Like "[A-Za-z]" Then
            result = result & UCase(ch)
        End If
    Next i
    
    Lic指定文字列から秘密コード = result
    Debug.Print "Lic指定文字列から秘密コード >> " & result
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function




' **************************************************************
' 文字列をランダムに並び替えるヘルパー関数
' **************************************************************
Private Function str文字列をバラバラソート(ByVal src As String) As String
    Const methodName As String = "str文字列をバラバラソート"
    Dim i As Integer, j As Integer
    Dim arr() As String
    Dim temp As String
    Dim result As String
On Error GoTo ErrHandler
    
    ReDim arr(1 To Len(src))
    
    ' 文字を配列に格納
    For i = 1 To Len(src)
        arr(i) = Mid(src, i, 1)
    Next i
    
    ' Fisher?Yates shuffle
    For i = Len(src) To 2 Step -1
        j = Int(i * Rnd) + 1
        temp = arr(i)
        arr(i) = arr(j)
        arr(j) = temp
    Next i
    
    ' 配列を文字列に戻す
    result = ""
    For i = 1 To Len(src)
        result = result & arr(i)
    Next i
    
    str文字列をバラバラソート = result
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function

' **************************************************************
' 文字列を昇順で並び替える関数
' **************************************************************
Private Function str文字列を昇順ソート(ByVal src As String) As String
    Const methodName As String = "str文字列を昇順ソート"
    Dim arr() As String
    Dim i As Long, j As Long
    Dim temp As String
    Dim result As String
On Error GoTo ErrHandler
    
    ' 文字列を配列に分解
    ReDim arr(1 To Len(src))
    For i = 1 To Len(src)
        arr(i) = Mid(src, i, 1)
    Next i
    
    ' バブルソート（単純ソート）
    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' 配列を結合
    result = ""
    For i = 1 To UBound(arr)
        result = result & arr(i)
    Next i
    
    str文字列を昇順ソート = result
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）

End Function


' **************************************************
Private Function str指定文字数ごとに指定文字を挟む(src As String, strChar As String, num As Integer) As String
    Const methodName As String = "str指定文字数ごとに指定文字を挟む"
    Dim i As Long
    Dim result As String
On Error GoTo ErrHandler
    
    For i = 1 To Len(src)
        result = result & Mid(src, i, 1)
        ' 指定数ごとにハイフンを追加（末尾以外）
        If i Mod num = 0 And i <> Len(src) Then
            result = result & strChar
        End If
    Next i
    
    str指定文字数ごとに指定文字を挟む = result
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function



'' ==============================================================
'' 文字列全体をシフト
'' ==============================================================
'Private Function strシフト文字列(ByVal s As String, ByVal Shift As Integer) As String
'    Dim i As Long
'    Dim result As String
'    result = ""
'
'    For i = 1 To Len(s)
'        result = result & strシフト文字(Mid$(s, i, 1), Shift)
'    Next i
'
'    strシフト文字列 = result
'
'End Function
'' ==============================================================
'' 補助：文字をシフト
'' ==============================================================
'Private Function strシフト文字(ByVal ch As String, ByVal Shift As Integer) As String
'    Dim code As Integer
'    code = Asc(ch)
'
'    ' 0-9 -> 48-57
'    If code >= 48 And code <= 57 Then
'        strシフト文字 = Chr(((code - 48 + Shift) Mod 10) + 48)
'        Exit Function
'    End If
'
'    ' A-Z -> 65-90
'    If code >= 65 And code <= 90 Then
'        strシフト文字 = Chr(((code - 65 + Shift) Mod 26) + 65)
'        Exit Function
'    End If
'
'    ' a-z -> 97-122
'    If code >= 97 And code <= 122 Then
'        strシフト文字 = Chr(((code - 97 + Shift) Mod 26) + 97)
'        Exit Function
'    End If
'
'    ' それ以外はそのまま返す
'    strシフト文字 = ch
'
'End Function






