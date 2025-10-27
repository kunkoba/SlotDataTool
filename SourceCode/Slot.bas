Attribute VB_Name = "Slot"
Option Compare Database
Option Explicit


' ***************************************************************
Public Function クレンジング処理(ByVal src As Variant) As Variant
    Const methodName As String = "クレンジング処理"
    ' 宣言部
    Dim s As String
On Error GoTo ErrHandler
    
    クレンジング処理 = ""
    
    ' メイン処理
    If IsNull(src) Then
        Exit Function
    End If

    s = CStr(src)
    
    ' 全角＋ → 削除
    s = Replace(s, "＋", "")
    ' 半角+ → 削除
    s = Replace(s, "+", "")
    ' カンマ , → 削除
    s = Replace(s, ",", "")
    
    クレンジング処理 = s
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 設定判別_BB
' ***************************************************************
Public Function 汎用設定判別(ゲーム数 As Long, 成立回数 As Long, _
                          設定値1 As Double, 設定値2 As Double, 設定値3 As Double, _
                          設定値4 As Double, 設定値5 As Double, 設定値6 As Double) As Variant
    Const methodName As String = "汎用設定判別"
    ' 宣言部
    Dim 実測分母 As Double
    Dim 設定値(1 To 6) As Double
    Dim i As Integer
    Dim val1 As Double, val2 As Double
    Dim 設定推定値 As Double
On Error GoTo ErrHandler

    汎用設定判別 = 1
    
    ' メイン処理
    If 成立回数 <= 0 Then
        Exit Function
    End If

    実測分母 = ゲーム数 / 成立回数

    ' 設定値配列化
    設定値(1) = 設定値1
    設定値(2) = 設定値2
    設定値(3) = 設定値3
    設定値(4) = 設定値4
    設定値(5) = 設定値5
    設定値(6) = 設定値6

    ' 線形補間
    For i = 1 To 5
        val1 = 設定値(i)
        val2 = 設定値(i + 1)
        If val1 > 0 And val2 > 0 Then
            If (実測分母 >= val1 And 実測分母 <= val2) Or (実測分母 >= val2 And 実測分母 <= val1) Then
                設定推定値 = i + (実測分母 - val1) / (val2 - val1)
                汎用設定判別 = Round(設定推定値, 1)
                Exit Function
            End If
        End If
    Next i

    ' 最も近い値にする（補間範囲外）
    Dim 最小差 As Double: 最小差 = 999999
    Dim 最小設定 As Integer
    For i = 1 To 6
        If 設定値(i) > 0 Then
            If Abs(実測分母 - 設定値(i)) < 最小差 Then
                最小差 = Abs(実測分母 - 設定値(i))
                最小設定 = i
            End If
        End If
    Next i

    汎用設定判別 = Round(最小設定, 1)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 設定総合評価
' ***************************************************************
Public Function 設定総合評価(設定_BB As Variant, 設定_RB As Variant, 設定_合算 As Variant, _
                            wBB As Variant, wRB As Variant, w合算 As Variant, _
                            ゲーム数 As Long, 回転数_しきい値 As Long, 設定調整値 As Integer) As Variant
    Const methodName As String = "設定総合評価"
    ' 宣言部
    Dim ret As Variant
    Dim 評価値 As Double
On Error GoTo ErrHandler

    設定総合評価 = Null
    
    ' メイン処理
    If IsNull(設定_BB) Then
        Exit Function
    End If

    評価値 = 設定_BB * wBB + 設定_RB * wRB + 設定_合算 * w合算
    設定総合評価 = Round(評価値 / (wBB + wRB + w合算), 1)
    
    
    ' 設定調整
    If ゲーム数 < 回転数_しきい値 Then
        設定総合評価 = 設定総合評価 - 設定調整値
        If 設定総合評価 < 1 Then 設定総合評価 = 1
    End If
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 機械割差枚数
' ***************************************************************
Public Function 機械割差枚数(ゲーム数 As Long, 想定設定 As Double, _
                    割1 As Double, 割2 As Double, 割3 As Double, _
                    割4 As Double, 割5 As Double, 割6 As Double) As Long
    Const methodName As String = "機械割差枚数"
    ' 宣言部
    Dim 設定下 As Integer, 設定上 As Integer
    Dim 差枚 As Double
    Dim 割補間 As Double
    Dim arr割 As Variant
On Error GoTo ErrHandler

    機械割差枚数 = 0
    
    ' メイン処理
    If ゲーム数 <= 0 Or 想定設定 < 1 Or 想定設定 > 6 Then
        機械割差枚数 = 0
        Exit Function
    End If

    ' 機械割を配列に（添字1～6）
    arr割 = Array(0, 割1, 割2, 割3, 割4, 割5, 割6)

    設定下 = Int(想定設定)
    If 設定下 + 1 > 6 Then
        設定上 = 6
    Else
        設定上 = 設定下 + 1
    End If

    If 設定下 = 設定上 Then
        割補間 = arr割(設定下)
    Else
        割補間 = arr割(設定下) + (arr割(設定上) - arr割(設定下)) * (想定設定 - 設定下)
    End If

    差枚 = ゲーム数 * 3 * (割補間 / 100 - 1)
    機械割差枚数 = CLng(差枚)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 想定差枚数
' ***************************************************************
Public Function 想定差枚数(ゲーム数 As Long, BB数 As Long, RB数 As Long, BB枚数 As Double, RB枚数 As Double, 差枚数係数 As Double) As Long
    Const methodName As String = "想定差枚数"
    ' 宣言部
    Dim 差枚数 As Double
On Error GoTo ErrHandler

    想定差枚数 = 0

    差枚数 = BB数 * BB枚数 + RB数 * RB枚数 + ゲーム数 * 差枚数係数 - ゲーム数 * 3

    想定差枚数 = Round(差枚数)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 判定結果信頼度
' ***************************************************************
Public Function 判定結果信頼度(ゲーム数 As Long) As String
    Const methodName As String = "判定結果信頼度"
    ' 宣言部
On Error GoTo ErrHandler

    判定結果信頼度 = ""
    
    ' メイン処理
    If ゲーム数 < 2000 Then
        判定結果信頼度 = "1低"
    
    ElseIf ゲーム数 < 5000 Then
        判定結果信頼度 = "2中"
    
    Else
        判定結果信頼度 = "3高"
    
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 出玉傾向
' ***************************************************************
Public Function 出玉傾向(有効差枚 As Long, 機械割差枚 As Long, しきい値 As Long) As String
    Const methodName As String = "出玉傾向"
    ' 宣言部
    Dim diff As Long
On Error GoTo ErrHandler
    
    出玉傾向 = "想定内→"
        
    ' メイン処理
    If 機械割差枚 = 0 Then
        Exit Function
    End If

    diff = Abs(有効差枚 - 機械割差枚)

    ' 評価ロジック
    If diff <= しきい値 Then
        出玉傾向 = "想定内→"
    ElseIf 有効差枚 > 機械割差枚 Then
        出玉傾向 = "上振れ↑"
    Else
        出玉傾向 = "下振れ↓"
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function



' ******************************************************
' 汎用：指定の変換対象に応じて値を変換する（エラー時 -1、文字列は Trim）
' ******************************************************
Public Function 値変換(sDate As Variant, convType As String) As Variant
    Const methodName As String = "値変換"
    ' 宣言部
    Dim result As Variant
On Error GoTo ErrHandler

    値変換 = ""
    
    ' メイン処理
    If sDate = "" Or sDate = Null Then
        Exit Function
    End If

    Select Case LCase(convType)
        Case "末尾"
            result = Right(CStr(sDate), 1)
            
        Case "日付末尾"
            result = Right(CStr(sDate), 1)
        Case "台番号末尾"
            result = Right(CStr(sDate), 1)
        Case "年月"
            result = Format(sDate, "yyyymm")
        Case "月"
            result = Month(sDate)
        Case "曜日"
            sDate = CDate(sDate)
            result = Weekday(sDate) & Format(sDate, "aaa")
        Case "月旬"
            result = 月旬変換(CDate(sDate)) ' ※別途 月旬 関数が必要
        Case "月旬6"
            result = 月旬6分割変換(CDate(sDate)) ' ※別途 月旬6分割 関数が必要
        Case "日"
            result = day(sDate)
    End Select

    ' 文字列なら Trim、数値ならそのまま返す
    If VarType(result) = vbString Then
        値変換 = Trim(result)
    Else
        値変換 = result
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' 月旬（上・中・下）を算出
' ***************************************************************
Public Function 月旬変換(日付 As Date) As String
    Const methodName As String = "月旬変換"
    ' 宣言部
    Dim 日 As Integer
On Error GoTo ErrHandler

    月旬変換 = ""
    
    ' メイン処理
    日 = day(日付)

    If 日 <= 10 Then
        月旬変換 = "1上旬"
    ElseIf 日 <= 20 Then
        月旬変換 = "2中旬"
    Else
        月旬変換 = "3下旬"
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function

' ***************************************************************
' 月旬（5日ごと6分割）を算出
' ***************************************************************
Public Function 月旬6分割変換(日付 As Date) As String
    Const methodName As String = "月旬6分割変換"
    ' 宣言部
    Dim 日 As Integer
On Error GoTo ErrHandler

    月旬6分割変換 = ""
    
    ' メイン処理
    日 = day(日付)
    
    Dim 区分 As Integer
    区分 = Int((日 - 1) / 5) + 1

    If 区分 > 6 Then 区分 = 6 ' 安全処理（最大でも第6旬）

    月旬6分割変換 = "第" & 区分 & "旬"

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' ボーナス当選率
' 　引数二つなら合算値
' ***************************************************************
Function ボーナス当選率(ByVal bonus1 As Long, ByVal ゲーム数 As Long, Optional ByVal bonus2 As Long = 0) As String
    Const methodName As String = "ボーナス当選率"
    ' 宣言部
    Dim totalBonus As Long
    Dim rate As Double
    Dim denominator As Long
On Error GoTo ErrHandler
    
    ボーナス当選率 = "1/999"
        
    ' メイン処理
    If ゲーム数 <= 0 Then
        Exit Function
    End If

    totalBonus = bonus1 + bonus2

    If totalBonus <= 0 Then
        ボーナス当選率 = "1/999"
    Else
        rate = totalBonus / ゲーム数
        denominator = CLng(1 / rate)

        If denominator >= 1000 Then
            ボーナス当選率 = "1/999"
        Else
            ボーナス当選率 = "1/" & Format(denominator, "000")  ' 3桁固定（例: 1/007, 1/250）
        End If
    End If
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
Public Function 当選率算出(ByVal ゲーム数 As Long, ByVal ボーナス数 As Long) As String
    Const methodName As String = "当選率算出"
    Dim rate As Double
On Error GoTo ErrHandler
    
    当選率算出 = "--"
    
    ' BB数が0の場合
    If ボーナス数 = 0 Then
        Exit Function
    End If
    
    ' 当選率計算
    rate = ゲーム数 / ボーナス数
    
    ' 四捨五入して整数化
    rate = Round(rate, 0)
    
    ' 分子（計算結果）が1000以上なら999に制限
    If rate >= 1000 Then
        rate = 999
    End If
    
    ' 文字列化して返却
    当選率算出 = "1/" & CStr(rate)
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function



