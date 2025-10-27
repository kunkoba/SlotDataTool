Attribute VB_Name = "ShowMessage"
Option Compare Database
Option Explicit


' ***********************************************
' トースト表示（mode指定付き）
' ***********************************************
Public Sub ShowToast(msg As String, Optional modeColor As Integer = 0, Optional interval As Integer = 30000)
    Const methodName As String = "ShowToast"
On Error GoTo ErrHandler

    DoCmd.OpenForm F03_Toast, acNormal, , , , acWindowNormal

    ' トースト設定
    With Forms(F03_Toast)
        .lblMessage.caption = msg
        .TimerInterval = interval
        
        Select Case modeColor
            Case E_Color_Black ' 黒
                .boxBackBoard.BackColor = RGB(0, 0, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Blue ' 青
                .boxBackBoard.BackColor = RGB(0, 122, 204)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Red ' 赤
                .boxBackBoard.BackColor = RGB(204, 0, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Green ' 緑
                .boxBackBoard.BackColor = RGB(0, 153, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
        End Select
        
        .Repaint
    End With

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ***********************************************
' 伝達ダイアログ生成
' ***********************************************
Public Sub ShowInfomation(title As String, msg As String)
    Const methodName As String = "ShowInfomation"
    Dim frm As String
    Dim ret As Integer
    Dim args As String
On Error GoTo ErrHandler
    
    args = title & vbTab & msg    ' 区切り付きで渡す（簡易）

    frm = F12_InfomationDialog
    DoCmd.OpenForm frm, , , , , acDialog, args    'await処理

    ret = Nz(Forms(frm).選択結果, 0)
    DoCmd.Close acForm, frm    'ここで閉じる
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Sub


' ***********************************************
' 確認ダイアログ生成
' ***********************************************
Public Function ShowConfirm(title As String, msg As String, Optional mode As Integer = vbYesNo) As Integer
    Const methodName As String = "ShowConfirm"
    Dim frm As String
    Dim ret As Integer
    Dim args As String
On Error GoTo ErrHandler
    
    args = title & vbTab & msg & vbTab & mode   ' 区切り付きで渡す（簡易）

    frm = F13_ConfirmDialog
    DoCmd.OpenForm frm, , , , , acDialog, args    'await処理

    ret = Nz(Forms(frm).選択結果, 0)
    DoCmd.Close acForm, frm    'ここで閉じる

    ShowConfirm = ret

ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "疑似エラー"  ' 必ず最後(後続処理が動かないため）
    
End Function




