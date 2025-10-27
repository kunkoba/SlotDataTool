VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_P_タイトル"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim blnOverButton As Boolean

' ======================================================
Private Sub Form_Load()
    Dim flg As Boolean
On Error Resume Next
    
    Call Procアプリ状態更新
    
    btn閉じる.Visible = True
    If Me.Parent.name = F11_MainMenu Then btn閉じる.Visible = False
    
End Sub
' ======================================================
Private Sub Form_Close()
    
    
End Sub
' ======================================================
Private Sub btn閉じる_Click()
On Error Resume Next
    DoCmd.Close acForm, Me.name
    DoCmd.Close acForm, Me.Parent.name

End Sub


' ======================================================
Public Sub Procアプリ状態更新()
    Const methodName As String = "Procアプリ状態更新"
    Dim flg As Boolean
On Error GoTo ErrHandler

    ' 制限解除チェック
    flg = Not IsNull(App_アプリ解除日)
    lab無料版.Visible = Not flg
    lab有料版.Visible = flg
    
ErrHandler:
    Call ErrorSave(methodName, True) '必ず先頭（発生したエラーを確実にキャッチするため）
    
End Sub



' ******************************************************
Public Sub Procタイトル更新(caption As String, mode As Integer)
    Const methodName As String = "Procタイトル更新"
On Error Resume Next
    labタイトル.caption = caption
    
    Select Case mode
    Case E_Color_Blue
        labタイトル.BackColor = ConvertHexColorToRGB("#4BACC6") '青っぽい（取り込み）
    Case E_Color_DeepBlue
        labタイトル.BackColor = ConvertHexColorToRGB("#56769D") '深青っぽい（集計）
    Case E_Color_Green
        labタイトル.BackColor = ConvertHexColorToRGB("#7c9647") '緑っぽい（マスタ）
    Case E_Color_Red
        labタイトル.BackColor = ConvertHexColorToRGB("#CF7B79") '赤っぽい（削除）
    Case E_Color_Black
        labタイトル.BackColor = ConvertHexColorToRGB("#000000") '黒（システム）
    Case E_Color_Purple
        labタイトル.BackColor = ConvertHexColorToRGB("#4B0082") '濃いパープル（グラフ）
    End Select
    
    '確認ダイアログを確実に閉じたい
    DoCmd.Close acForm, F13_ConfirmDialog    '念のため
    DoCmd.Close acForm, F12_InfomationDialog       '念のため
    
ErrHandler:
    Call ErrorSave(methodName)  '必ず先頭（発生したエラーを確実にキャッチするため）
    
End Sub

