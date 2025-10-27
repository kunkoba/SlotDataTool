VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F01_最適化"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Dim frmName As String
    
' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Me.TimerInterval = 1000
    
'    Dim args As Variant
'
'    If Not IsNull(Me.OpenArgs) Then
'        args = Split(Me.OpenArgs, vbTab)
'        frmName = args(0)
'    End If
    
End Sub



' ======================================================
Private Sub Form_Timer()
On Error Resume Next
    Call Procファイル最適化
    flg最適化 = False
'
'    DoCmd.OpenForm frmName
    DoCmd.Close acForm, Me.name

End Sub
