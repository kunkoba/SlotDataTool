VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_FC9_ÉfÅ[É^íPï[sub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


' ======================================================
Private Sub Form_Load()
On Error Resume Next
    Form.caption = App_ÉAÉvÉäñº
    Me.KeyPreview = True
    
End Sub
' ======================================================
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    If KeyCode = vbKeyEscape Then
        DoCmd.Close acForm, Me.name
        DoCmd.Close acForm, Me.Parent.name
    End If
    
End Sub




' ======================================================
Private Sub textâÒé˚äz_AfterUpdate()
On Error Resume Next
    texté˚éxäz.Requery
    
End Sub

' ======================================================
Private Sub textìäéëäz_AfterUpdate()
On Error Resume Next
    texté˚éxäz.Requery
    
End Sub
