Attribute VB_Name = "ShowMessage"
Option Compare Database
Option Explicit


' ***********************************************
' �g�[�X�g�\���imode�w��t���j
' ***********************************************
Public Sub ShowToast(msg As String, Optional modeColor As Integer = 0, Optional interval As Integer = 30000)
    Const methodName As String = "ShowToast"
On Error GoTo ErrHandler

    DoCmd.OpenForm F03_Toast, acNormal, , , , acWindowNormal

    ' �g�[�X�g�ݒ�
    With Forms(F03_Toast)
        .lblMessage.caption = msg
        .TimerInterval = interval
        
        Select Case modeColor
            Case E_Color_Black ' ��
                .boxBackBoard.BackColor = RGB(0, 0, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Blue ' ��
                .boxBackBoard.BackColor = RGB(0, 122, 204)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Red ' ��
                .boxBackBoard.BackColor = RGB(204, 0, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
                
            Case E_Color_Green ' ��
                .boxBackBoard.BackColor = RGB(0, 153, 0)
                .lblMessage.ForeColor = RGB(255, 255, 0)
        End Select
        
        .Repaint
    End With

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ***********************************************
' �`�B�_�C�A���O����
' ***********************************************
Public Sub ShowInfomation(title As String, msg As String)
    Const methodName As String = "ShowInfomation"
    Dim frm As String
    Dim ret As Integer
    Dim args As String
On Error GoTo ErrHandler
    
    args = title & vbTab & msg    ' ��؂�t���œn���i�ȈՁj

    frm = F12_InfomationDialog
    DoCmd.OpenForm frm, , , , , acDialog, args    'await����

    ret = Nz(Forms(frm).�I������, 0)
    DoCmd.Close acForm, frm    '�����ŕ���
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ***********************************************
' �m�F�_�C�A���O����
' ***********************************************
Public Function ShowConfirm(title As String, msg As String, Optional mode As Integer = vbYesNo) As Integer
    Const methodName As String = "ShowConfirm"
    Dim frm As String
    Dim ret As Integer
    Dim args As String
On Error GoTo ErrHandler
    
    args = title & vbTab & msg & vbTab & mode   ' ��؂�t���œn���i�ȈՁj

    frm = F13_ConfirmDialog
    DoCmd.OpenForm frm, , , , , acDialog, args    'await����

    ret = Nz(Forms(frm).�I������, 0)
    DoCmd.Close acForm, frm    '�����ŕ���

    ShowConfirm = ret

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function




