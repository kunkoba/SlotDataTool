Attribute VB_Name = "CommonForm"
Option Compare Database
Option Explicit


' ======================================================
Public Sub Proc�N���b�v�{�[�h�֓]��(ctrl As TextBox)
    Const methodName As String = "Proc�N���b�v�{�[�h�֓]��"
On Error GoTo ErrHandler

    ' ���C������
    With ctrl
        ' �R�s�[�������R���g���[���i��F�e�L�X�g0�j�Ƀt�H�[�J�X���ړ�
        .SetFocus
        
        ' �e�L�X�g�{�b�N�X���̓��e�S�̂�I��
        .SelStart = 0
        .SelLength = Len(.text)
        
        DoCmd.RunCommand acCmdCopy
    End With
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub



'' ******************************************************
'' �R���{�{�b�N�X�̑I�����ړ�������i�㉺�ړ��A���o���Ή��j
'' ******************************************************
'Public Sub changeCombbox(cmb As ComboBox, addValue As Long, Optional isHeader As Boolean = True)
'    Const methodName As String = "changeCombbox"
'    ' �錾��
'    Dim idx As Long
'    Dim newIdx As Long
'    Dim offset As Long
'On Error GoTo ErrHandler
'
'    ' ���C������
'    offset = IIf(isHeader, 1, 0)
'
'    ' ���݂̃C���f�b�N�X
'    idx = cmb.ListIndex + offset
'    If idx < 0 Then Exit Sub  ' ���I��
'
'    newIdx = idx + addValue
'
'    ' �͈̓`�F�b�N�i�w�b�_�[�l���j
'    If newIdx < offset Then Exit Sub
'    If newIdx >= cmb.ListCount Then Exit Sub
'
'    ' �I��ύX
'    cmb.value = cmb.ItemData(newIdx)
'
'ErrHandler:
'    ' �G���[
'    If Err.Number <> 0 Then
'        Call LogError(methodName)
'        Err.Raise Err.Number, Err.source, Err.Description
'    End If
'ExitHandler:
'    ' �I������
''
'
'End Sub



' **************************************************
' �T�u�t�H�[���X�V�iFilter / OrderBy ��ێ��j
' subFormControl: �T�u�t�H�[���R���g���[�����i������j
' **************************************************
Public Sub Proc�t�B���^����ێ������܂܃T�u�t�H�[���X�V(frm As Form, subFormControl As String)
    Const methodName As String = "Proc�t�B���^����ێ������܂܃T�u�t�H�[���X�V"
    ' �錾��
    Dim sf As subForm
    Dim ctrl As subForm
    Dim currentFilter As String
    Dim currentOrder As String
On Error GoTo ErrHandler
    
    ' �T�u�t�H�[���R���g���[�����擾
    Set ctrl = frm.Controls(subFormControl)
    
    ' ���݂̃t�B���^�E�I�[�_�[��ޔ�
    currentFilter = ctrl.Form.filter
    currentOrder = ctrl.Form.OrderBy
    
    ' SourceObject ���Đݒ�i�X�V�j
    ctrl.SourceObject = ctrl.SourceObject
    
    ' �t�B���^�E�I�[�_�[�𕜌�
    ctrl.Form.filter = currentFilter
    ctrl.Form.OrderBy = currentOrder
    
    ' �t�B���^�E�I�[�_�[�̗L����Ԃ�����
    If currentFilter <> "" Then ctrl.Form.FilterOn = True
    If currentOrder <> "" Then ctrl.Form.OrderByOn = True

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ******************************************************
' ���X�g�{�b�N�X�̒l�\�[�X���ꌳ�z��ɕϊ�����
' ******************************************************
Function Proc���X�g�{�b�N�X�̒l�擾�z��(lst As ListBox, Optional delimiter As String = ";") As Variant
    Const methodName As String = "Proc���X�g�{�b�N�X�̒l�擾�z��"
    ' �錾��
    Dim raw() As String
    Dim totalItems As Long
    Dim i As Long
    Dim arr1D() As String
On Error GoTo ErrHandler
    
    ' RowSource ����Ȃ��z���Ԃ�
    If lst.RowSource = "" Then
        Proc���X�g�{�b�N�X�̒l�擾�z�� = Array()
        Exit Function
    End If
    
    ' RowSource �� delimiter �ŕ���
    raw = Split(lst.RowSource, delimiter)
    totalItems = UBound(raw) - LBound(raw) + 1
    
    ' 1�����z��ɋl�߂�
    ReDim arr1D(0 To totalItems - 1)
    For i = 0 To totalItems - 1
        arr1D(i) = raw(i)
    Next i
    
    Proc���X�g�{�b�N�X�̒l�擾�z�� = arr1D

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



'=============================================
' ���X�g�{�b�N�X�̑I���s�́u�A����v�̒l��z��
' lst : ListBox �R���g���[��
'=============================================
Public Function Proc�I�����ꂽ�v�f���擾�z��(lst As Access.ListBox, Optional is�����p As Boolean = False) As Variant
    Const methodName As String = "Proc�I�����ꂽ�v�f���擾�z��"
    ' �錾��
    Dim i As Long
    Dim tmp As Collection
    Dim result() As Variant
On Error GoTo ErrHandler
    
    ' ���C������
    Set tmp = New Collection
    
    ' �I������Ă���s�́uValue�v= �A����̒l���E��
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) Then
            tmp.Add lst.ItemData(i)   ' �� ���ꂪ�A����̒l
        End If
    Next i
    
    ' �z��
    If tmp.count > 0 Then
        ReDim result(0 To tmp.count - 1)
        For i = 1 To tmp.count
            If is�����p Then
                result(i - 1) = Convert�����l������(tmp(i))
            Else
                result(i - 1) = tmp(i)
            End If
        Next i
        Proc�I�����ꂽ�v�f���擾�z�� = result
    Else
        Proc�I�����ꂽ�v�f���擾�z�� = Array()
    End If

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



'=============================================
' ���X�g�{�b�N�X�ɍ��ڂ�ǉ�����֐�
'=============================================
Public Sub Proc���X�g�{�b�N�X�v�f�ǉ�(ByVal newItem As String, lst As Access.ListBox)
    Const methodName As String = "Proc���X�g�{�b�N�X�v�f�ǉ�"
    ' �錾��
On Error GoTo ErrHandler

    With lst
        If Len(.RowSource) > 0 Then
            .RowSource = .RowSource & ";" & newItem
        Else
            .RowSource = newItem
        End If
    End With
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


'=============================================
' ���X�g�{�b�N�X���בւ�
'=============================================
Public Sub Proc���X�g�{�b�N�X���בւ�(lst As ListBox)
    Const methodName As String = "Proc���X�g�{�b�N�X�v�f�ǉ�"
    ' �錾��
    Dim arr() As String
    Dim i As Long
    Dim itemCount As Long
On Error GoTo ErrHandler
    
    ' ���X�g�{�b�N�X�̍s���擾
    itemCount = lst.ListCount
    If itemCount = 0 Then Exit Sub
    
    ' �z��ɗ�1�̒l���i�[
    ReDim arr(0 To itemCount - 1)
    For i = 0 To itemCount - 1
        arr(i) = lst.Column(0, i) ' ��1��Column(0)
    Next i
    
    ' �z��������\�[�g�i�o�u���\�[�g�j
    Dim j As Long
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
    
    ' ���X�g�{�b�N�X���N���A���čĒǉ�
    lst.RowSourceType = "Value List"
    lst.RowSource = ""
    For i = LBound(arr) To UBound(arr)
        If lst.RowSource = "" Then
            lst.RowSource = arr(i)
        Else
            lst.RowSource = lst.RowSource & ";" & arr(i)
        End If
    Next i
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' **************************************************
' �����ȊO�̂��ׂẴt�H�[�������
' **************************************************
Public Sub Proc���ׂẴt�H�[�������(Optional keepFormName As String = "")
    Const methodName As String = "Proc���ׂẴt�H�[�������"
    ' �錾��
    Dim frm As AccessObject
    Dim target As String
On Error GoTo ErrHandler
    
    ' �f�t�H���g�F�Ăяo�����̃t�H�[�����c��
    If keepFormName = "" Then
        If Screen.ActiveForm Is Nothing Then
            keepFormName = ""
        Else
            keepFormName = Screen.ActiveForm.name
        End If
    End If
    
    For Each frm In CurrentProject.AllForms
        target = frm.name
        If CurrentProject.AllForms(target).IsLoaded Then
            If target <> keepFormName Then
                DoCmd.Close acForm, target, acSaveNo
            End If
        End If
    Next

ErrHandler:
    DoCmd.Echo True
    Call ErrorSave(methodName, True) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' **************************************************
' �����ȊO�̂��ׂẴt�H�[�������
' **************************************************
Public Sub Proc���X�g�{�b�N�X�����I������(ctrl As ListBox)
On Error Resume Next
    Dim i As Long
    With ctrl
        For i = 0 To .ListCount - 1
            .Selected(i) = False
        Next i
        If .ListCount > 0 Then
            .Selected(0) = True
            .Selected(0) = False
        End If
    End With
    
End Sub


