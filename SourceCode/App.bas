Attribute VB_Name = "App"
Option Compare Database
Option Explicit


Public Const T_AppSetting As String = "S_�A�v���ݒ�"


' **************************************************
Public Function App_�ڑ��`�F�b�N()
    Const methodName As String = "App_�ڑ��`�F�b�N"
    Dim newPath As String
'    Dim flg As Boolean
On Error Resume Next

    ' �K��̃f�[�^�t�@�C���ɐڑ�����i�z����Data��D�悵�Đڑ�����j
    newPath = CurrentProject.path & "\" & PATH_DATA & "\" & App_�f�[�^�t�@�C���� & SYS�g���q
    If Proc�t�@�C�����݊m�F(newPath) Then
        ' ����̏ꏊ�Ƀf�[�^�t�@�C��������΁A�����N����X�V����i������΍X�V���Ȃ��j
        Call Proc�����N�e�[�u���ꊇ�X�V(newPath)
    
    End If

    App_�ڑ��`�F�b�N = DCount("*", "M_�X�܃}�X�^") > 0
    If Not App_�ڑ��`�F�b�N Then Call ShowConfirm("�V�X�e��", "�f�[�^�t�@�C����������܂���B�@�ēx�A�ݒ�����Ă��������B", vbYes)
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function

' **************************************************
Public Function App_�A�v����() As String
On Error Resume Next
    App_�A�v���� = DLookup("�A�v����", T_AppSetting)
        
End Function
' **************************************************
Public Function App_���C�Z���X�}�[�N() As String
On Error Resume Next
    App_���C�Z���X�}�[�N = DLookup("���C�Z���X�}�[�N", T_AppSetting)
        
End Function
' **************************************************
Public Function App_�o�[�W����() As String
On Error Resume Next
    App_�o�[�W���� = DLookup("�o�[�W����", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MAC�A�h���X1() As String
On Error Resume Next
    App_MAC�A�h���X1 = DLookup("MAC�A�h���X1", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MAC�A�h���X2() As String
On Error Resume Next
    App_MAC�A�h���X2 = DLookup("MAC�A�h���X2", T_AppSetting)
        
End Function
' **************************************************
Public Function App_MAC�A�h���X3() As String
On Error Resume Next
    App_MAC�A�h���X3 = DLookup("MAC�A�h���X3", T_AppSetting)
        
End Function
' **************************************************
Public Function App_�A�v���L������()
On Error Resume Next
    App_�A�v���L������ = DLookup("�A�v���L������", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�閧�R�[�h() As String
On Error Resume Next
    App_�閧�R�[�h = DLookup("�閧�R�[�h", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�Í��L�[() As String
On Error Resume Next
    App_�Í��L�[ = DLookup("�Í��L�[", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�A�v��������()
On Error Resume Next
    App_�A�v�������� = DLookup("�A�v��������", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�����R�[�h() As String
On Error Resume Next
    App_�����R�[�h = DLookup("�����R�[�h", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�����[�X��()
On Error Resume Next
    App_�����[�X�� = DLookup("�����[�X��", T_AppSetting)
    
End Function
' **************************************************
Public Function App_�f�[�^�t�@�C����() As String
On Error Resume Next
    App_�f�[�^�t�@�C���� = DLookup("�f�[�^�t�@�C����", T_AppSetting)
    
End Function



' **************************************************
' App�f�[�^������
' **************************************************
Public Sub App�f�[�^������()
    Const methodName As String = "App�f�[�^������"
    ' �錾��
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim tablesToDelete As Collection
    Dim tblName As Variant
On Error GoTo ErrHandler
    
    ' ���C������
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
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(tdf)
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ***********************************************
' MAC�A�h���X�X�V
' ***********************************************
Public Sub MAC�A�h���X�X�V()
    Const methodName As String = "MAC�A�h���X�X�V"
    ' �錾��
    Dim macs As Variant
    Dim i As Long
    Dim sql As String
    Dim vals(1 To 3) As String
On Error GoTo ErrHandler
    
    ' ���C������
    macs = �[��MAC�A�h���X�擾�z��()

    ' --- �ő�3�������󕶎��ŏ����� ---
    For i = 1 To 3
        If IsArray(macs) And UBound(macs) >= i Then
            vals(i) = macs(i)
        Else
            vals(i) = ""
        End If
    Next

    ' --- UPDATE���쐬 ---
    sql = "UPDATE S_�A�v���ݒ� SET " & _
          "[MAC�A�h���X1]='" & Replace(vals(1), "'", "''") & "', " & _
          "[MAC�A�h���X2]='" & Replace(vals(2), "'", "''") & "', " & _
          "[MAC�A�h���X3]='" & Replace(vals(3), "'", "''") & "';"

    CurrentDb.Execute sql, dbFailOnError

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Sub


' ***********************************************
' �[��MAC�A�h���X�擾�z��i�����Ȃ��j
' ***********************************************
Public Function �[��MAC�A�h���X�擾�z��() As Variant
    Const methodName As String = "�[��MAC�A�h���X�擾�z��"
    ' �錾��
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
    
    ' ���C������
    count = 0
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapter WHERE MACAddress IS NOT NULL AND PhysicalAdapter=True")
    
    ' �ꎞ�i�[
    For Each objAdapter In colAdapters
        mac = Trim(objAdapter.MACAddress)
        name = LCase(Trim(objAdapter.name & " " & objAdapter.NetConnectionID))
        If mac <> "" Then
            ' --- ��ʂ𕪗ށi�������������D��j ---
            If InStr(name, "ethernet") > 0 Or InStr(name, "lan") > 0 Then
                category = 1   ' �L��LAN
            ElseIf InStr(name, "wireless") > 0 Or InStr(name, "wi-fi") > 0 Or InStr(name, "wifi") > 0 Then
                category = 2   ' ����LAN
            Else
                category = 3   ' ���̑��i���z�Ȃǁj
            End If
            
            ' --- �d���`�F�b�N ---
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
    
    ' --- �D�揇�Ƀ\�[�g�icategory���j ---
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
    
    ' --- ���ʂ�Ԃ� ---
    �[��MAC�A�h���X�擾�z�� = tmpMacs

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function



' ******************************************************
' �A�v�����p�`�F�b�N
' ******************************************************
Public Function Auth�A�v�����p�`�F�b�N() As String
    Const methodName As String = "Auth�A�v�����p�`�F�b�N"
    Dim errMsg As String
On Error GoTo ErrHandler

    If Not AuthMAC�A�h���X�F�؃`�F�b�N Then
        Auth�A�v�����p�`�F�b�N = "�����ꂽ�[���ȊO�ł̗��p�͔F�߂Ă��܂���B"
        Exit Function
    End If
    
    If Not Auth�L�������F�؃`�F�b�N Then
        Auth�A�v�����p�`�F�b�N = "�A�v�����p�̗L���������؂�܂����B"
        Exit Function
    End If
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function


' ******************************************************
' �L�������F�؃`�F�b�N�i�A�v���L�������A�A�v���������j
' ******************************************************
Function Auth�L�������F�؃`�F�b�N() As Boolean
    Const methodName As String = "Auth�L�������F�؃`�F�b�N"
On Error GoTo ErrHandler

    Auth�L�������F�؃`�F�b�N = True
    
    If Not IsNull(App_�A�v��������) Then Exit Function
    If Date <= App_�A�v���L������ Then Exit Function

    Auth�L�������F�؃`�F�b�N = False
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function


' ******************************************************
' MAC�A�h���X�F�؃`�F�b�N�i�A�v���L�������A�A�v���������j
' ******************************************************
Function AuthMAC�A�h���X�F�؃`�F�b�N() As Boolean
    Const methodName As String = "Auth�L�������F�؃`�F�b�N"
    Dim ary1 As Variant
    Dim result As Boolean
On Error GoTo ErrHandler
    
    ary1 = �[��MAC�A�h���X�擾�z��
    
    result = �z��v�f���݃`�F�b�N(ary1, Nz(App_MAC�A�h���X1))
    AuthMAC�A�h���X�F�؃`�F�b�N = result
    If result Then Exit Function
    
    result = �z��v�f���݃`�F�b�N(ary1, Nz(App_MAC�A�h���X2))
    AuthMAC�A�h���X�F�؃`�F�b�N = result
    If result Then Exit Function
    
    result = �z��v�f���݃`�F�b�N(ary1, Nz(App_MAC�A�h���X3))
    AuthMAC�A�h���X�F�؃`�F�b�N = result
    If result Then Exit Function
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Function


