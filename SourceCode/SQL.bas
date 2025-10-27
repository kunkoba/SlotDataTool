Attribute VB_Name = "SQL"
Option Compare Database
Option Explicit


' ***********************************************
' ���I��SQL���Z�b�g����
' ***********************************************
Public Sub Proc�N�G����SQL����������(queryName As String, newSQL As String)
    Const methodName As String = "Proc�N�G����SQL����������"
    Dim qdf As DAO.QueryDef
On Error GoTo ErrHandler

    ' ���C������
    Set qdf = CurrentDb.QueryDefs(queryName)
    qdf.sql = newSQL
    
    Call ProcNothing(qdf)
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub



' ***********************************************
' �N�G�� D_�W�v���f�[�^ ��SQL�𓮓I�ɕύX����
' ***********************************************
Public Sub Update�W�v��Query(Optional �@��ID As Variant = Null, _
                             Optional ��ԍ�ID As Variant = Null, _
                             Optional �J�n�� As Variant = Null, _
                             Optional �I���� As Variant = Null, _
                             Optional ���̑� As Variant = Null)
    Const methodName As String = "Update�W�v��Query"
    Const queryName As String = "D_�W�v���f�[�^_�t�B���^"
    Dim sql As String
    Dim ���� As String
On Error GoTo ErrHandler

    ' ���C������
    ���� = ""

    If Not IsNull(�@��ID) And �@��ID <> 0 And �@��ID <> "" Then
        ���� = JoinText(����, " AND �@��ID = " & �@��ID, True)
    End If

    If Not IsNull(��ԍ�ID) And ��ԍ�ID <> 0 And ��ԍ�ID <> "" Then
        ���� = JoinText(����, " AND ��ԍ� = " & ��ԍ�ID, True)
    End If

    If Not IsNull(�J�n��) And �J�n�� <> "" Then
        ���� = JoinText(����, " AND ���t >= #" & Format(�J�n��, "yyyy/mm/dd") & "#", True)
    End If

    If Not IsNull(�I����) And �I���� <> "" Then
        ���� = JoinText(����, " AND ���t <= #" & Format(�I����, "yyyy/mm/dd") & "#", True)
    End If
    
    If Not IsNull(���̑�) And ���̑� <> "" Then
        ���� = JoinText(����, " AND " & ���̑�, True)
    End If

    If ���� <> "" Then
        ���� = " WHERE " & Mid(����, 6)  ' �擪�� AND ���폜���� WHERE ��
    End If

    sql = "SELECT * " & _
          "FROM D_�W�v���f�[�^ " & ���� & ";"
    
    ' �����N�G�����擾����SQL����������
    Call Proc�N�G����SQL����������(queryName, sql)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub

' ***********************************************
' �N�G�� D_�W�v���f�[�^ ��SQL�𓮓I�ɕύX����
' ***********************************************
Public Sub Update�W�v��Query_�ŐV(Optional �@��ID As Variant = Null, _
                                Optional ��ԍ�ID As Variant = Null, _
                                Optional �J�n�� As Variant = Null, _
                                Optional �I���� As Variant = Null, _
                                Optional ���̑� As Variant = Null)
    Const methodName As String = "Update�W�v��Query_�ŐV"
    ' �錾��
    Const queryName As String = "D_�W�v���f�[�^_�ŐV_�t�B���^"
    Dim sql As String
    Dim ���� As String
On Error GoTo ErrHandler
    
    ' ���C������
    ���� = ""

    If Not IsNull(�@��ID) And �@��ID <> 0 And �@��ID <> "" Then
        ���� = JoinText(����, " AND �@��ID = " & �@��ID, True)
    End If

    If Not IsNull(��ԍ�ID) And ��ԍ�ID <> 0 And ��ԍ�ID <> "" Then
        ���� = JoinText(����, " AND ��ԍ� = " & ��ԍ�ID, True)
    End If

    If Not IsNull(�J�n��) And �J�n�� <> "" Then
        ���� = JoinText(����, " AND ���t >= #" & Format(�J�n��, "yyyy/mm/dd") & "#", True)
    End If

    If Not IsNull(�I����) And �I���� <> "" Then
        ���� = JoinText(����, " AND ���t <= #" & Format(�I����, "yyyy/mm/dd") & "#", True)
    End If
    
    If Not IsNull(���̑�) And ���̑� <> "" Then
        ���� = JoinText(����, " AND " & ���̑�, True)
    End If

    If ���� <> "" Then
        ���� = " WHERE " & Mid(����, 6)  ' �擪�� AND ���폜���� WHERE ��
    End If

    sql = "SELECT * " & _
          "FROM D_�W�v���f�[�^_�ŐV " & ���� & ";"
    
    ' �����N�G�����擾����SQL����������
    Call Proc�N�G����SQL����������(queryName, sql)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub



' ***********************************************
' �O���[�v���t�B�[���h�ƏW�v�Ώۃt�B�[���h����SQL������𐶐����ĕԂ�
' ***********************************************
Public Function GenerateGraphSQL_A(���ڗ� As String, �W�v�� As String) As String
    Const methodName As String = "GenerateGraphSQL_A"
    ' �錾��
    Dim sql As String
On Error GoTo ErrHandler

    ' ���C������
    sql = "SELECT " & ���ڗ� & ", Avg(" & �W�v�� & ") AS �W�v�l " & _
          "FROM D_�W�v���f�[�^_�t�B���^ " & _
          "GROUP BY " & ���ڗ�

'    Call LogDebug("GenerateGraphSQL_A", sql)
    
    GenerateGraphSQL_A = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function

' ***********************************************
' �O���[�v���t�B�[���h�ƏW�v�Ώۗ񂩂�SQL������𐶐����ĕԂ�
' ���{�̏����𓮓I�ɕt�^�\
' ***********************************************
Public Function GenerateGraphSQL_B(���ڗ�1 As String, ���ڗ�2 As String, �W�v�� As String, _
                                   Optional �t�B���^�l As Variant) As String
    Const methodName As String = "GenerateGraphSQL_B"
    ' �錾��
    Dim sql As String
    Dim whereStr As String
On Error GoTo ErrHandler

    ' ���C������
    If Not IsMissing(�t�B���^�l) Then
        If Not IsNull(�t�B���^�l) And Trim(�t�B���^�l & "") <> "" Then
            whereStr = " WHERE CStr([" & ���ڗ�2 & "]) = """ & �t�B���^�l & """"
        End If
    End If
    
    ' SQL�g�ݗ���
    sql = "SELECT " & ���ڗ�1 & ", Avg(" & �W�v�� & ") AS �W�v�l " & _
      "FROM D_�W�v���f�[�^_�t�B���^" & whereStr & _
      " GROUP BY " & ���ڗ�1 & ";"
    
    GenerateGraphSQL_B = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



' ***********************************************
' �N�G�� L_��ԍ����X�g ��SQL�𓮓I�ɕύX����
' ***********************************************
Public Sub Update��ԍ�Query(Optional �@��ID As Variant)
    Const methodName As String = "Update��ԍ�Query"
    ' �錾��
    Dim sql As String
    Const queryName As String = "L_��ԍ����X�g_�t�B���^"
On Error GoTo ErrHandler

    ' ���C������
    sql = "SELECT * FROM L_��ԍ����X�g"
    
    ' �@��ID���w�肳��Ă����WHERE���t�^
    If Not IsNull(�@��ID) And �@��ID <> "" Then
        sql = sql & " WHERE �@��ID = " & �@��ID
    Else
        sql = sql & " WHERE 1 = 0"
    End If
    
    ' �����N�G�����擾����SQL����������
    Call Proc�N�G����SQL����������(queryName, sql)
    
    Debug.Print sql
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ***********************************************
' �N�G�� L_�i�荞�݃��X�g ��SQL�𓮓I�ɕύX����
' ***********************************************
Public Sub Update�i�荞��Query(ByVal fieldName As String)
    Const methodName As String = "Update�i�荞��Query"
    ' �錾��
    Dim sql As String
    Const queryName As String = "L_�i�荞�݃��X�g"
On Error GoTo ErrHandler

    ' �t�B�[���h������Ȃ珈�����Ȃ�
    If Nz(fieldName, "") = "" Then Exit Sub

    ' ���C������
    sql = "SELECT DISTINCT " & fieldName & " AS �i�荞�� " & _
          "FROM T_SLOT�W�v�敪 " & _
          "ORDER BY " & fieldName & ";"

    ' �����N�G�����擾����SQL����������
    Call Proc�N�G����SQL����������(queryName, sql)
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


' ***********************************************
' ��������SQL������𐶐����ĕԂ�
' ***********************************************
Public Function Generate�ςݏグ����SQL(�@��ID As String, ���я� As String) As String
    Const methodName As String = "Generate�ςݏグ����SQL"
    ' �錾��
    Dim sql As String
On Error GoTo ErrHandler

    ' ���C������
    sql = "SELECT �@�햼, ���t, ����, �ςݏグ���� " & _
          "FROM Check_�ςݏグ��_�@��� " & _
          "WHERE �@��ID = " & �@��ID & " " & _
          "ORDER BY " & ���я� & ";"
    
    Generate�ςݏグ����SQL = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



