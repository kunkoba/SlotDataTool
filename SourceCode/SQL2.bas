Attribute VB_Name = "SQL2"
Option Compare Database
Option Explicit


' ***********************************************
' ���I�N�G�������i�W�v�N�G���j
' ***********************************************
Public Function Proc���I�W�v�N�G��SQL����(�s1 As Variant, �s2 As Variant, �s3 As Variant, where�� As String) As String
    Const methodName As String = "Proc���I�W�v�N�G��SQL����"
    ' �錾��
    Dim sql As String
    Dim groupFields As String
    Dim selectFields As String
    Dim f1 As String, f2 As String, f3 As String
On Error GoTo ErrHandler
    
    ' ���C������
    ' --- Null/�󔒃`�F�b�N ---
    If IsNull(�s1) Or Trim(�s1 & "") = "" Then
        f1 = """"""
    Else
        f1 = �s1
    End If
    
    If IsNull(�s2) Or Trim(�s2 & "") = "" Then
        f2 = """"""
    Else
        f2 = �s2
    End If
    
    If IsNull(�s3) Or Trim(�s3 & "") = "" Then
        f3 = """"""
    Else
        f3 = �s3
    End If
    
    ' --- SELECT�� ---
    selectFields = f1 & " AS ����1, " & _
                   f2 & " AS ����2, " & _
                   f3 & " AS ����3, " & _
                   "Count(�@��ID) AS �f�[�^����, " & _
                   "Sum(���x) AS ���x�̍��v, " & _
                   "Sum(�Q�[����) AS �Q�[�����̍��v, " & _
                   "Sum(BB��) AS BB���̍��v, " & _
                   "Sum(RB��) AS RB���̍��v, " & _
                   "Sum(������) AS �������̍��v, " & _
                   "Avg(������) AS �������̕���, " & _
                   "Avg(�ݒ蔻��) AS �ݒ蔻�ʂ̕���, " & _
                   "Avg(�ݒ�4�ȏ�) AS �ݒ�4������, " & _
                   "Avg(�ݒ�5�ȏ�) AS �ݒ�5������, " & _
                   "Avg(�ݒ�6) AS �ݒ�6������"
    
    ' --- GROUP BY�� ---
    groupFields = f1 & ", " & f2 & ", " & f3

    ' --- SQL�g�ݗ��� ---
    sql = " SELECT " & selectFields & _
          " FROM D_�W�v���f�[�^"

    If Trim(where�� & "") <> "" Then
        sql = sql & " WHERE " & where�� & vbCrLf
    End If

    sql = sql & " GROUP BY " & groupFields & _
                " ORDER BY " & groupFields

    ' ���O�o��
    Call LogDebug("Proc���I�W�v�N�G��SQL����", sql)
    Proc���I�W�v�N�G��SQL���� = sql

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' ���I�N�G�������i�N���X�W�v�N�G���j
' ***********************************************
Public Function Proc���I�N���X�W�vSQL����( _
                    �s1 As String, �s2 As String, _
                    �� As String, �l As String, where�� As String) As String
    Const methodName As String = "Proc���I�N���X�W�vSQL����"
    ' �錾��
    Dim sql As String
    Dim groupClause As String
    Dim selectClause As String
    Dim whereClause As String
On Error GoTo ErrHandler
        
    ' SELECT��GROUP BY
    If �s1 <> "" Then
        selectClause = �s1
        groupClause = " GROUP BY " & �s1
    End If

    If �s2 <> "" Then
        If selectClause <> "" Then
            selectClause = selectClause & ", " & �s2
            groupClause = groupClause & ", " & �s2
        Else
            selectClause = �s2
            groupClause = " GROUP BY " & �s2
        End If
    End If

    ' WHERE��
    If where�� <> "" Then
        whereClause = " WHERE " & where��
    Else
        whereClause = ""
    End If

    ' SQL�\�z
sql = " TRANSFORM Round(Avg(" & �l & "), 2) AS ��1 " & _
      " SELECT " & selectClause & "," & _
      " Count(" & �� & ") AS �f�[�^��," & _
      " Round(Avg(" & �l & "), 2) AS �S�� " & _
      " FROM D_�W�v���f�[�^" & _
      whereClause & _
      groupClause & _
      " ORDER BY " & selectClause & _
      " PIVOT " & �� & ";"
    
    Proc���I�N���X�W�vSQL���� = sql

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' ���I�����N�G���i�`���[�g�p�W�v�N�G�� + �t�B���^�Ή��j
' ***********************************************
Public Function Proc���I�`���[�g�pSQL����( _
                field1 As String, field2 As String, _
                sumField As String, aggFunc As String, _
                filterValue1 As String, filterValue2 As String) As String
    Const methodName As String = "Proc���I�`���[�g�pSQL����"
    ' �錾��
    Dim sql As String
    Dim validAgg As String
    Dim selectField1 As String
    Dim selectField2 As String
    Dim groupField1 As String
    Dim groupField2 As String
    Dim whereClause As String
On Error GoTo ErrHandler
    
    ' �W�v�֐��̃o���f�[�V����
    Select Case LCase(aggFunc)
        Case "sum": validAgg = "Sum"
        Case "avg": validAgg = "Avg"
        Case Else:  validAgg = "Avg"
    End Select

    ' field1, field2 ���󔒂Ȃ� '' �ɒu������
    selectField1 = IIf(Trim(field1) = "", "''", field1)
    groupField1 = selectField1
    selectField2 = IIf(Trim(field2) = "", "''", field2)
    groupField2 = selectField2

    ' HAVING�吶��
    If Trim(field1) <> "" And Trim(filterValue1) <> "" Then
        whereClause = field1 & " = " & Convert�����l������(filterValue1)
    End If

    If Trim(field2) <> "" And Trim(filterValue2) <> "" Then
        If whereClause <> "" Then whereClause = whereClause & " AND "
        whereClause = whereClause & field2 & " = " & Convert�����l������(filterValue2)
    End If

    ' SQL �g�ݗ���
    sql = "SELECT TOP 50 " & _
          selectField1 & " AS ���ڂP, " & _
          selectField2 & " AS ���ڂQ, " & _
          validAgg & "(" & sumField & ") AS �l " & _
          "FROM D_�W�v���f�[�^_�t�B���^ "
    
    ' WHERE��ǉ�
    If whereClause <> "" Then
        sql = sql & "WHERE " & whereClause
    End If
    
    sql = sql & _
          " GROUP BY " & _
          groupField1 & ", " & _
          groupField2 & _
          " ORDER BY " & groupField1 & " DESC, " & groupField2 & " DESC"
        
    Proc���I�`���[�g�pSQL���� = sql

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



' ******************************************************
' �Œ�e�[�u���FT_SLOT�W�v�敪 �̃��X�g�擾�pSQL����
' ******************************************************
Public Function Proc���X�g�pSQL����(fieldName As String) As String
    Const methodName As String = "Proc���X�g�pSQL����"
    ' �錾��
    Dim sql As String
On Error GoTo ErrHandler

    ' ���C������
    sql = " SELECT DISTINCT " & fieldName & _
          " FROM D_�W�v���f�[�^" & _
          " ORDER BY " & fieldName & ";"
    
    Proc���X�g�pSQL���� = sql
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function


' ***********************************************
' D_�W�v���f�[�^_�t�B���^�i���I�ύX�j
' ***********************************************
Public Sub Proc�O���t�p�N�G�����I�ύX( _
                Optional �@��ID As Variant, _
                Optional ��ԍ� As Variant, _
                Optional �J�n�� As Variant, _
                Optional �I���� As Variant)
    Const methodName As String = "Proc�O���t�p�N�G�����I�ύX"
    ' �錾��
    Const �N�G���� As String = "D_�W�v���f�[�^_�t�B���^"
    Const �e�[�u���� As String = "D_�W�v���f�[�^"
    Dim ���� As String
    Dim �VSQL As String
On Error GoTo ErrHandler
    
    ' �����쐬
    If Not IsNull(�@��ID) And Trim(�@��ID & "") <> "" Then
        ���� = ���� & " AND [�@��ID] = " & �@��ID
    End If

    If Not IsNull(��ԍ�) And Trim(��ԍ� & "") <> "" Then
        ���� = ���� & " AND [��ԍ�] = " & ��ԍ�
    End If

    If Not IsNull(�J�n��) And Trim(�J�n�� & "") <> "" Then
        ���� = ���� & " AND [���t] >= #" & Format(�J�n��, "yyyy/mm/dd") & "#"
    End If

    If Not IsNull(�I����) And Trim(�I���� & "") <> "" Then
        ���� = ���� & " AND [���t] <= #" & Format(�I����, "yyyy/mm/dd") & "#"
    End If

    ' WHERE��쐬
    If ���� <> "" Then
        ���� = " WHERE " & Mid(����, 6)  ' �擪�� AND ���폜���� WHERE ��
    End If

    ' SQL�쐬
    �VSQL = "SELECT * FROM " & �e�[�u���� & ����

    ' SQL �u��
    Call Proc�N�G����SQL����������(�N�G����, �VSQL)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub


