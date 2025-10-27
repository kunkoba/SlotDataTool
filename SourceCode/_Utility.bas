Attribute VB_Name = "_Utility"
Option Compare Database
Option Explicit


' **************************************************
' ���ׂẴt�H�[���𓝈�d�l�ɂ��悤1
' **************************************************
Public Sub utl�t�H�[���d�l����(���O�t�H�[�� As String, isPopUp As Boolean)
    Const methodName As String = "utl�t�H�[���d�l����"
    ' �錾��
    Dim accObj As AccessObject
    Dim frm As Form
    Dim ctl As Control
On Error GoTo ErrHandler
    
    ' ���C������
    For Each accObj In CurrentProject.AllForms
    
        If accObj.name <> ���O�t�H�[�� And Left(accObj.name, 2) <> "F_" Then
        
            ' �t�H�[�����f�U�C���r���[����\���ŊJ��
            DoCmd.OpenForm accObj.name, acDesign, , , , acHidden
            Set frm = Forms(accObj.name)
    
            ' �t�H�[���S�̂̃v���p�e�B�ύX
            With frm
                .PopUp = isPopUp                      ' �|�b�v�A�b�v�ɂ���
                .Modal = isPopUp                      ' ���[�_���ɂ͂��Ȃ��i��ƃE�B���h�E�ɌŒ�j
                .ShortcutMenu = Not isPopUp       ' ���E�N���b�N�������i�����ǉ��I�j
    '            .RecordSelectors = False           ' ���R�[�h�Z���N�^��\��
    '            .AllowDesignChanges = True         ' �f�U�C���ύX���i�O�̂��߁j
    '            .AllowDatasheetView = False        ' �f�[�^�V�[�g�\���͖���
    '            .BorderStyle = 1
    '            .MinMaxButtons = 0
    '            .ScrollBars = 0                    ' �X�N���[���o�[�Ȃ��i�K�v�Ȃ璲���j
    '            .AutoCenter = True
    '            .AllowLayoutView = False
    '            .NavigationButtons = False
    '            .Moveable = True
            End With
    
            ' �ۑ����ĕ���
            DoCmd.Close acForm, accObj.name, acSaveYes
        
        End If
        
    Next accObj

    Call ShowToast("���ׂẴt�H�[���̎d�l�𓝈ꂵ�܂����B", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    
End Sub



' **************************************************
' ���ׂĂ̊֘A�I�u�W�F�N�g���擾����i������ԋp�Łj
' ���� delimiter : ��؂蕶���i�� vbCrLf, vbTab, "," �Ȃǁj
' **************************************************
Public Function utl�ˑ��I�u�W�F�N�g���o������(���O�t�H�[�� As String, ByVal �������[�h As String, Optional ByVal delimiter As String = ",") As String
    Const methodName As String = "utl�ˑ��I�u�W�F�N�g���o������"
    ' �錾��
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim doc As AccessObject
    Dim frm As Form
    Dim rpt As Report
    Dim objName As String
    Dim ctl As Control
    Dim result As String
On Error GoTo ErrHandler
    
    ' ���C������
    Set db = CurrentDb
    Application.Echo False

    ' === �N�G�� =============================
    For Each qdf In db.QueryDefs
        If InStr(1, qdf.sql, �������[�h, vbTextCompare) > 0 Then
            If Left(qdf.name, 1) <> "~" Then
                result = result & "Query:" & qdf.name & delimiter
            End If
        End If
    Next

    ' === �t�H�[�� ===========================
    For Each doc In Application.CurrentProject.AllForms
        objName = doc.name
        
        If Left(objName, 1) <> "~" And Left(objName, 2) <> "F_" Then
                
            If objName <> ���O�t�H�[�� Then
                
                DoCmd.OpenForm objName, acDesign, , , , acHidden
                Set frm = Forms(objName)
        
                If InStr(1, frm.RecordSource, �������[�h, vbTextCompare) > 0 Then
                    result = result & "Form:" & objName & "�iRecordSource�j" & delimiter
                End If
        
                For Each ctl In frm.Controls
                    If ctl.ControlType = acComboBox Or ctl.ControlType = acListBox Then
                        If InStr(1, ctl.RowSource, �������[�h, vbTextCompare) > 0 Then
                            result = result & "Form:" & objName & "�iRowSource: " & ctl.name & "�j" & delimiter
                        End If
                    End If
                Next ctl
        
                DoCmd.Close acForm, objName, acSaveNo
            
            End If
            
        End If
    Next

    ' === ���|�[�g ===========================
    For Each doc In Application.CurrentProject.AllReports
        objName = doc.name
        DoCmd.OpenReport objName, acDesign, , , acHidden
        Set rpt = Reports(objName)

        If InStr(1, rpt.RecordSource, �������[�h, vbTextCompare) > 0 Then
            If Left(objName, 1) <> "~" Then
                result = result & "Report:" & objName & "�iRecordSource�j" & delimiter
            End If
        End If

        DoCmd.Close acReport, objName, acSaveNo
    Next

    ' �߂�l�ɐݒ�
    utl�ˑ��I�u�W�F�N�g���o������ = result
    
    Call ShowToast("�֘A�I�u�W�F�N�g�̒��o�͊������܂����B", E_Color_Blue)

ErrHandler:
    Application.Echo True
    Call ErrorSave(methodName, False, objName) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
'    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function



' **************************************************
' ���ׂẴI�u�W�F�N�g���𕶎���ŕԂ�
' ���� mode
'   1 = �e�[�u���ꗗ
'   2 = �N�G���ꗗ
'   3 = �t�H�[���ꗗ
' ���� delim ��؂蕶��
'   ��: "," �� vbCrLf �� vbTab
' **************************************************
Public Function utl�I�u�W�F�N�g���ꗗ�擾(ByVal mode As Integer, Optional ByVal delim As String = ",") As String
    Const methodName As String = "utl�I�u�W�F�N�g���ꗗ�擾"
    Dim obj As AccessObject
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim qdf As DAO.QueryDef
    Dim result As String
On Error GoTo ErrHandler

    result = ""

    Set db = CurrentDb
    Select Case mode
        Case 1 ' �e�[�u��
            Set db = CurrentDb
            For Each tdf In db.TableDefs
                ' �V�X�e���e�[�u���iMSys�`�j�͏��O
                If Left(tdf.name, 4) <> "MSys" Then
                    If result <> "" Then result = result & delim
                    result = result & tdf.name
                End If
            Next
        
        Case 2 ' �N�G��
            For Each qdf In db.QueryDefs
                ' �ꎞ�N�G���i~...�j�͏��O
                If Left(qdf.name, 1) <> "~" Then
                    If result <> "" Then result = result & delim
                    result = result & qdf.name
                End If
            Next
        
        Case 3 ' �t�H�[��
            For Each obj In CurrentProject.AllForms
                If result <> "" Then result = result & delim
                result = result & obj.name
            Next
        
        Case Else
            result = "mode�� 1=�e�[�u�� / 2=�N�G�� / 3=�t�H�[�� ���w�肵�Ă�������"
    End Select

    utl�I�u�W�F�N�g���ꗗ�擾 = result
    
'    MsgBox "�I�u�W�F�N�g���ꗗ�擾���������܂����B", vbInformation
    
ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(db)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Function




' ***********************************************
' �e�[�u����`����
' ***********************************************
Sub utl�e�[�u����`����(tableName As String)
    Const methodName As String = "utl�e�[�u����`����"
    ' �錾��
    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim fld As DAO.field
    Dim idx As DAO.index
    Dim pkFields As Collection
On Error GoTo ErrHandler
    
    ' ���C������
    Set pkFields = New Collection
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tableName)
    
    ' ��L�[�̃t�B�[���h�����W
    For Each idx In tdf.Indexes
        If idx.Primary Then
            Dim f As DAO.field
            For Each f In idx.fields
                pkFields.Add f.name
            Next f
        End If
    Next idx
    
    ' CREATE TABLE ���쐬
    Dim sql As String
    sql = "CREATE TABLE [" & tableName & "] (" & vbCrLf
    
    Dim fieldSQL As String
    For Each fld In tdf.fields
        fieldSQL = "  [" & fld.name & "] " & Convert�t�B�[���h�^�C�v(fld.Type, fld.size)
        If fld.Required Then fieldSQL = fieldSQL & " NOT NULL"
        fieldSQL = fieldSQL & ","
        sql = sql & fieldSQL & vbCrLf
    Next fld
    
    ' ��L�[�ݒ�
    If pkFields.count > 0 Then
        Dim pkSQL As String
        pkSQL = "  PRIMARY KEY ("
        Dim i As Long
        For i = 1 To pkFields.count
            pkSQL = pkSQL & "[" & pkFields(i) & "]"
            If i < pkFields.count Then pkSQL = pkSQL & ", "
        Next i
        pkSQL = pkSQL & ")"
        sql = sql & pkSQL & vbCrLf
    Else
        ' �Ō�̃J���}���폜
        sql = Left(sql, Len(sql) - 3) & vbCrLf
    End If
    
    sql = sql & ");"
    
    Call LogDebug("utl�e�[�u����`����", sql)

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(db)
    Call ProcNothing(tdf)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
        
End Sub


' ***********************************************
' XXX
' ***********************************************
Function Convert�t�B�[���h�^�C�v(type_ As Long, size As Long) As String
On Error Resume Next
    Select Case type_
        Case dbText
            Convert�t�B�[���h�^�C�v = "TEXT(" & size & ")"
        Case dbMemo
            Convert�t�B�[���h�^�C�v = "MEMO"
        Case dbByte
            Convert�t�B�[���h�^�C�v = "BYTE"
        Case dbInteger
            Convert�t�B�[���h�^�C�v = "INTEGER"
        Case dbLong
            Convert�t�B�[���h�^�C�v = "LONG"
        Case dbCurrency
            Convert�t�B�[���h�^�C�v = "CURRENCY"
        Case dbSingle
            Convert�t�B�[���h�^�C�v = "SINGLE"
        Case dbDouble
            Convert�t�B�[���h�^�C�v = "DOUBLE"
        Case dbDate
            Convert�t�B�[���h�^�C�v = "DATETIME"
        Case dbBoolean
            Convert�t�B�[���h�^�C�v = "YESNO"
        Case Else
            Convert�t�B�[���h�^�C�v = "TEXT"
    End Select
    
End Function



' ************************************************
' �\�[�X�R�[�h�o��
' ************************************************
Sub utl�\�[�X�R�[�h�o��()
    Const methodName As String = "utl�\�[�X�R�[�h�o��"
    Const Path_Source As String = "SourceCode\"
    ' �錾��
    Dim vbcmp As Object
    Dim strFileName As String
    Dim strExt As String
    Dim batPath As String
    Dim strCmd As String
    Dim outPath As String
On Error GoTo ErrHandler
    
    ' �o�̓t�H���_�ݒ�
    outPath = CurrentProject.path & "\" & Path_Source
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(outPath)
    
    ' ���C������
    For Each vbcmp In VBE.ActiveVBProject.VBComponents
        With vbcmp
            '�t�@�C�����܂ł�ݒ�
            strFileName = outPath & .name
            '�g���q��ݒ�
            Select Case .Type
                Case 1        '�W�����W���[��
                    strExt = ".bas"
                Case 2        '�N���X���W���[��
                    strExt = ".cls"
                Case 100      '�t�H�[��/���|�[�g�̃��W���[��
                    strExt = ".cls"
                Case Else
                    strExt = ".txt"
            End Select
            ' ���W���[�����G�N�X�|�[�g
            .Export strFileName & strExt
        End With
    Next vbcmp

    Sleep 1000
    
    batPath = CurrentProject.path & "\ExportedObjects\__Convert_UTF8.bat"
    strCmd = "cmd /c """ & batPath & """"
    shell strCmd, vbHide

    Call ShowToast("�G�N�X�|�[�g���������܂����B", E_Color_Blue)

ErrHandler:
    Call ErrorSave(methodName, True, strFileName) '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(vbcmp)
'    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j
    
End Sub




