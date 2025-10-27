Attribute VB_Name = "Slot"
Option Compare Database
Option Explicit


' ***************************************************************
Public Function �N�����W���O����(ByVal src As Variant) As Variant
    Const methodName As String = "�N�����W���O����"
    ' �錾��
    Dim s As String
On Error GoTo ErrHandler
    
    �N�����W���O���� = ""
    
    ' ���C������
    If IsNull(src) Then
        Exit Function
    End If

    s = CStr(src)
    
    ' �S�p�{ �� �폜
    s = Replace(s, "�{", "")
    ' ���p+ �� �폜
    s = Replace(s, "+", "")
    ' �J���} , �� �폜
    s = Replace(s, ",", "")
    
    �N�����W���O���� = s
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �ݒ蔻��_BB
' ***************************************************************
Public Function �ėp�ݒ蔻��(�Q�[���� As Long, ������ As Long, _
                          �ݒ�l1 As Double, �ݒ�l2 As Double, �ݒ�l3 As Double, _
                          �ݒ�l4 As Double, �ݒ�l5 As Double, �ݒ�l6 As Double) As Variant
    Const methodName As String = "�ėp�ݒ蔻��"
    ' �錾��
    Dim �������� As Double
    Dim �ݒ�l(1 To 6) As Double
    Dim i As Integer
    Dim val1 As Double, val2 As Double
    Dim �ݒ萄��l As Double
On Error GoTo ErrHandler

    �ėp�ݒ蔻�� = 1
    
    ' ���C������
    If ������ <= 0 Then
        Exit Function
    End If

    �������� = �Q�[���� / ������

    ' �ݒ�l�z��
    �ݒ�l(1) = �ݒ�l1
    �ݒ�l(2) = �ݒ�l2
    �ݒ�l(3) = �ݒ�l3
    �ݒ�l(4) = �ݒ�l4
    �ݒ�l(5) = �ݒ�l5
    �ݒ�l(6) = �ݒ�l6

    ' ���`���
    For i = 1 To 5
        val1 = �ݒ�l(i)
        val2 = �ݒ�l(i + 1)
        If val1 > 0 And val2 > 0 Then
            If (�������� >= val1 And �������� <= val2) Or (�������� >= val2 And �������� <= val1) Then
                �ݒ萄��l = i + (�������� - val1) / (val2 - val1)
                �ėp�ݒ蔻�� = Round(�ݒ萄��l, 1)
                Exit Function
            End If
        End If
    Next i

    ' �ł��߂��l�ɂ���i��Ԕ͈͊O�j
    Dim �ŏ��� As Double: �ŏ��� = 999999
    Dim �ŏ��ݒ� As Integer
    For i = 1 To 6
        If �ݒ�l(i) > 0 Then
            If Abs(�������� - �ݒ�l(i)) < �ŏ��� Then
                �ŏ��� = Abs(�������� - �ݒ�l(i))
                �ŏ��ݒ� = i
            End If
        End If
    Next i

    �ėp�ݒ蔻�� = Round(�ŏ��ݒ�, 1)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �ݒ葍���]��
' ***************************************************************
Public Function �ݒ葍���]��(�ݒ�_BB As Variant, �ݒ�_RB As Variant, �ݒ�_���Z As Variant, _
                            wBB As Variant, wRB As Variant, w���Z As Variant, _
                            �Q�[���� As Long, ��]��_�������l As Long, �ݒ蒲���l As Integer) As Variant
    Const methodName As String = "�ݒ葍���]��"
    ' �錾��
    Dim ret As Variant
    Dim �]���l As Double
On Error GoTo ErrHandler

    �ݒ葍���]�� = Null
    
    ' ���C������
    If IsNull(�ݒ�_BB) Then
        Exit Function
    End If

    �]���l = �ݒ�_BB * wBB + �ݒ�_RB * wRB + �ݒ�_���Z * w���Z
    �ݒ葍���]�� = Round(�]���l / (wBB + wRB + w���Z), 1)
    
    
    ' �ݒ蒲��
    If �Q�[���� < ��]��_�������l Then
        �ݒ葍���]�� = �ݒ葍���]�� - �ݒ蒲���l
        If �ݒ葍���]�� < 1 Then �ݒ葍���]�� = 1
    End If
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �@�B��������
' ***************************************************************
Public Function �@�B��������(�Q�[���� As Long, �z��ݒ� As Double, _
                    ��1 As Double, ��2 As Double, ��3 As Double, _
                    ��4 As Double, ��5 As Double, ��6 As Double) As Long
    Const methodName As String = "�@�B��������"
    ' �錾��
    Dim �ݒ艺 As Integer, �ݒ�� As Integer
    Dim ���� As Double
    Dim ����� As Double
    Dim arr�� As Variant
On Error GoTo ErrHandler

    �@�B�������� = 0
    
    ' ���C������
    If �Q�[���� <= 0 Or �z��ݒ� < 1 Or �z��ݒ� > 6 Then
        �@�B�������� = 0
        Exit Function
    End If

    ' �@�B����z��Ɂi�Y��1�`6�j
    arr�� = Array(0, ��1, ��2, ��3, ��4, ��5, ��6)

    �ݒ艺 = Int(�z��ݒ�)
    If �ݒ艺 + 1 > 6 Then
        �ݒ�� = 6
    Else
        �ݒ�� = �ݒ艺 + 1
    End If

    If �ݒ艺 = �ݒ�� Then
        ����� = arr��(�ݒ艺)
    Else
        ����� = arr��(�ݒ艺) + (arr��(�ݒ��) - arr��(�ݒ艺)) * (�z��ݒ� - �ݒ艺)
    End If

    ���� = �Q�[���� * 3 * (����� / 100 - 1)
    �@�B�������� = CLng(����)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �z�荷����
' ***************************************************************
Public Function �z�荷����(�Q�[���� As Long, BB�� As Long, RB�� As Long, BB���� As Double, RB���� As Double, �������W�� As Double) As Long
    Const methodName As String = "�z�荷����"
    ' �錾��
    Dim ������ As Double
On Error GoTo ErrHandler

    �z�荷���� = 0

    ������ = BB�� * BB���� + RB�� * RB���� + �Q�[���� * �������W�� - �Q�[���� * 3

    �z�荷���� = Round(������)

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' ���茋�ʐM���x
' ***************************************************************
Public Function ���茋�ʐM���x(�Q�[���� As Long) As String
    Const methodName As String = "���茋�ʐM���x"
    ' �錾��
On Error GoTo ErrHandler

    ���茋�ʐM���x = ""
    
    ' ���C������
    If �Q�[���� < 2000 Then
        ���茋�ʐM���x = "1��"
    
    ElseIf �Q�[���� < 5000 Then
        ���茋�ʐM���x = "2��"
    
    Else
        ���茋�ʐM���x = "3��"
    
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �o�ʌX��
' ***************************************************************
Public Function �o�ʌX��(�L������ As Long, �@�B������ As Long, �������l As Long) As String
    Const methodName As String = "�o�ʌX��"
    ' �錾��
    Dim diff As Long
On Error GoTo ErrHandler
    
    �o�ʌX�� = "�z�����"
        
    ' ���C������
    If �@�B������ = 0 Then
        Exit Function
    End If

    diff = Abs(�L������ - �@�B������)

    ' �]�����W�b�N
    If diff <= �������l Then
        �o�ʌX�� = "�z�����"
    ElseIf �L������ > �@�B������ Then
        �o�ʌX�� = "��U�ꁪ"
    Else
        �o�ʌX�� = "���U�ꁫ"
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function



' ******************************************************
' �ėp�F�w��̕ϊ��Ώۂɉ����Ēl��ϊ�����i�G���[�� -1�A������� Trim�j
' ******************************************************
Public Function �l�ϊ�(sDate As Variant, convType As String) As Variant
    Const methodName As String = "�l�ϊ�"
    ' �錾��
    Dim result As Variant
On Error GoTo ErrHandler

    �l�ϊ� = ""
    
    ' ���C������
    If sDate = "" Or sDate = Null Then
        Exit Function
    End If

    Select Case LCase(convType)
        Case "����"
            result = Right(CStr(sDate), 1)
            
        Case "���t����"
            result = Right(CStr(sDate), 1)
        Case "��ԍ�����"
            result = Right(CStr(sDate), 1)
        Case "�N��"
            result = Format(sDate, "yyyymm")
        Case "��"
            result = Month(sDate)
        Case "�j��"
            sDate = CDate(sDate)
            result = Weekday(sDate) & Format(sDate, "aaa")
        Case "���{"
            result = ���{�ϊ�(CDate(sDate)) ' ���ʓr ���{ �֐����K�v
        Case "���{6"
            result = ���{6�����ϊ�(CDate(sDate)) ' ���ʓr ���{6���� �֐����K�v
        Case "��"
            result = day(sDate)
    End Select

    ' ������Ȃ� Trim�A���l�Ȃ炻�̂܂ܕԂ�
    If VarType(result) = vbString Then
        �l�ϊ� = Trim(result)
    Else
        �l�ϊ� = result
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' ���{�i��E���E���j���Z�o
' ***************************************************************
Public Function ���{�ϊ�(���t As Date) As String
    Const methodName As String = "���{�ϊ�"
    ' �錾��
    Dim �� As Integer
On Error GoTo ErrHandler

    ���{�ϊ� = ""
    
    ' ���C������
    �� = day(���t)

    If �� <= 10 Then
        ���{�ϊ� = "1��{"
    ElseIf �� <= 20 Then
        ���{�ϊ� = "2���{"
    Else
        ���{�ϊ� = "3���{"
    End If

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function

' ***************************************************************
' ���{�i5������6�����j���Z�o
' ***************************************************************
Public Function ���{6�����ϊ�(���t As Date) As String
    Const methodName As String = "���{6�����ϊ�"
    ' �錾��
    Dim �� As Integer
On Error GoTo ErrHandler

    ���{6�����ϊ� = ""
    
    ' ���C������
    �� = day(���t)
    
    Dim �敪 As Integer
    �敪 = Int((�� - 1) / 5) + 1

    If �敪 > 6 Then �敪 = 6 ' ���S�����i�ő�ł���6�{�j

    ���{6�����ϊ� = "��" & �敪 & "�{"

ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
' �{�[�i�X���I��
' �@������Ȃ獇�Z�l
' ***************************************************************
Function �{�[�i�X���I��(ByVal bonus1 As Long, ByVal �Q�[���� As Long, Optional ByVal bonus2 As Long = 0) As String
    Const methodName As String = "�{�[�i�X���I��"
    ' �錾��
    Dim totalBonus As Long
    Dim rate As Double
    Dim denominator As Long
On Error GoTo ErrHandler
    
    �{�[�i�X���I�� = "1/999"
        
    ' ���C������
    If �Q�[���� <= 0 Then
        Exit Function
    End If

    totalBonus = bonus1 + bonus2

    If totalBonus <= 0 Then
        �{�[�i�X���I�� = "1/999"
    Else
        rate = totalBonus / �Q�[����
        denominator = CLng(1 / rate)

        If denominator >= 1000 Then
            �{�[�i�X���I�� = "1/999"
        Else
            �{�[�i�X���I�� = "1/" & Format(denominator, "000")  ' 3���Œ�i��: 1/007, 1/250�j
        End If
    End If
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function


' ***************************************************************
Public Function ���I���Z�o(ByVal �Q�[���� As Long, ByVal �{�[�i�X�� As Long) As String
    Const methodName As String = "���I���Z�o"
    Dim rate As Double
On Error GoTo ErrHandler
    
    ���I���Z�o = "--"
    
    ' BB����0�̏ꍇ
    If �{�[�i�X�� = 0 Then
        Exit Function
    End If
    
    ' ���I���v�Z
    rate = �Q�[���� / �{�[�i�X��
    
    ' �l�̌ܓ����Đ�����
    rate = Round(rate, 0)
    
    ' ���q�i�v�Z���ʁj��1000�ȏ�Ȃ�999�ɐ���
    If rate >= 1000 Then
        rate = 999
    End If
    
    ' �����񉻂��ĕԋp
    ���I���Z�o = "1/" & CStr(rate)
    
ErrHandler:
    If Err.Number <> 0 Then Debug.Print Now, methodName, Err.Number, Err.Description
End Function



