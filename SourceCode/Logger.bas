Attribute VB_Name = "Logger"
Option Compare Database
Option Explicit


Const delim As String = " >> "
Dim outputMsg As String

'********************************************************************************
' ���O�͂����܂ŃG���[�ӏ�����肷�邽�߂����̂���
'********************************************************************************
Public Sub LogOpen(objName As String)
On Error Resume Next
    
    Call ErrorClear     ' �ŏ��ɃN���A
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, "=======================================================================================", True)
    outputMsg = JoinText(outputMsg, Now & delim & "[OPEN ]" & delim & objName)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
End Sub

'********************************************************************************
Public Sub LogStart(methodName As String)
On Error Resume Next
    
    Call ErrorClear     ' �ŏ��ɃN���A
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, "-------------------------------------------------", True)
    outputMsg = JoinText(outputMsg, Now & delim & "[START]" & delim & methodName)
    Debug.Print outputMsg
    Call Output(outputMsg)
        
End Sub

'********************************************************************************
Public Sub LogEnd(methodName As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[END  ]" & delim & methodName)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
    Call ErrorMessage   '�o�͂��ăN���A
    
End Sub

'********************************************************************************
Public Sub LogClose(objName As String)
On Error Resume Next
    
    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[CLOSE]" & delim & objName, True)
    outputMsg = JoinText(outputMsg, "=======================================================================================")
    Debug.Print outputMsg
    Call Output(outputMsg)
    
    Call ErrorMessage   '�o�͂��ăN���A
    
End Sub

'********************************************************************************
Public Sub LogError(methodName As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[ ERR ]" & delim & methodName & " >>>> " & ErrObj.Number & vbTab & ErrObj.Description)
    Debug.Print outputMsg
    Call Output(outputMsg)
    
End Sub

'********************************************************************************
Public Sub LogDebug(methodName As String, ByVal log1 As String)
On Error Resume Next

    outputMsg = ""
    outputMsg = JoinText(outputMsg, Now & delim & "[DEBUG]" & delim & methodName & delim & log1)
    Debug.Print outputMsg
    Call Output(outputMsg)

End Sub


'********************************************************************************
' ���O�t�@�C���ɒǋL����i���t�ʃt�@�C���j
'********************************************************************************
Private Sub Output(logText As String)
    Const methodName As String = "Output"
'    Const logFolder As String = "..\Log\"  ' App ���猩�����΃p�X
    Dim fso As Object
    Dim ts As Object
    Dim outPath As String
    Dim filaName As String
On Error Resume Next

    filaName = "log_" & Format(Date, "yyyymmdd") & ".log"
    outPath = CurrentProject.path & "\" & PATH_LOG

    Set fso = CreateObject("Scripting.FileSystemObject")
    Call Proc�t�H���_�����݂��Ȃ��ꍇ�͍쐬(outPath)

    Set ts = fso.OpenTextFile(outPath & "\" & filaName, 8, True)  ' 8 = ForAppending, True = create if not exists
    ts.WriteLine logText

ErrHandler:
    Call ErrorSave(methodName)  '�K���擪�i���������G���[���m���ɃL���b�`���邽�߁j
    Call ProcNothing(ts)
    Call ProcNothing(fso)
    If ErrObj.Number <> 0 Then Err.Raise ERR_TMP, "", "�^���G���["  ' �K���Ō�(�㑱�����������Ȃ����߁j

End Sub



