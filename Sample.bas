Attribute VB_Name = "Sample"
Option Explicit


Public Sub Sample_Is�j��()
    '�w�肵�����t���j�����m�F���܂��B


    Dim ���t As Date
    ���t = ���t�����()

    On Error GoTo Catch

    Debug.Print Is�j��(���t)

    Exit Sub
    
Catch:
    Call MsgBox("���s���G���[: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Public Sub Sample_�j�������擾()
    '�w�肳�ꂽ���t���j���ł���Ώj���̖��O���Ԃ���܂��B
    '�j���ł͂Ȃ��ꍇ�A�󕶎��񂪕Ԃ���܂��B
    '���̂��߁A�󕶎���ł���Ώj���ł͂Ȃ��Ɣ��f�ł��܂��B


    Dim ���t As Date
    ���t = ���t�����()

    On Error GoTo Catch

    '�w�肵�����t�̏j�������擾���܂��B
    Dim �j�� As String
    �j�� = �j�������擾(���t)
    
    If �j�� = "" Then
        Debug.Print "�j���ł͂���܂���B"
    Else
        Debug.Print �j��
    End If

    Exit Sub
    
Catch:
    Call MsgBox("���s���G���[: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Private Function ���t�����() As Date

    Dim ���t As Date
    ���t = InputBox("���t����͂��Ă��������B")
    ���t = DateValue(���t)
    ���t����� = ���t

End Function


Public Sub Sample_�j��csv���쐬()
    '�w�肵���K�w�ɏj����csv�t�@�C�����쐬���܂��B

    Dim �N As Long
    �N = �N�����()

    Dim �ۑ��� As String
    �ۑ��� = �ۑ�������()

    On Error GoTo Catch

    Call �j��csv���쐬(�N, �ۑ���)

    Exit Sub

Catch:
    Call MsgBox("���s���G���[: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)

End Sub


Public Sub Sample_�j���ꗗ���擾()
    '�j���iDictionary�^�j���擾���� For Each �Ŏ��o���܂��B

    Dim �N As Long
    �N = �N�����()
    
    On Error GoTo Catch

    Dim �j���ꗗ As Dictionary
    Set �j���ꗗ = �j���ꗗ���擾(�N)
    
    Dim ���t
    For Each ���t In �j���ꗗ
        Debug.Print ���t, �j���ꗗ.item(���t)
    Next

    Exit Sub
    
Catch:
    Call MsgBox("���s���G���[: " & Err & vbNewLine _
        & Err.Description, vbCritical, Err.Source)
    
End Sub


Private Function �N�����() As Long

    Dim �N As Long
    �N = InputBox("�N����͂��Ă��������B")
    �N����� = �N

End Function


Private Function �ۑ�������() As String

    Dim �ۑ��� As String
    �ۑ��� = InputBox("�ۑ�����΃p�X�Ŏw�肵�Ă��������B")
    �ۑ������� = �ۑ���

End Function
