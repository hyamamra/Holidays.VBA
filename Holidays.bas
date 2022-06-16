Attribute VB_Name = "Holidays"

'Holidays JP API ���� json �𕶎���Ƃ��Ď擾���A
'Dictionary�I�u�W�F�N�g�ɉ��H���Ă��珈�����܂��B
'Holidays JP API �ɑ��݂��Ȃ��N���w�肳�ꂽ�ꍇ�A
'���[�U�[��`�̎��s���G���[: 513 �𔭐������܂��B
'�g�p����ۂ͗�O�����̓������������Ă��������B

'�Q�Ɛݒ�
    'Microsoft XML, v6.0
    'Microsoft Scripting Runtime

'Holidays JP API �����T�C�g
    'https://holidays-jp.github.io


Option Explicit


Public Function Is�j��(ByVal ���t As Date) As Boolean
    '�n���ꂽ���t���j�����ǂ�����^�U�l�ŕԂ��܂��B


    Dim �N As Long
    �N = Year(���t)

    Dim �j���ꗗ As Dictionary
    Set �j���ꗗ = �j���ꗗ���擾(�N)

    Is�j�� = �j���ꗗ.exists(���t)

End Function


Public Function �j�������擾(ByVal ���t As Date) As String
    '�n���ꂽ���t���j���Ȃ�j���̖��O��Ԃ��܂��B
    '�j���łȂ���΋󕶎����Ԃ��܂��B


    Dim �N As Long
    �N = Year(���t)

    Dim �j���ꗗ As Dictionary
    Set �j���ꗗ = �j���ꗗ���擾(�N)

    �j�������擾 = �j���ꗗ.item(���t)

End Function


Public Function �j��csv���쐬(Optional ByVal �N As Long, Optional ByVal �ۑ��� As String)
    '�w�肳�ꂽ�N�̏j���ꗗ��csv�����ĕۑ����܂��B
    '�N���w�肳��Ȃ���΍��N���w�肵�܂��B
    '�ۑ��悪�w�肳��Ȃ���Γ��K�w�ɕۑ����܂��B
    '���łɓ����̃t�@�C��������Ύ��s���G���[: 58 ���������܂��B
    '�ۑ���t�H���_�����݂��Ȃ���Ύ��s���G���[: 76 ���������܂��B


    If �N = 0 Then
        �N = Year(Now)
    End If

    If �ۑ��� = "" Then
        �ۑ��� = ThisWorkbook.Path & "\" & �N & "�N�̏j��.csv"
    Else
        �ۑ��� = �����ꕶ�����폜(�ۑ���, "/", "\")
        �ۑ��� = �ۑ��� & "\" & �N & "�N�̏j��.csv"
    End If

    Dim �j���ꗗ As Dictionary
    Set �j���ꗗ = �j���ꗗ���擾(�N)

    Dim FSO As New FileSystemObject

    On Error GoTo Finally

    Dim csv As TextStream
    Set csv = FSO.CreateTextFile(�ۑ���, False)

    Dim ���t
    For Each ���t In �j���ꗗ
        Call csv.WriteLine(���t & "," & �j���ꗗ.item(���t))
    Next

Finally:
    If Not csv Is Nothing Then
        Call csv.Close
    End If
    Set FSO = Nothing

    If Err Then
        Call Err.Raise(Err, "�j��csv���쐬", Err.Description)
    End If

End Function


Public Function �j���ꗗ���擾(Optional ByVal �N As Long) As Dictionary
    '�w�肳�ꂽ�N�̏j�����擾���܂��B
    '�N���w�肳��Ȃ���΍��N���w�肵�܂��B


    Dim �j��json As String
    �j��json = �j��json���擾(�N)

    Set �j���ꗗ���擾 = Json��Dictionay�ɕϊ�(�j��json)

End Function


Private Function �j��json���擾(ByVal �N As Long) As String
    '�w�肵���N�̏j����json�t�@�C���𕶎���Ƃ��Ď擾���܂��B
    '�N���w�肳��Ă��Ȃ���΍��N���w�肵�܂��B
    '�w�肳�ꂽ�N��json�t�@�C�����Ȃ���΃G���[�ɂȂ�܂��B


    Dim URL As String
    URL = Holidays_JP_API��URL���擾(�N)
    
    �j��json���擾 = Web�T�C�g�̑S��������擾(URL)
    
    '�擾����������json�`���łȂ���΃G���[�𔭐������܂��B
    If �j��json���擾 Like "*html*" Then
        Call Err.Raise(513, "�j��json���擾", _
            "�w�肳�ꂽ�N��WebAPI�����݂��܂���B�i " & �N & " �N )")
    End If
    
End Function


Private Function Holidays_JP_API��URL���擾(ByVal �N As Long) As String
    '�w�肳�ꂽ�N��Holidays JP API��URL��Ԃ��܂��B


    '�N���w�肳��Ă��Ȃ���΍��N���w�肵�܂��B
    If �N = 0 Then
        �N = Year(Now)
    End If

    Holidays_JP_API��URL���擾 = _
        "https://holidays-jp.github.io/api/v1/" & �N & "/date.json"

End Function


Private Function Web�T�C�g�̑S��������擾(ByVal URL As String) As String
    '�w�肳�ꂽURL�̃y�[�W����S��������擾���܂��B


    With New MSXML2.XMLHTTP60
        Call .Open("get", URL)
        Call .send
        Web�T�C�g�̑S��������擾 = .responseText()
    End With

End Function


Private Function Json��Dictionay�ɕϊ�(ByVal json As String) As Object
    'json��Dictionary�I�u�W�F�N�g�ɕϊ����܂��B
    
    '����
        'json : string

    '�߂�l
        'Object : Dictionary


    json = Json����]���ȕ����������(json)

    Dim ���ڈꗗ() As String
    ���ڈꗗ() = Split(json, ",")

    Dim jsonDictionary As New Dictionary

    Dim ����
    For Each ���� In ���ڈꗗ()
        Dim key As Date
        Dim item As String

        ' "���ږ�:�l" ���R�����ŕ����đ�����܂��B
        key = Split(����, ":")(0)
        item = Split(����, ":")(1)

        Call jsonDictionary.Add(key, item)
    Next

    Set Json��Dictionay�ɕϊ� = jsonDictionary

End Function


Private Function Json����]���ȕ����������(ByVal json As String) As String
    '�l�ƃR�����ƃJ���}�ȊO���폜���܂��B�i�����̃J���}�������j


    Dim �������镶���ꗗ() As Variant
    �������镶���ꗗ() = Array("{", "}", """", " ", vbNewLine)

    Dim �������镶��
    For Each �������镶�� In �������镶���ꗗ()
        json = Replace(json, �������镶��, "")
    Next

    json = �����ꕶ�����폜(json, ",")

    Json����]���ȕ���������� = json

End Function


Private Function �����ꕶ�����폜(ByVal ������ As String, ParamArray �폜���镶��()) As String
    '�폜���镶���ɍ��v���Ă���Ζ����ꕶ�����폜���܂��B


    Dim ������ As Long
    ������ = Len(������)
    
    Dim ����
    For Each ���� In �폜���镶��()
        If (Right(������, 1) = ����) Then
            �����ꕶ�����폜 = Left(������, ������ - 1)
            Exit Function
        End If
    Next

    �����ꕶ�����폜 = ������

End Function
