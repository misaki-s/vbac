Attribute VB_Name = "Module1"
Dim cnt As Long
Dim depth As Long
Dim max_depth As Long

'�C�~�f�B�G�C�g�E�B���h�E�N���A
Sub Cls()
    Dim i
    Dim t As String
    t = ""
    For i = 1 To 200
         t = t & vbCrLf
    Next
    Debug.Print t
End Sub

'�t�H���_�̍ċA����
Sub getDirR(path As String)
    Dim buf As String, f As Object
'    buf = Dir(path & "\*.*")
'    Do While buf <> ""
'        cnt = cnt + 1
'        Cells(cnt, 1) = path & "\" & buf
'        buf = Dir()
'    Loop
    cnt = cnt + 1
    Cells(cnt, 1) = path
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(path).SubFolders
            If .FolderExists(f.path) Then
                Debug.Print (f.path & " isDir")
                If FindCharCount(f.path, "\") <= max_depth Then
                    Call getDirR(f.path)
                End If
'            Else
'                Debug.Print (f.path & " isFile")
            End If

        Next f
    End With
End Sub


'�t�@�C���̍ċA����
Sub getFileR(path As String)
    Dim f As Object
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(path).Files
            cnt = cnt + 1
            Debug.Print f.path
            Cells(cnt, 1) = f.path
        Next f
        For Each f In .GetFolder(path).SubFolders
            If .FolderExists(f.path) Then
                If FindCharCount(f.path, "\") <= max_depth Then
                    Call getFileR(f.path)
                End If
            End If
        Next f
    End With
End Sub





Function FindCharCount(text, c)
    Dim count As Long '�����J�E���g��
    count = Len(text) - Len(Replace(text, c, ""))
    FindCharCount = count
End Function
'Sub Test2()
'    cnt = FindCharCount("C:\works\vbac\.git\COMMITMESSAGE", "\")
'    MsgBox "�y\�z��" & cnt & "����܁[��"
'End Sub


Sub Test()
    Call Cls
    Dim ignore As Variant
    ignore = Range("A1").Value
    Columns("A").Clear
    Range("A1").Value = ignore
    
    depth = FindCharCount("C:\works\vbac", "\")
    Debug.Print depth
    max_depth = 2
    
    depth = depth
    max_depth = max_depth + depth
    Debug.Print max_depth


    cnt = 1 '�f�[�^�J�n�s
'    Call getDirR("C:\works\vbac")
    Call getFileR("C:\works\vbac")
    

End Sub


' ���[�U�[��`�֐�(m_)

' �T�������A�o����
Function m_FindCharCount(address, c)
    m_FindCharCount = FindCharCount(address, c)
End Function

' �g���q����(MSOffice)
Function m_isMsOffice(address)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\.xls$|\.xlsx$|\.xlsm$|\.ppt$|\.pptx$|\.doc$|\.docx$"
    m_isMsOffice = re.Test(address)
End Function

'�h�L�������g�쐬�Ҏ擾
Function m_isMsOfficeAuthor(address)
'    Dim re
'    Set re = CreateObject("VBScript.RegExp")
'    re.Pattern = "\.xls$|\.xlsx$|\.xlsm$|\.ppt$|\.pptx$|\.doc$|\.docx$"
'    m_isMsOffice = re.Test(address)
    Dim obj As Object, s As String
    s = ""
    Set obj = GetObject(address)
    If Err.Number <> 0 Then
        'Office�̃h�L�������g�ł͂Ȃ��Ƃ�������
        Err.Clear
    Else
        s = obj.BuiltinDocumentProperties(3)  'Author https://excel-ubara.com/excelvba4/EXCEL256.html
        obj.Close
    End If
    m_isMsOfficeAuthor = s
End Function

'�t�@�C����ʂ��擾
Function m_fileType(address)
    Dim fso As Object, fs As Object, r As String
    r = ""
    Dim attType As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    r = fso.GetFile(address).Type
    m_fileType = r
    
'    Attribute �Ȃ����S���e�L�X�g��zip�t�@�C����32��Ԃ��̂Ŏg���Ȃ��B https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13146844163
'    If attType And 0 Then
'      r = "�W���t�@�C��"
'    ElseIf attType And 1 Then
'      r = "�ǂݎ���p�t�@�C��"
'    ElseIf attType And 2 Then
'      r = "�B���t�@�C��"
'    ElseIf attType And 4 Then
'      r = "�V�X�e���t�@�C��"
'    ElseIf attType And 8 Then
'      r = "�f�B�X�N�h���C�u�{�����[�����x��"
'    ElseIf attType And 16 Then
'      r = "�t�H���_�܂��̓f�B���N�g��"
'    ElseIf attType And 32 Then
'      r = "�A�[�J�C�u�t�@�C��"
'    ElseIf attType And 64 Then
'      r = "�����N�܂��̓V���[�g�J�b�g"
'    ElseIf attType And 128 Then
'      r = "���k�t�@�C��"
'    End If
'    Set fso = Nothing
'    m_fileAttribute = r
End Function

'�t�@�C�������擾
Function m_fileName(address)
    Dim fso As Object, fs As Object, r As String
    r = ""
    Dim attType As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    r = fso.GetFileName(address)
    m_fileName = r
End Function
