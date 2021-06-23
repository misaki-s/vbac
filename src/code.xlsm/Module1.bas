Attribute VB_Name = "Module1"
Dim cnt As Long
Dim depth As Long
Dim max_depth As Long

'イミディエイトウィンドウクリア
Sub Cls()
    Dim i
    Dim t As String
    t = ""
    For i = 1 To 200
         t = t & vbCrLf
    Next
    Debug.Print t
End Sub

'フォルダの再帰検索
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


'ファイルの再帰検索
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
    Dim count As Long '文字カウント数
    count = Len(text) - Len(Replace(text, c, ""))
    FindCharCount = count
End Function
'Sub Test2()
'    cnt = FindCharCount("C:\works\vbac\.git\COMMITMESSAGE", "\")
'    MsgBox "【\】は" & cnt & "個ありまーす"
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


    cnt = 1 'データ開始行
'    Call getDirR("C:\works\vbac")
    Call getFileR("C:\works\vbac")
    

End Sub


' ユーザー定義関数(m_)

' 探索文字、出現回数
Function m_FindCharCount(address, c)
    m_FindCharCount = FindCharCount(address, c)
End Function

' 拡張子判定(MSOffice)
Function m_isMsOffice(address)
    Dim re
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "\.xls$|\.xlsx$|\.xlsm$|\.ppt$|\.pptx$|\.doc$|\.docx$"
    m_isMsOffice = re.Test(address)
End Function

'ドキュメント作成者取得
Function m_isMsOfficeAuthor(address)
'    Dim re
'    Set re = CreateObject("VBScript.RegExp")
'    re.Pattern = "\.xls$|\.xlsx$|\.xlsm$|\.ppt$|\.pptx$|\.doc$|\.docx$"
'    m_isMsOffice = re.Test(address)
    Dim obj As Object, s As String
    s = ""
    Set obj = GetObject(address)
    If Err.Number <> 0 Then
        'Officeのドキュメントではないということ
        Err.Clear
    Else
        s = obj.BuiltinDocumentProperties(3)  'Author https://excel-ubara.com/excelvba4/EXCEL256.html
        obj.Close
    End If
    m_isMsOfficeAuthor = s
End Function

'ファイル種別を取得
Function m_fileType(address)
    Dim fso As Object, fs As Object, r As String
    r = ""
    Dim attType As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    r = fso.GetFile(address).Type
    m_fileType = r
    
'    Attribute なぜか全部テキストもzipファイルも32を返すので使えない。 https://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q13146844163
'    If attType And 0 Then
'      r = "標準ファイル"
'    ElseIf attType And 1 Then
'      r = "読み取り専用ファイル"
'    ElseIf attType And 2 Then
'      r = "隠しファイル"
'    ElseIf attType And 4 Then
'      r = "システムファイル"
'    ElseIf attType And 8 Then
'      r = "ディスクドライブボリュームラベル"
'    ElseIf attType And 16 Then
'      r = "フォルダまたはディレクトリ"
'    ElseIf attType And 32 Then
'      r = "アーカイブファイル"
'    ElseIf attType And 64 Then
'      r = "リンクまたはショートカット"
'    ElseIf attType And 128 Then
'      r = "圧縮ファイル"
'    End If
'    Set fso = Nothing
'    m_fileAttribute = r
End Function

'ファイル名を取得
Function m_fileName(address)
    Dim fso As Object, fs As Object, r As String
    r = ""
    Dim attType As Integer
    Set fso = CreateObject("Scripting.FileSystemObject")
    r = fso.GetFileName(address)
    m_fileName = r
End Function
