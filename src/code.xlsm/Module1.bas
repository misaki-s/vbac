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


Sub sample()
    Dim i As Long
    Dim strFile As String
    Dim strPath As String
    Dim obj As Object
  
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = "C:\"
        .AllowMultiSelect = False
        .Title = "フォルダの選択"
        If .Show = False Then
            Exit Sub
        End If
        strPath = .SelectedItems(1) & "\"
    End With
  
    Application.EnableEvents = False '起動時のOpenイベント等を停止
    On Error Resume Next 'GetObjectで取得できないファイルの対策
    strFile = Dir(strPath)
    i = 2
    Do While strFile <> ""
        Cells(i, 1) = strFile 'ファイル名
        Cells(i, 2) = FileDateTime(strPath & strFile) '更新日時
        Cells(i, 3) = FileLen(strPath & strFile) 'サイズ
        Set obj = GetObject(strPath & strFile)
        If Err.Number <> 0 Then
            'Officeのドキュメントではないということ
            Err.Clear
        Else
            Cells(i, 4).Value = obj.BuiltinDocumentProperties(3) 'Author
            Cells(i, 5).Value = obj.BuiltinDocumentProperties(7) 'Last Author
            obj.Close
        End If
        strFile = Dir()
        i = i + 1
    Loop
    Set obj = Nothing
    Application.EnableEvents = True
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
    Call getDirR("C:\works\vbac")
'    Call getFileR("C:\works\vbac")
    

End Sub


' ユーザー定義関数(m_)
Function m_FindCharCount(address, c)
    m_FindCharCount = FindCharCount(address, CStr(c))
End Function

