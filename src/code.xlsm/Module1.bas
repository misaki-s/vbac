Attribute VB_Name = "Module1"
Dim cnt As Long

'イミディエイトウィンドウクリア
Sub mCls()
    Dim i
    Dim t As String
    t = ""
    For i = 1 To 200
         t = t & vbCrLf
    Next
    Debug.Print t
End Sub

'フォルダの再帰検索
Sub getFolderR(Path As String)
    Dim buf As String, f As Object
    buf = Dir(Path & "\*.*")
    Do While buf <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = buf
        buf = Dir()
    Loop
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            If .FolderExists(f.Path) Then
                Debug.Print (f.Path & " isDir")
            Else
                Debug.Print (f.Path & " isFile")
            End If
            
            Call getFolderR(f.Path)
        Next f
    End With
End Sub


Sub Test()
    Call mCls
    cnt = 0
    Call getFolderR("C:\works\vbac")
End Sub
