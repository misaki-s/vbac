Attribute VB_Name = "Module1"
Dim cnt As Long

Sub Sample3233(Path As String)
    Dim buf As String, f As Object
    buf = Dir(Path & "\*.*")
    Do While buf <> ""
        cnt = cnt + 1
        Cells(cnt, 1) = buf
        buf = Dir()
    Loop
    With CreateObject("Scripting.FileSystemObject")
        For Each f In .GetFolder(Path).SubFolders
            Call Sample3(f.Path)
        Next f
    End With
End Sub

