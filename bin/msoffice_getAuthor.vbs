'   *-------------------------------------------------------*
'   * EXCELファイルのプロパティを取得                       *
'   *                                                       *
'   * ID : sample.vbs                                       *
'   * 使用方法 : X:\>cscript sample.vbs xxxxx.xls           *
'   *-------------------------------------------------------*
Option Explicit
On Error Resume Next

Dim objExcel        ' EXCEL    オブジェクト
Dim objBook         ' Book     オブジェクト
Dim objPpty         ' Property オブジェクト
Dim strFileName     ' EXCELファイル名

'   - コマンドラインパラメータの取得 -
strFileName = WScript.Arguments(0)

'   - Fileの存在確認 -
If IsEmpty(strFileName) OR strFileName = "" Then
    Wscript.Echo "Excelファイルが存在しませんでした。" & Err.Description & "(" & Err.Number & ")"
    Wscript.Quit
Else
    Wscript.Echo "Excelファイル=" & strFileName
End If

'   - EXCELファイルの読み込みとプロパティ属性取得 -
Set objExcel = CreateObject("Excel.Application")             ' EXCELの起動
If Err.Number = 0 Then
    objExcel.Application.DisplayAlerts = False                ' 保存確認ダイアログを非表示
    Set objBook = objExcel.Workbooks.Open(strFileName)        ' EXCELファイルのオープン
    If Err.Number = 0 Then
        For Each objPpty In objBook.BuiltInDocumentProperties  ' プロパティを全て表示
            WScript.Echo objPpty.Name & ": " & objPpty.Value
        Next
        objBook.Close
        objExcel.Quit
    Else
        Wscript.Echo "Excelファイルを開けませんでした。" & Err.Description & "(" & Err.Number & ")"
    End If
Else
    Wscript.Echo "Excelを起動できませんでした。" & Err.Description & "(" & Err.Number & ")"
End If

Set objBook = Nothing
Set objExcel = Nothing