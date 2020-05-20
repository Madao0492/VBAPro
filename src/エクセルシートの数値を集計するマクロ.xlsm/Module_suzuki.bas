Attribute VB_Name = "Module_suzuki"
'---------------------------------------------------------------
'結果出力ファイルの準備をするマクロ
'（１）エクセルブックを「新規作成」する
'（２）「新規作成」したエクセルブックを開き、シートをアクティブにする
'（３）名前を付けて保存する（上書き保存を確認する）
'Owner suzuki
'---------------------------------------------------------------

Function OpenResultSheet(resultFile As Workbook, resultSheet As Worksheet)
    '新規ファイル作成
    Set resultFile = Workbooks.Add
    '新規ファイルの1 番目のシートを変数に格納
    Set resultSheet = resultFile.Sheets(1)
 End Function
  
Function OutputResultFile(resultFile As Workbook)
    Dim resultFileName As String
 
   '保存するファイルの名前を入力（名前を付けて保存）
      resultFileName = Application.GetSaveAsFilename( _
          InitialFileName:="集計結果.xlsx", FileFilter:="Excelファイル, *.xlsx")

  'ファイル名が入力されたか判定
   If resultFileName <> "False" Then
  
     'ファイル名が入力された場合
      '名前を付けて保存
       On Error GoTo Error1
       ActiveWorkbook.SaveAs Filename:=resultFileName
       Exit Function
    End If

'エラー処理
Error1:
    'ActiveWorkbook.Close
    ActiveWorkbook.Close savechanges:=False
    MsgBox "保存されませんでした"
    Err.Clear

End Function


