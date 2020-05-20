Attribute VB_Name = "Module_kinoshita"
'---------------------------------------------------------------
'セル位置・シート名を指定する機能
'（１）エクセルシート指定して開く
'（２）開いたエクセルシートのうち「セル位置」「シート名」を指定
'（３）指定位置のセルをアクティブにして値を任意の変数に代入する
'Owner kinoshita
'---------------------------------------------------------------


Function Kagebunshin(openSheet As String, rangeSelect As String) As Long

' kagebunshin Macro

    'Dim sheetname As String
    'Dim rangeselect As String
    'Dim opensheet As String
    Dim targetvalue As Long

    Worksheets(openSheet).Select
    'MsgBox "openSheet を開きました"
    
    Worksheets(openSheet).Activate
    Range(rangeSelect).Activate

    targetvalue = Range(rangeSelect).Value
    
    Kagebunshin = targetvalue

    'MsgBox "変数の値は" & targetvalue & "です"
    '    Worksheets(sheetname).Range(rangeselect) = "test2"
    '    Sheets.Add After:=ActiveSheet
    '    ActiveCell.Formula = "a"
    '
    '      MsgBox "完了 "
      
End Function
