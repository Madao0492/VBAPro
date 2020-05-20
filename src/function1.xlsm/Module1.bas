Attribute VB_Name = "Module1"

'---------------------------------------------------------------
'�@セル位置・シート名を指定するマクロ
'（１）エクセルシート指定して開く
'（２）開いたエクセルシートのうち「セル位置」「シート名」を指定
'（３）指定位置のセルをアクティブにして値を任意の変数に代入する
'Owner kinoshita
'---------------------------------------------------------------


Sub kagebunshin()
Attribute kagebunshin.VB_ProcData.VB_Invoke_Func = " \n14"
'
' kagebunshin Macro

Dim sheetname As String
Dim rangeselect As String
Dim opensheet As String
Dim targetvalue As Long

sheetname = "Sheet1"

rangeselect = "A1"

opensheet = "test01"



Worksheets(opensheet).Select
    MsgBox "opensheet を開きました"
    
Worksheets(opensheet).Activate
Range("A3").Activate

targetvalue = Range("C3").Value

MsgBox "変数の値は" & targetvalue & "です"
'    Worksheets(sheetname).Range(rangeselect) = "test2"
'    Sheets.Add After:=ActiveSheet
'    ActiveCell.Formula = "a"
'
'      MsgBox "完了 "
      
End Sub
