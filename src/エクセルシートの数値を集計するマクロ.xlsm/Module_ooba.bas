Attribute VB_Name = "Module_ooba"
'---------------------------------------------------------------
'合計を出力するマクロ
'（１）エクセルシート指定して開く
'（２）ある２つのセルに入力された値を２つの変数に読み込む
'（３）合計値を計算する
'（４）合計値を別のセルに出力する
'Owner ooba
'---------------------------------------------------------------


 Function Sumcells(fcell As Double, scell As Double, resultSheet As Worksheet)

'-----変数宣言-----
    'Dim Path As String '対象エクセルシートのファイルパス
    
    'Dim fcell As Double '１つ目の値
    'Dim scell As Double '２つ目の値

    ' Path = "C:\Users\xxxx\Desktop\hokan2\simnor2.xlsm" '仮のファイルバス
    
    '（１）エクセルシート指定して開く
    ' Workbooks.Open Path
    
    '（２）ある２つのセルに入力された値を２つの変数に読み込む
    ' fcell = Cells(1, 1) 'A1の値
    ' scell = Cells(1, 2) 'B1の値
    
    '（３）合計値を計算する
    '（４）合計値を別のセルに出力する
    resultSheet.Cells(1, 1).Value = fcell + scell  '合計値を出力
    fcell = fcell + scell  '合計値を保管

    ' MsgBox Cells(1, 3) '計算結果出力

 End Function

