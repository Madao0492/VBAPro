Attribute VB_Name = "Module_ooba"
'---------------------------------------------------------------
'合計を出力するマクロ
'（１）エクセルシート指定して開く
'（２）ある２つのセルに入力された値を２つの変数に読み込む
'（３）合計値を計算する
'（４）合計値を別のセルに出力する
'Owner ooba
'---------------------------------------------------------------


 Function Sumcells(fCell As Double, resultSheet As Worksheet, cPosition As String)

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
    resultSheet.Range(cPosition).Value = resultSheet.Range(cPosition).Value + fCell  '合計値を出力

    ' MsgBox Cells(1, 3) '計算結果出力

 End Function

 Function WriteLog(resultFile As Workbook, rNum As Integer, message As String)
    
    Dim ws As Worksheet
    
    If rNum = 1 Then '初回ログ出力（1行目）の場合
        Set ws = resultFile.Worksheets.Add
        ws.Name = "Log"
    End If
 
    resultFile.Worksheets("Log").Range("A" & rNum).Value = message
    rNum = rNum + 1

 End Function
