Attribute VB_Name = "Main"
Sub エクセルシートの数値を集計するマクロ()

'-----変数宣言開始-----
    Dim targetFolderPath As String  '作業対象フォルダ
    Dim targetFileNames() As String '作業対象エクセルシート名（配列）
    Dim targetSheet As String '作業対象シート名
    
    Dim resultFileName As String  '結果出力エクセルファイル名
    Dim resultFile As Workbook    '結果出力エクセルブック
    Dim resultSheet As Worksheet  '結果出力エクセルシート
    
    Dim sCell As String '集計セルの最初の位置
    Dim eCell As String '集計セルの最後の位置
'-----変数宣言終了-----

'-----変数値設定開始-----
    targetSheet = "テスト"
    sCell = "H3"
    eCell = "BE5"
'-----変数値設定終了-----


'-----準備処理開始-----
    ' フォルダを指定して変数に格納（Owner kitazima）
    Call SelectBooks(targetFolderPath, targetFileNames)
    ' 結果出力用シートを設定（Owner suzuki）
    Call OpenResultSheet(resultFile, resultSheet)
'-----準備処理終了-----
    
'-----メイン処理開始-----
    ' 一つ一つファイルを開き、処理を実行を取得（Owner kitazima）
    Call ProcessBooks(targetFolderPath, targetFileNames, targetSheet, resultSheet, sCell, eCell)
'-----メイン処理終了-----
    
'-----最終処理開始-----
    ' 結果出力ファイルの保管（Owner suzuki）
    Call OutputResultFile(resultFile)
'-----最終処理終了-----

End Sub
