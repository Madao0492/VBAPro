Attribute VB_Name = "Module_kitazima"
'---------------------------------------------------------------
'指定フォルダ内のエクセルシートを順番に開いて閉じるマクロ
'（１）作業対象フォルダパスを指定
'（２）作業対象フォルダパス内のエクセルシートを取得
'（３）（２）のエクセルシートを順番に「開く⇒閉じる」
'Owner kitazima
'---------------------------------------------------------------

Function SelectBooks(foPath As String, fiName() As String)

'-----変数宣言-----
    Dim fiNum As Long           '対象フォルダ内に保管されているエクセルファイルの数
    Dim tempfiName As String    '一時的なエクセルファイル名保管箇所
    
'-----フォルダパス・ファイルパス取得（3パターン）-----
    'フォルダパス選択1：ダイアログから選択
    With Application.FileDialog(msoFileDialogFolderPicker) 'フォルダ選択画面を表示
        If .Show = 0 Then '未選択の場合
            Exit Function 'マクロを終了
        Else '選択した場合
            foPath = .SelectedItems(1) '選択したフォルダパスを取得
        End If
    End With
    
    'フォルダパス選択2：マクロを置いたフォルダとする
    'foPath = ThisWorkbook.Path
    
    'フォルダパス選択3：パス固定（直接入力）
    'foPath = "C:\Users\Kitajima\Desktop"
    
    
    fiNum = 0
    tempfiName = Dir(foPath & "\*.xls*") '対象フォルダの最初のファイル名
    
'-----フォルダ内のエクセルファイル名をすべて取得-----
    Do While tempfiName <> "" 'フォルダにエクセルファイルがある場合
        ReDim Preserve fiName(fiNum)
        fiName(fiNum) = tempfiName
        tempfiName = Dir '次のファイルの検索
        fiNum = fiNum + 1
    Loop
    
End Function

Function ProcessBooks(foPath As String, fiName() As String, resultSheet As Worksheet)
    Dim i As Integer
    Dim sum As Double
    Dim fcell As Double
    
    sum = 0
    
    'ファイル名のみの比較
    '（フォルダパスは比較しないが、ファイル名が同一のファイルを開くとエラーとなるため回避要）
    'SelectBooksで取得したファイル名がすでに開かれていないかチェックする
    'Filter : fiName内にwb.Nameが含まれていないと-1を返す
    'End : プログラム全体を終了
    For Each wb In Workbooks
        If UBound(Filter(fiName, wb.Name)) <> -1 Then
            MsgBox "処理対象のファイルがすでに開かれているため処理を中止しました", vbCritical
            End
        End If
    Next wb
    
    For i = 0 To UBound(fiName)
        Workbooks.Open foPath & "\" & fiName(i) '開く
        ' セルを指定して、値を返す（Owner kinoshita）
        fcell = Kagebunshin("テスト", "H3")
        ' 取得した値を足して出力する（Owner ooba）
        Call Sumcells(sum, fcell, resultSheet)
        Workbooks(fiName(i)).Close SaveChanges:=False   '上書きせずファイルを閉じる
    Next i
    
End Function
    
