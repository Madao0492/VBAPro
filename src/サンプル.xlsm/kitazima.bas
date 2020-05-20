Attribute VB_Name = "kitazima"
'---------------------------------------------------------------
'②指定フォルダ内のエクセルシートを順番に開いて閉じるマクロ
'（１）作業対象フォルダパスを指定
'（２）作業対象フォルダパス内のエクセルシートを取得
'（３）（２）のエクセルシートを順番に「開く⇒閉じる」
'Owner kitazima
'---------------------------------------------------------------

Sub OpenAndCloseBooks()

'-----変数宣言-----
    Dim foPath As String '対象のフォルダパス
    Dim fiName As String 'foPath内のエクセルファイル名
    
'-----フォルダパス・ファイルパス取得（3パターン）-----
    'フォルダパス選択1：ダイアログから選択
    With Application.FileDialog(msoFileDialogFolderPicker) 'フォルダ選択画面を表示
        If .Show = 0 Then '未選択の場合
            Exit Sub 'マクロを終了
        Else '選択した場合
            foPath = .SelectedItems(1) '選択したフォルダパスを取得
        End If
    End With
    
    'フォルダパス選択2：マクロを置いたフォルダとする
    foPath = ThisWorkbook.Path
    
    'フォルダパス選択3：パス固定（直接入力）
    foPath = "C:\Users\Kitajima\Desktop"
    
    fiName = Dir(foPath & "\*.xls*") '対象フォルダの最初のファイル名
    
'-----ファイルを開いて閉じる-----
    Do While fiName <> "" 'フォルダにエクセルファイルがある場合
        Workbooks.Open foPath & "\" & fiName '開く
        
        'ここに開いたあとの処理を記載
        
        Workbooks(fiName).Close SaveChanges:=False '上書きせずファイルを閉じる
        fiName = Dir '次のファイルの検索
    Loop
    
End Sub

