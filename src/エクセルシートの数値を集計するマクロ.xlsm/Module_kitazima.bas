Attribute VB_Name = "Module_kitazima"
'---------------------------------------------------------------
'�w��t�H���_���̃G�N�Z���V�[�g�����ԂɊJ���ĕ���}�N��
'�i�P�j��ƑΏۃt�H���_�p�X���w��
'�i�Q�j��ƑΏۃt�H���_�p�X���̃G�N�Z���V�[�g���擾
'�i�R�j�i�Q�j�̃G�N�Z���V�[�g�����ԂɁu�J���˕���v
'Owner kitazima
'---------------------------------------------------------------

Function SelectBooks(foPath As String, fiName() As String)

'-----�ϐ��錾-----
    Dim fiNum As Long           '�Ώۃt�H���_���ɕۊǂ���Ă���G�N�Z���t�@�C���̐�
    Dim tempfiName As String    '�ꎞ�I�ȃG�N�Z���t�@�C�����ۊǉӏ�
    
'-----�t�H���_�p�X�E�t�@�C���p�X�擾�i3�p�^�[���j-----
    '�t�H���_�p�X�I��1�F�_�C�A���O����I��
    With Application.FileDialog(msoFileDialogFolderPicker) '�t�H���_�I����ʂ�\��
        If .Show = 0 Then '���I���̏ꍇ
            Exit Function '�}�N�����I��
        Else '�I�������ꍇ
            foPath = .SelectedItems(1) '�I�������t�H���_�p�X���擾
        End If
    End With
    
    '�t�H���_�p�X�I��2�F�}�N����u�����t�H���_�Ƃ���
    'foPath = ThisWorkbook.Path
    
    '�t�H���_�p�X�I��3�F�p�X�Œ�i���ړ��́j
    'foPath = "C:\Users\Kitajima\Desktop"
    
    
    fiNum = 0
    tempfiName = Dir(foPath & "\*.xls*") '�Ώۃt�H���_�̍ŏ��̃t�@�C����
    
'-----�t�H���_���̃G�N�Z���t�@�C���������ׂĎ擾-----
    Do While tempfiName <> "" '�t�H���_�ɃG�N�Z���t�@�C��������ꍇ
        ReDim Preserve fiName(fiNum)
        fiName(fiNum) = tempfiName
        tempfiName = Dir '���̃t�@�C���̌���
        fiNum = fiNum + 1
    Loop
    
End Function

Function ProcessBooks(foPath As String, fiName() As String, targetSheet As String, resultFile As Workbook, resultSheet As Worksheet, sCell As String, eCell As String)
    Dim i As Integer
    Dim sum As Double
    Dim fCell As Double
    Dim rNum As Integer '���O�o�͍s
    
    sum = 0
    rNum = 1
    
    '�t�@�C�����݂̂̔�r
    '�i�t�H���_�p�X�͔�r���Ȃ����A�t�@�C����������̃t�@�C�����J���ƃG���[�ƂȂ邽�߉��v�j
    'SelectBooks�Ŏ擾�����t�@�C���������łɊJ����Ă��Ȃ����`�F�b�N����
    'Filter : fiName����wb.Name���܂܂�Ă��Ȃ���-1��Ԃ�
    'End : �v���O�����S�̂��I��
    For Each wb In Workbooks
        If UBound(Filter(fiName, wb.Name)) <> -1 Then
            MsgBox "�����Ώۂ̃t�@�C�������łɊJ����Ă��邽�ߏ����𒆎~���܂���", vbCritical
            End
        End If
    Next wb
    
    For i = 0 To UBound(fiName)
        Workbooks.Open foPath & "\" & fiName(i) '�J��
        Call WriteLog(resultFile, rNum, "�t�@�C�� " & fiName(i) & " ���J���܂���")
        Worksheets(targetSheet).Range(sCell & ":" & eCell).Select
        For Each targetCell In Selection.Cells
            ' �Z�����w�肵�āA�l��Ԃ��iOwner kinoshita�j
            fCell = Kagebunshin("�e�X�g", targetCell.Address)
            Call WriteLog(resultFile, rNum, "�t�@�C�� " & fiName(i) & " �� " & targetCell.Address & " �Z����ǂݍ��݂܂���")
            ' �擾�����l�𑫂��ďo�͂���iOwner ooba�j
            Call Sumcells(fCell, resultSheet, targetCell.Address)
            Call WriteLog(resultFile, rNum, "�o�͐�t�@�C���� " & targetCell.Address & " �ɒl���������݂܂���")
        Next
        Workbooks(fiName(i)).Close savechanges:=False   '�㏑�������t�@�C�������
    Next i
    
End Function
    
