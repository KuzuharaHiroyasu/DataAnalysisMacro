Attribute VB_Name = "Clear"
Sub Clear()
    Dim cnt As Long
    
    '�O���t�폜'
    If Sheets(retSheetName).ChartObjects.Count > 0 Then
        Sheets(retSheetName).ChartObjects.Delete
    End If
    
    '�f�[�^�V�[�g'
    cnt = 2 '2�s�ڂ���'
    While IsEmpty(Sheets(dataSheetName).Cells(cnt, rawRow)) = False
        Sheets(dataSheetName).Rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    '���ʃV�[�g'
    Sheets(retSheetName).Rows(3).Clear
'    Sheets(retSheetName).Range("C3").Clear
'    Sheets(retSheetName).Range("D3").Clear
'    Sheets(retSheetName).Range("E3").Clear
'    Sheets(retSheetName).Range("F3").Clear

    cnt = 44    '44�s�ڂ���'
    While IsEmpty(Sheets(retSheetName).Cells(cnt, retStartTimeRow)) = False
        Sheets(retSheetName).Rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    MsgBox "�폜�������܂����B"
End Sub

