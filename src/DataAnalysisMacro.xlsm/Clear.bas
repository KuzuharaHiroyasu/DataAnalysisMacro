Attribute VB_Name = "Clear"
Sub Clear()
    Dim cnt As Long
    
    '�O���t�폜'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    
    '�f�[�^�V�[�g'
    cnt = 2 '2�s�ڂ���'
    While IsEmpty(Sheets(constDataSheetName).Cells(cnt, constRawRow)) = False
        Sheets(constDataSheetName).rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    '���ʃV�[�g'
    Sheets(constRetSheetName).rows(3).Clear
'    Sheets(constRetSheetName).Range("C3").Clear
'    Sheets(constRetSheetName).Range("D3").Clear
'    Sheets(constRetSheetName).Range("E3").Clear
'    Sheets(constRetSheetName).Range("F3").Clear

    cnt = 44    '44�s�ڂ���'
    While IsEmpty(Sheets(constRetSheetName).Cells(cnt, constRetStartTimeRow)) = False
        Sheets(constRetSheetName).rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    MsgBox "�폜�������܂����B"
End Sub

