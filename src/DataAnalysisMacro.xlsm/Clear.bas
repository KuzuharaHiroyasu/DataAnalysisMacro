Attribute VB_Name = "Clear"
Sub dataClear()
    Dim cnt As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    '�O���t�폜'
    If ChartObjects.Count > 0 Then
        ChartObjects.Delete
    End If
    
    '���ʃV�[�g'
    rows(3).Clear
    rows(9).Clear
    rows(14).Clear
    rows(19).Clear
    rows(24).Clear
    rows(28).Clear
    rows(32).Clear
    rows(36).Clear
    rows(40).Clear

    cnt = 44    '44�s�ڂ���'
    While IsEmpty(.Cells(cnt, constRetStartTimeRow)) = False
        rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "�폜�������܂����B"
End Sub

Sub retClear()
    Dim endLine As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    '�f�[�^�V�[�g'
    endLine = Cells(rows.Count, 2).End(xlUp).row
    Range(Cells(2, 1), Cells(endLine, 17)).Clear
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    MsgBox "�폜�������܂����B"
End Sub
