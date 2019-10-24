Attribute VB_Name = "Clear"
Sub retClear()
    Dim cnt As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    '�O���t�폜'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    
    '���ʃV�[�g'
    Sheets(constRetSheetName).rows(3).Clear
    Sheets(constRetSheetName).rows(9).Clear
    Sheets(constRetSheetName).rows(14).Clear
    Sheets(constRetSheetName).rows(19).Clear
    Sheets(constRetSheetName).rows(24).Clear
    Sheets(constRetSheetName).rows(28).Clear
    Sheets(constRetSheetName).rows(32).Clear
    Sheets(constRetSheetName).rows(36).Clear
    Sheets(constRetSheetName).rows(40).Clear
    
    '�������Ԃ̕\�̘g���𕜊�������'
    Sheets(constRetSheetName).Range("B9:J20").BorderAround LineStyle:=xlContinuous
 
    cnt = 44    '44�s�ڂ���'
    While IsEmpty(Sheets(constRetSheetName).Cells(cnt, constRetStartTimeRow)) = False
        Sheets(constRetSheetName).rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    'copy�V�[�g�̏o�͌��ʍ폜'
'    Sheets(constCopySheetName).rows(1).Clear
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox "�폜�������܂����B"
End Sub

Sub dataClear()
    Dim endLine As Long
    Dim lineCnt As Long

    endLine = 1
    lineCnt = 1
    
    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    '�f�[�^�V�[�g'
    While endLine < 2
        endLine = Cells(rows.Count, lineCnt).End(xlUp).row
        If lineCnt > 18 Then
            endLine = 2
            GoTo Break
        End If
        lineCnt = lineCnt + 1
    Wend

Break:
    Range(Cells(2, 1), Cells(endLine, 18)).Clear
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
'    MsgBox "�폜�������܂����B"
End Sub
