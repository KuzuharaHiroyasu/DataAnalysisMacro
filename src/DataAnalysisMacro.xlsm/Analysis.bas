Attribute VB_Name = "Analysis"
'
'�e�����̎��ԗp�\����
'
Type directionTime
    up As Integer
    rightUp As Integer
    right As Integer
    rightDown As Integer
    down As Integer
    leftDown As Integer
    left As Integer
    leftUp As Integer
End Type

'
'�f�[�^���
'
Sub dataAnalysis()
    Dim snoreState As Integer           '���т�������'
    Dim apneaState As Integer           '���ċz������'
    Dim beforeSnoreState As Integer     '�P�O�̂��т�������'
    Dim beforeApneaState As Integer     '�P�O�̖��ċz������'
    Dim time As Integer                 '�o�ߎ���(�b)'
    Dim snoreCnt As Integer             '���т���'
    Dim apneaCnt As Integer             '���ċz��'
    Dim startTime As Date               '�J�n����'
    Dim dataLine As Long                '���݉�͒��̃f�[�^�̍s'
    Dim retLine As Long                 '���݌��ʓ��͒��̍s'
    Dim no As Long                      '�i���o�['
    Dim remark As Long                  '���l�p'
    Dim breath As directionTime         '�ʏ�ċz�̌����\����'
    Dim snore As directionTime          '���т��̌����\����'
    Dim apnea As directionTime          '���ċz�̌����\����'
    
    ''''''������''''''
    snoreState = 0
    apneaState = 0
    time = 0
    snoreCnt = 0
    apneaCnt = 0
    no = 0
    dataLine = constInitDataLine
    retLine = constInitRetLine
    
    ''''''�J�n�����ݒ�''''''
    startTime = Sheets(constRetSheetName).Range("B3").Value
    
    ''''''���''''''
    While IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow)) = False
        Sheets(constDataSheetName).Cells(dataLine, constNoRow).Value = no '�i���o�[�}��'
        beforeSnoreState = snoreState                       '�P�O�̂��т������Ԃ�ۑ�'
        beforeApneaState = apneaState                       '�P�O�̖��ċz�����Ԃ�ۑ�'
        snoreState = Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow).Value   '���т���Ԏ擾'
        apneaState = Sheets(constDataSheetName).Cells(dataLine, constApneaStateRow).Value   '���ċz��Ԏ擾'
        
        '�ċz�̈ړ�����'
        Call movAverage(dataLine)
        
        If snoreState = 1 Then
        '���т����肠��'
            If beforeApneaState = 1 Or beforeApneaState = 2 Then
            '�P�O�Ŗ��ċz���肠�肾����'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
        
            If beforeSnoreState = 0 Then
            '�P�O�͂��т�����Ȃ�'
                Call setStart(retLine, startTime, time, constSnore)
                snoreCnt = snoreCnt + 1
                remark = no
            End If
            
            '���т��̃g�[�^������'
            Call calculationDirectionTime(no, snore)
        ElseIf apneaState = 1 Or apneaState = 2 Then
        '���ċz���肠��'
            If beforeSnoreState = 1 Then
            '�P�O�ł��т����肠�肾����'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
        
            If beforeApneaState = 0 Then
            '�P�O�͖��ċz����Ȃ�'
                Call setStart(retLine, startTime, time, constApnea)
                apneaCnt = apneaCnt + 1
                remark = no
            End If
            
            '���ċz�̃g�[�^������'
            Call calculationDirectionTime(no, apnea)
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '�P�O�Ŗ��ċz���肠��A�������͂��т����肠�肾����'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
            
            '�ʏ�ċz�̃g�[�^������'
            Call calculationDirectionTime(no, breath)
        End If
        
        no = no + 1
        time = time + 10    '���Ԃ�10�b���₷'
        dataLine = dataLine + 1     '���̍s�̃f�[�^02��'
    Wend
    
    If IsEmpty(Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value) = False Then
        '�Ō�̔���̒�~�����Ȃ�'
        Call setRemarks(retLine, startTime, time, remark, no)
    End If
    
    ''''''�e��f�[�^�L��''''''
    Call setData(time, startTime, snoreCnt, apneaCnt)
    
    '�ʏ�ċz'
    Call setDirectionTime(breath, 9, 3)
    
    '���т�'
    Call setDirectionTime(snore, 14, 3)
    
    '���ċz'
    Call setDirectionTime(apnea, 19, 3)
    
    ''''''�����x�Z���T�[''''''
    Dim endLine As Long
    Dim i As Long
    i = 1
    
    '��������̍ŏI�s'
    endLine = Sheets(constDataSheetName).Cells(rows.Count, constRetAcceRow).End(xlUp).row
    
    '�ŏI�̌����̍s������'
    While i <= 7
        If endLine <= Sheets(constDataSheetName).Cells(rows.Count, constRetAcceRow + i).End(xlUp).row Then
            endLine = Sheets(constDataSheetName).Cells(rows.Count, constRetAcceRow + i).End(xlUp).row
        End If
        i = i + 1
    Wend
    
    ''''''�O���t�쐬''''''
    '���ɃO���t������Έ�U�폜'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    '�O���t�쐬'
    Call createGraph(endLine)
    
    MsgBox "�������܂����B"
End Sub

'
'���т��E���ċz�̊J�n�����Z�b�g
'
Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value = DateAdd("s", time, startTime)   '�J�n�����Z�b�g'
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value = kind                                '��ʃZ�b�g'
End Sub

'
'���т��E���ċz�̏I�������Z�b�g
'
Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value = DateAdd("s", time, startTime)   '��~�����Z�b�g'
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value - Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value   '�p������'
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '�p�����ԏ����ݒ�'
    If retLine = constInitRetLine Then
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = "-"
    Else
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value - Sheets(constRetSheetName).Cells(retLine - 1, constRetStopTimeRow).Value '�O���~���獡�񔭐��܂ł̎���'
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '�O���~���獡�񔭐��܂ł̎��ԏ����ݒ�'
    End If
End Sub


'
'�ړ����ς����߂�
'
Sub movAverage(ByVal dataLine As Long)
    If no >= 4 Then
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = WorksheetFunction.Sum(Range(Sheets(constDataSheetName).Cells(dataLine - 4, constRawRow), Sheets(constDataSheetName).Cells(dataLine, constRawRow))) / 5
    Else
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = "-"
    End If
End Sub

'
'���l���ɋL��
'
Sub setRemarks(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal remark As Long, ByVal no As Long)
    Call setEnd(retLine, startTime, time)
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).Value = remark & "����" & no
End Sub

'
'�e��f�[�^���Z�b�g����
'
Sub setData(ByVal time As Long, ByVal startTime As Date, ByVal snoreCnt As Integer, ByVal apneaCnt As Integer)
    '�I������'
    Sheets(constRetSheetName).Range("C3").Value = DateAdd("s", time, startTime)
    
    '�f�[�^�擾����'
    Sheets(constRetSheetName).Range("D3").Value = CStr(CDate(DateDiff("s", startTime, Sheets(constRetSheetName).Range("C3").Value) / 86400#))
    
    '���т���'
    Sheets(constRetSheetName).Range("E3").Value = snoreCnt
    
    '���ċz��'
    Sheets(constRetSheetName).Range("F3").Value = apneaCnt
End Sub

'
'�e��Ԃ��Ƃ̊e�����̎��Ԃ��Z�b�g����
'
Sub setDirectionTime(directTime As directionTime, ByVal line As Integer, ByVal row As Integer)
    '��'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.up)
    row = row + 1
    
    '�E��'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.rightUp)
    row = row + 1
    
    '�E'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.right)
    row = row + 1
    
    '�E��'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.rightDown)
    row = row + 1
    
    '��'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.down)
    row = row + 1
    
    '����'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.leftDown)
    row = row + 1
    
    '��'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.left)
    row = row + 1
    
    '����'
    Sheets(constRetSheetName).Cells(line, row).Value = DateAdd("s", time, directTime.leftUp)
End Sub

'
'�O���t�쐬
'
Sub createGraph(ByVal endLine As Long)
'���т�/�ċz�̑傫��'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawSnoreRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow), Sheets(constDataSheetName).Cells(rows.Count, constRawSnoreRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H7").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("H7").left
            .SeriesCollection(1).Name = "=""���т�"""
            .SeriesCollection(2).Name = "=""�ċz��"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1024
            .Axes(xlValue).MajorUnit = 256
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/�W��"
        End With
    End If
    
    '���т�/�ċz�̔���'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constApneaStateRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow), Sheets(constDataSheetName).Cells(rows.Count, constApneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H19").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("H19").left
            .SeriesCollection(1).Name = "=""���т�"""
            .SeriesCollection(2).Name = "=""���ċz"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 2
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/�W��"
        End With
    End If
    
    '�̂̌���'
    If endLine > 1 Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine - 1, constRetAcceRow), Sheets(constDataSheetName).Cells(endLine, 17))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H31").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("H31").left
            .SeriesCollection(1).Name = "=""��"""
            .SeriesCollection(2).Name = "=""�E��"""
            .SeriesCollection(3).Name = "=""�E"""
            .SeriesCollection(4).Name = "=""�E��"""
            .SeriesCollection(5).Name = "=""��"""
            .SeriesCollection(6).Name = "=""����"""
            .SeriesCollection(7).Name = "=""��"""
            .SeriesCollection(8).Name = "=""����"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 7
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/�W��"
        End With
    End If
    
    '�Z���T�[�l'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceYRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceZRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow), Sheets(constDataSheetName).Cells(rows.Count, constAcceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H43").Top
            .ChartArea.left = Sheets(constRetSheetName).Range("H43").left
            .SeriesCollection(1).Name = "=""�w��"""
            .SeriesCollection(2).Name = "=""�x��"""
            .SeriesCollection(3).Name = "=""�y��"""
            .Axes(xlValue).MinimumScale = -100
            .Axes(xlValue).MaximumScale = 100
            .Axes(xlValue).MajorUnit = 50
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/�W��"
        End With
    End If
End Sub

'
'�e��Ԃ��Ƃ̊e�����̎��Ԃ����߂�
'
Sub calculationDirectionTime(ByVal no As Long, directTime As directionTime)
    Dim line As Long
    Dim rows As Integer
    
    rows = 10
    
    '�Y���̌����̍s'
    line = (no * 20) + 20

    '����������'
    While WorksheetFunction.CountA(Sheets(constDataSheetName).Cells(line, rows)) = 0
        '��'
        rows = rows + 1
    Wend

    Select Case rows
        Case 10 'J��(��)'
            directTime.up = directTime.up + 10
        Case 11 'K��(�E��)'
            directTime.rightUp = directTime.rightUp + 10
        Case 12 'L��(�E)'
            directTime.right = directTime.right + 10
        Case 13 'M��(�E��)'
            directTime.rightDown = directTime.rightDown + 10
        Case 14 'N��(��)'
            directTime.down = directTime.down + 10
        Case 15 'O��(����)'
            directTime.leftDown = directTime.leftDown + 10
        Case 16 'P��(��)'
            directTime.left = directTime.left + 10
        Case 17 'Q��(����)'
            directTime.leftUp = directTime.leftUp + 10
    End Select
End Sub





















