Attribute VB_Name = "Analysis"
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)

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
    Dim time As Long                    '�o�ߎ���(�b)'
    Dim snoreCnt As Integer             '���т���'
    Dim apneaCnt As Integer             '���ċz��'
    Dim startTime As Date               '�J�n����'
    Dim dataLine As Long                '���݉�͒��̃f�[�^�̍s'
    Dim retLine As Long                 '���݌��ʓ��͒��̍s'
    Dim no As Long                      '�i���o�['
    Dim lastNo As Long                  '�ŏI����i���o�['
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
    
    ''''''�����l�ݒ�''''''
    '��0������'
    Sheets(constRetSheetName).Range("B24:H24").Value = 0
    Sheets(constRetSheetName).Range("B28:H28").Value = 0
    
    ''''''�J�n�����ݒ�''''''
    startTime = Sheets(constRetSheetName).Range("B3").Value
    
    ''''''���''''''
    Application.ScreenUpdating = False
    Application.Calculation = xlManual

    While IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constApneaStateRow)) = False
        DoEvents
        Sheets(constDataSheetName).Cells(dataLine, constNoRow).Value = no '�i���o�[�}��'

        '�ċz�̈ړ�����'
'        Call movAverage(dataLine, no)'

        If IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow)) = False Then
            '���т����茋�ʂ����͂���Ă���'
            beforeSnoreState = snoreState                       '�P�O�̂��т������Ԃ�ۑ�'
            beforeApneaState = apneaState                       '�P�O�̖��ċz�����Ԃ�ۑ�'
            snoreState = Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow).Value   '���т���Ԏ擾'
            apneaState = Sheets(constDataSheetName).Cells(dataLine, constApneaStateRow).Value   '���ċz��Ԏ擾'

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
                Call calculationDirectionTime(no, dataLine, snore)
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
                Call calculationDirectionTime(no, dataLine, apnea)
            Else
                If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
                '�P�O�Ŗ��ċz���肠��A�������͂��т����肠�肾����'
                    Call setRemarks(retLine, startTime, time, remark, no)
                    retLine = retLine + 1     '���ʓ��͂����̍s��'
                End If

                '�ʏ�ċz�̃g�[�^������'
                Call calculationDirectionTime(no, dataLine, breath)
            End If
            time = time + 10    '���Ԃ�10�b���₷'
            lastNo = no + 1
        End If

        no = no + 1
        dataLine = dataLine + 1     '���̍s�̃f�[�^02��'
    Wend

    ''''''�e��f�[�^�L��''''''
    Call setData(time, startTime, snoreCnt, apneaCnt)

    '�e�������Ƃ̒ʏ�ċz�̎���'
    Call setDirectionTime(breath, 9, 2)

    '�e�������Ƃ̂��т��̎���'
    Call setDirectionTime(snore, 14, 2)

    '�e�������Ƃ̖��ċz�̎���'
    Call setDirectionTime(apnea, 19, 2)

    '�������Ԃ̊���'
    Call sleepTimeRatio

    '���т��}���̊���'
    Call perOfSuppression(24, 36, 2, Sheets(constRetSheetName).Range("E3").Value)

    '���ċz�}���̊���'
    Call perOfSuppression(28, 40, 2, Sheets(constRetSheetName).Range("F3").Value)

    '�f�[�^��1�s�ɃR�s�['
    Call copyData
    
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    
    'MsgBox "�������܂����B"
End Sub

'
'���т��E���ċz�̊J�n�����Z�b�g
'
Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
'    Sheets(constRetSheetName).Range(Cells(retLine, constRetStartTimeRow), Cells(retLine, constRetRemarkRow)).Font.Size = 10
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value = DateAdd("s", time, startTime)   '�J�n�����Z�b�g'
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value = kind                                '��ʃZ�b�g'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).HorizontalAlignment = xlCenter
End Sub

'
'���т��E���ċz�̏I�������Z�b�g
'
Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Dim kind As String
    Dim duration As Date
    
    '��~�����Z�b�g'
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value = DateAdd("s", time, startTime)
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    
    '�p�����ԃZ�b�g'
    duration = Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value - Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).Value = duration
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '�p�����ԏ����ݒ�'
    
    '�Ĕ��o�ߎ��ԃZ�b�g'
    If retLine = constInitRetLine Then
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = "-"
    Else
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value - Sheets(constRetSheetName).Cells(retLine - 1, constRetStopTimeRow).Value
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '�O���~���獡�񔭐��܂ł̎��ԏ����ݒ�'
    End If
    Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).HorizontalAlignment = xlRight

    
    '�p�����Ԃ��Ƃɉ񐔂��Z�b�g'
    '���'
    kind = Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value
    If kind = "���т�" Then
        Call setNumPerDuration(duration, 24)
    Else
        Call setNumPerDuration(duration, 28)
    End If
End Sub

'
'�p�����Ԃ��Ƃɉ񐔂��Z�b�g
'
Sub setNumPerDuration(ByVal duration As Date, ByVal line As Integer)
    Dim durationInt As Integer
    
    'Date��Integer�ɕϊ�'
    durationInt = duration * 86400

    If durationInt = 10 Then
        '10�b'
        Sheets(constRetSheetName).Cells(line, 2).Value = Sheets(constRetSheetName).Cells(line, 2).Value + 1
    ElseIf durationInt = 20 Then
        '20�b'
        Sheets(constRetSheetName).Cells(line, 3).Value = Sheets(constRetSheetName).Cells(line, 3).Value + 1
    ElseIf durationInt >= 30 And durationInt < 60 Then
        '30�b�ȏ�1������'
        Sheets(constRetSheetName).Cells(line, 4).Value = Sheets(constRetSheetName).Cells(line, 4).Value + 1
    ElseIf durationInt >= 60 And durationInt < 120 Then
        '1���ȏ�2������'
        Sheets(constRetSheetName).Cells(line, 5).Value = Sheets(constRetSheetName).Cells(line, 5).Value + 1
    ElseIf durationInt >= 120 And durationInt < 300 Then
        '2���ȏ�5������'
        Sheets(constRetSheetName).Cells(line, 6).Value = Sheets(constRetSheetName).Cells(line, 6).Value + 1
    ElseIf durationInt >= 300 And durationInt < 600 Then
        '5���ȏ�10������'
        Sheets(constRetSheetName).Cells(line, 7).Value = Sheets(constRetSheetName).Cells(line, 7).Value + 1
    Else
        '10���ȏ�'
        Sheets(constRetSheetName).Cells(line, 8).Value = Sheets(constRetSheetName).Cells(line, 8).Value + 1
    End If
End Sub

'
'���l���ɋL��
'
Sub setRemarks(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal remark As Long, ByVal lastNo As Long)
    Call setEnd(retLine, startTime, time)
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).Value = remark & "����" & lastNo
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).HorizontalAlignment = xlRight
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
    Dim time As Date
    Dim totalTime As Integer
    
    '��'
    time = TimeSerial(0, 0, directTime.up)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '�E��'
    time = TimeSerial(0, 0, directTime.rightUp)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '�E'
    time = TimeSerial(0, 0, directTime.right)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '�E��'
    time = TimeSerial(0, 0, directTime.rightDown)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '��'
    time = TimeSerial(0, 0, directTime.down)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '����'
    time = TimeSerial(0, 0, directTime.leftDown)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '��'
    time = TimeSerial(0, 0, directTime.left)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '����'
    time = TimeSerial(0, 0, directTime.leftUp)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
    row = row + 1
    
    '���v'
    totalTime = directTime.up + directTime.rightUp + directTime.right + directTime.rightDown + directTime.down + directTime.leftDown + directTime.left + directTime.leftUp
    time = TimeSerial(0, 0, totalTime)
    Sheets(constRetSheetName).Cells(line, row).NumberFormatLocal = "hh:mm:ss"
    Sheets(constRetSheetName).Cells(line, row).Value = time
End Sub

'
'�e��Ԃ��Ƃ̊e�����̎��Ԃ����߂�
'
Sub calculationDirectionTime(ByVal no As Long, ByVal dataLine As Long, directTime As directionTime)
'    Dim line As Long
    Dim rows As Integer
    
    rows = constRetAcceStartRow
    
    '�Y���̌����̍s'
'    line = (no * 20) + 20
    
    '�����x�Z���T�[�̒l���J���Ȃ��̍s�̒l������Ƃ���܂ők��'
'    While WorksheetFunction.CountA(Sheets(constDataSheetName).Cells(line, constAcceXRow)) = 0
'        line = line - 1
'    Wend
    
    '����������'
    While WorksheetFunction.CountA(Sheets(constDataSheetName).Cells(dataLine, rows)) = 0
        '��'
        rows = rows + 1
    Wend

    Select Case rows
        Case constRetAcceStartRow 'K��(��)'
            directTime.up = directTime.up + 10
        Case constRetAcceStartRow + 1 'L��(�E��)'
            directTime.rightUp = directTime.rightUp + 10
        Case constRetAcceStartRow + 2 'M��(�E)'
            directTime.right = directTime.right + 10
        Case constRetAcceStartRow + 3 'N��(�E��)'
            directTime.rightDown = directTime.rightDown + 10
        Case constRetAcceStartRow + 4 'O��(��)'
            directTime.down = directTime.down + 10
        Case constRetAcceStartRow + 5 'P��(����)'
            directTime.leftDown = directTime.leftDown + 10
        Case constRetAcceStartRow + 6 'Q��(��)'
            directTime.left = directTime.left + 10
        Case constRetAcceStartRow + 7 'R��(����)'
            directTime.leftUp = directTime.leftUp + 10
    End Select
End Sub

'
'�������Ԃ̊���
'
Sub sleepTimeRatio()
    Dim dataAcqTime As Variant '�f�[�^�擾����'
    
    dataAcqTime = Sheets(constRetSheetName).Range("D3").Value
    
    '�ʏ�ċz'
    Sheets(constRetSheetName).Range("B32").Value = Sheets(constRetSheetName).Range("J9").Value / dataAcqTime
    Sheets(constRetSheetName).Range("B32").NumberFormatLocal = "0.0%"
    
    '���т�'
    Sheets(constRetSheetName).Range("C32").Value = Sheets(constRetSheetName).Range("J14").Value / dataAcqTime
    Sheets(constRetSheetName).Range("C32").NumberFormatLocal = "0.0%"
    
    '���ċz'
    Sheets(constRetSheetName).Range("D32").Value = Sheets(constRetSheetName).Range("J19").Value / dataAcqTime
    Sheets(constRetSheetName).Range("D32").NumberFormatLocal = "0.0%"
End Sub

'
'���т��E���ċz�}���̊���
'
Sub perOfSuppression(ByVal line As Integer, ByVal retLine As Integer, ByVal row As Integer, ByVal totalCnt As Integer)
    Dim i As Integer
    '10�b�@�`�@10���ȏ�܂�7���ڕ�'
    If totalCnt = 0 Then
        'totalCnt��0'
        For i = 1 To 7
            Sheets(constRetSheetName).Cells(retLine, row).Value = Sheets(constRetSheetName).Cells(line, row).Value
            Sheets(constRetSheetName).Cells(retLine, row).NumberFormatLocal = "0.0%"
            row = row + 1
        Next i
    Else
        'totalCnt��0�ȊO'
        For i = 1 To 7
            Sheets(constRetSheetName).Cells(retLine, row).Value = Sheets(constRetSheetName).Cells(line, row).Value / totalCnt
            Sheets(constRetSheetName).Cells(retLine, row).NumberFormatLocal = "0.0%"
            row = row + 1
        Next i
    End If
End Sub

'
'��͌��ʃR�s�[
'
Sub copyData()
    Dim line As Integer
    Dim row As Integer
    
    line = 1
    row = 1
    
    Do While Len(Sheets(constCopySheetName).Cells(line, 1)) > 0
        line = line + 1
    Loop
    
    Sheets(constRetSheetName).Range("B3:F3").Copy Sheets(constCopySheetName).Cells(line, row)   '�J�n����, �I������, �f�[�^�擾����, ���т���, ���ċz�� + ���
    row = row + 6
    
    Sheets(constRetSheetName).Range("J9").Copy Sheets(constCopySheetName).Cells(line, row)      '�ʏ�ċz����
    row = row + 1
    
    Sheets(constRetSheetName).Range("J14").Copy Sheets(constCopySheetName).Cells(line, row)     '���т�����
    row = row + 1
    
    Sheets(constRetSheetName).Range("J19").Copy Sheets(constCopySheetName).Cells(line, row)     '���ċz���� + ���
    row = row + 2
    
    Sheets(constRetSheetName).Range("B24:H24").Copy Sheets(constCopySheetName).Cells(line, row) '���т����ԁi�񐔁j- 10�b, 20�b, 30�b�ȏ�1������, 1���ȏ�2������, 2���ȏ�5������, 5���ȏ�10������, 10���ȏ� + ���
    row = row + 8
    
    Sheets(constRetSheetName).Range("B28:H28").Copy Sheets(constCopySheetName).Cells(line, row) '���ċz���ԁi�񐔁j- 10�b, 20�b, 30�b�ȏ�1������, 1���ȏ�2������, 2���ȏ�5������, 5���ȏ�10������, 10���ȏ� + ���
    row = row + 8
    
    Sheets(constRetSheetName).Range("B32:D32").Copy Sheets(constCopySheetName).Cells(line, row) '���� - �ʏ�ċz, ���т�, ���ċz + ���
    row = row + 4
    
    Sheets(constRetSheetName).Range("B36:H36").Copy Sheets(constCopySheetName).Cells(line, row) '���т����ԁi�����j- 10�b, 20�b, 30�b�ȏ�1������, 1���ȏ�2������, 2���ȏ�5������, 5���ȏ�10������, 10���ȏ� + ���
    row = row + 8
    
    Sheets(constRetSheetName).Range("B40:H40").Copy Sheets(constCopySheetName).Cells(line, row) '���ċz���ԁi�����j- 10�b, 20�b, 30�b�ȏ�1������, 1���ȏ�2������, 2���ȏ�5������, 5���ȏ�10������, 10���ȏ� + ���
    row = row + 8
    
    Sheets(constRetSheetName).Range("B9:I9").Copy Sheets(constCopySheetName).Cells(line, row)   '�ʏ�ċz���� - ��, �E��, �E, �E��, ��, ����, ��, ���� + ���
    row = row + 9
    
    Sheets(constRetSheetName).Range("B14:I14").Copy Sheets(constCopySheetName).Cells(line, row)   '���т����� - ��, �E��, �E, �E��, ��, ����, ��, ���� + ���
    row = row + 9
    
    Sheets(constRetSheetName).Range("B19:I19").Copy Sheets(constCopySheetName).Cells(line, row)   '���ċz���� - ��, �E��, �E, �E��, ��, ����, ��, ���� + ���
    row = row + 9
End Sub











