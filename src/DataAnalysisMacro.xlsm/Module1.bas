Attribute VB_Name = "Module1"
'�萔'
'�f�[�^�V�[�g'
Const dataSheetName = "�f�[�^"  '�V�[�g��'
'�s'
Const initDataLine = 2          '�f�[�^�̍ŏ��̍s'
'��'
Const noRow = 1                 'No�̗�(A��)'
Const rawRow = 2                '�ċz���̗�(B��)'
Const rawSnoreRow = 3           '���т����̗�(C��)'
Const rawMovAvrRow = 4          '�ċz���̈ړ����ς̗�(D��)'
Const snoreStateRow = 5         '���т����茋�ʂ̓�������(E��)'
Const apneaStateRow = 6         '���ċz���茋�ʂ̓�������(F��)'
Const acceXRow = 7              '�����x(X)�̓�������(G��)'
Const acceYRow = 8              '�����x(Y)�̓�������(H��)'
Const acceZRow = 9              '�����x(Z)�̓�������(I��)'
Const retAcceRow = 10           '����(J��)'


'���ʃV�[�g'
Const retSheetName = "����"     '�V�[�g��'
'�s'
Const initRetLine = 7           '���ʂ����͂����ŏ��̍s'
'��'
Const retStartTimeRow = 2       '���莞���̗�(B��)'
Const retStopTimeRow = 3        '��~�����̗�(C��)'
Const retContinuTimeRow = 4     '�p�����Ԃ̗�(D��)'
Const retTypeRow = 5            '��ʂ̗�(E��)'
Const retStartFromStopTimeRow = 6   '�O���~���獡�񔭐��܂ł̎��Ԃ̗�(F��)'
Const retRemarkRow = 7          '���l�̗�(F��)'
'���͕���'
Const snore = "���т�"
Const apnea = "���ċz"

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
    
    '������'
    snoreState = 0
    apneaState = 0
    time = 0
    snoreCnt = 0
    apneaCnt = 0
    no = 1
    dataLine = initDataLine
    retLine = initRetLine
    
    '�J�n�����ݒ�'
    startTime = Sheets(retSheetName).Range("B3").Value
    
    '���'
    While IsEmpty(Sheets(dataSheetName).Cells(dataLine, snoreStateRow)) = False
        Sheets(dataSheetName).Cells(dataLine, noRow).Value = no 'No�}��'
        beforeSnoreState = snoreState                       '�P�O�̂��т������Ԃ�ۑ�'
        beforeApneaState = apneaState                       '�P�O�̖��ċz�����Ԃ�ۑ�'
        snoreState = Sheets(dataSheetName).Cells(dataLine, snoreStateRow).Value   '���т���Ԏ擾'
        apneaState = Sheets(dataSheetName).Cells(dataLine, apneaStateRow).Value   '���ċz��Ԏ擾'
        
        '�ċz�̈ړ�����'
        If no >= 5 Then
            Sheets(dataSheetName).Cells(dataLine, rawMovAvrRow).Value = WorksheetFunction.Sum(Range(Sheets(dataSheetName).Cells(dataLine - 4, rawRow), Sheets(dataSheetName).Cells(dataLine, rawRow))) / 5
        Else
            Sheets(dataSheetName).Cells(dataLine, rawMovAvrRow).Value = "-"
        End If
        
        
        If snoreState = 1 Then
        '���т����肠��'
            If beforeApneaState = 1 Or beforeApneaState = 2 Then
            '�P�O�Ŗ��ċz���肠�肾����'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "����" & no
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
        
            If beforeSnoreState = 0 Then
            '�P�O�͂��т�����Ȃ�'
                Call setStart(retLine, startTime, time, snore)
                snoreCnt = snoreCnt + 1
                remark = no
            End If
        ElseIf apneaState = 1 Or apneaState = 2 Then
        '���ċz���肠��'
            If beforeSnoreState = 1 Then
            '�P�O�ł��т����肠�肾����'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "����" & no
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
        
            If beforeApneaState = 0 Then
            '�P�O�͖��ċz����Ȃ�'
                Call setStart(retLine, startTime, time, apnea)
                apneaCnt = apneaCnt + 1
                remark = no
            End If
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '�P�O�Ŗ��ċz���肠��A�������͂��т����肠�肾����'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "����" & no
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
        End If
        
        no = no + 1
        time = time + 10    '���Ԃ�10�b���₷'
        dataLine = dataLine + 1     '���̍s�̃f�[�^02��'
    Wend
    
    If IsEmpty(Sheets(retSheetName).Cells(retLine, retTypeRow).Value) = False Then
        '�Ō�̔���̒�~�����Ȃ�'
        Call setEnd(retLine, startTime, time)
        Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "����" & no
    End If
    
    '�I������'
    Sheets(retSheetName).Range("C3").Value = DateAdd("s", time, startTime)
    
    '�f�[�^�擾����'
    Sheets(retSheetName).Range("D3").Value = CStr(CDate(DateDiff("s", startTime, Sheets(retSheetName).Range("C3").Value) / 86400#))
    
    '���т���'
    Sheets(retSheetName).Range("E3").Value = snoreCnt
    
    '���ċz��'
    Sheets(retSheetName).Range("F3").Value = apneaCnt
    
    '�O���t�폜(��x�폜����)'
    If Sheets(retSheetName).ChartObjects.Count > 0 Then
        Sheets(retSheetName).ChartObjects.Delete
    End If
    
    '�O���t�쐬'
    '���т�/�ċz�̑傫��'
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, rawRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, rawSnoreRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, rawRow), Sheets(dataSheetName).Cells(Rows.Count, rawSnoreRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H7").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H7").Left
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
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, snoreStateRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, apneaStateRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, snoreStateRow), Sheets(dataSheetName).Cells(Rows.Count, apneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H19").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H19").Left
            .SeriesCollection(1).Name = "=""���т�"""
            .SeriesCollection(2).Name = "=""���ċz"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 2
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/�W��"
        End With
    End If
    
    '�����x�Z���T�['
    '����'
    Call acceAnalysis
    
    Dim endLine As Long               '��������̍ŏI�s'
    Dim i As Long
    i = 1
    endLine = Sheets(dataSheetName).Cells(Rows.Count, retAcceRow).End(xlUp).Row
    
    '�ŏI�̌����̍s������'
    While i <= 7
        If endLine <= Sheets(dataSheetName).Cells(Rows.Count, retAcceRow + i).End(xlUp).Row Then
            endLine = Sheets(dataSheetName).Cells(Rows.Count, retAcceRow + i).End(xlUp).Row
        End If
        i = i + 1
    Wend
    
    If endLine > 1 Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine - 1, retAcceRow), Sheets(dataSheetName).Cells(endLine, 17))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H31").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H31").Left
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
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceXRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceYRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceZRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, acceXRow), Sheets(dataSheetName).Cells(Rows.Count, acceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H43").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H43").Left
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
    
    MsgBox "�������܂����B"
End Sub

Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
    Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value = DateAdd("s", time, startTime)   '�J�n�����Z�b�g'
    Sheets(retSheetName).Cells(retLine, retStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    Sheets(retSheetName).Cells(retLine, retTypeRow).Value = kind                                '��ʃZ�b�g'
End Sub

Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Sheets(retSheetName).Cells(retLine, retStopTimeRow).Value = DateAdd("s", time, startTime)   '��~�����Z�b�g'
    Sheets(retSheetName).Cells(retLine, retStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '���������ݒ�'
    Sheets(retSheetName).Cells(retLine, retContinuTimeRow).Value = Sheets(retSheetName).Cells(retLine, retStopTimeRow).Value - Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value   '�p������'
    Sheets(retSheetName).Cells(retLine, retContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '�p�����ԏ����ݒ�'
    If retLine = initRetLine Then
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).Value = "-"
    Else
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).Value = Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value - Sheets(retSheetName).Cells(retLine - 1, retStopTimeRow).Value '�O���~���獡�񔭐��܂ł̎���'
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '�O���~���獡�񔭐��܂ł̎��ԏ����ݒ�'
    End If
End Sub

Sub dataAndResultClear()
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

    cnt = 7
    While IsEmpty(Sheets(retSheetName).Cells(cnt, retStartTimeRow)) = False
        Sheets(retSheetName).Rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    MsgBox "�폜�������܂����B"
End Sub

Sub acceAnalysis()
    Dim x As Integer                    '�����x�Z���T�[_X��'
    Dim y As Integer                    '�����x�Z���T�[_Y��'
    Dim z As Integer                    '�����x�Z���T�[_Z��'
    Dim line As Long
    Dim x_abs As Integer
    Dim z_abs As Integer
    
    line = initDataLine
    
    While IsEmpty(Sheets(dataSheetName).Cells(line, acceXRow)) = False
        x = Sheets(dataSheetName).Cells(line, acceXRow).Value
        y = Sheets(dataSheetName).Cells(line, acceYRow).Value
        z = Sheets(dataSheetName).Cells(line, acceZRow).Value
        
        x_abs = Abs(x)
        z_abs = Abs(z)
        
        '�w�b�h���̏ꏊ�i�̂̌����ł͂Ȃ��j'
        If 0 <= x Then
            '�E��'
            If 0 <= z Then
                '�㑤'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    Else
                        '�E(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '����'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    Else
                        '�E(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    End If
                End If
            End If
        Else
            '����'
            If 0 <= z Then
                '�㑤'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '����(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '����(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    End If
                End If
            Else
                '����'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '����(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '����(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    Else
                        '��(�m)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 5).Value = 2
                    End If
                End If
            End If
        End If
        line = line + 1
    Wend
End Sub

Sub absoluteValue(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    x = Abs(x)
    y = Abs(y)
    z = Abs(z)
End Sub

Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    '�ċz��'
    If Not readText(ThisWorkbook.Path & "\raw_sum.txt", 2) Then
        msg = "raw_sum.txt "
    End If
    
    '���т���'
    If Not readText(ThisWorkbook.Path & "\rawsnore_sum.txt", 3) Then
        msg = msg + "rawsnore_sum.txt "
    End If
    
    '���т����'
    If Not readText(ThisWorkbook.Path & "\snore__sum.txt", 5) Then
        msg = msg + "snore__sum.txt "
    End If
    
    '���ċz���'
    If Not readText(ThisWorkbook.Path & "\apnea_sum.txt", 6) Then
        msg = msg + "apnea_sum.txt "
    End If
    
    'X��'
    If Not readText(ThisWorkbook.Path & "\acce_x_sum.txt", 7) Then
        msg = msg + "acce_x_sum.txt "
    End If
    
    'Y��'
    If Not readText(ThisWorkbook.Path & "\acce_y_sum.txt", 8) Then
        msg = msg + "acce_y_sum.txt "
    End If
    
    'Z��'
    If Not readText(ThisWorkbook.Path & "\acce_z_sum.txt", 9) Then
        msg = msg + "acce_z_sum.txt "
    End If
    
    If Not msg = "" Then
        msg = msg + "��ǂݍ��߂܂���ł����B"
    Else
        msg = "�������܂����B"
    End If
    
    MsgBox msg
End Sub


Public Function readText(ByVal fileName As String, ByVal inputRow As Long) As Boolean
    Dim a
    Dim inputLine As Long
    
    inputLine = 2
    
    a = Dir(fileName)
    If (a <> "") Then
        Open fileName For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            Sheets(dataSheetName).Cells(inputLine, inputRow) = buf
            inputLine = inputLine + 1
        Loop
        Close #1
        readText = True
    Else
        readText = False
    End If
End Function
