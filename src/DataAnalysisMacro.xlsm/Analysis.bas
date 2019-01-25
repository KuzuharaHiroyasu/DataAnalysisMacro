Attribute VB_Name = "Analysis"
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
            
            
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '�P�O�Ŗ��ċz���肠��A�������͂��т����肠�肾����'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '���ʓ��͂����̍s��'
            End If
            
            '�ʏ�ċz�̃g�[�^������'
            
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
    
    '��ԕʂ̑̂̌����̎���'
    
    'no�~20+20' '���т�1�̍s��T���˂�����no�~20+20�s�ڂ̌��������߂ĉ��Z�A�Ō�ɊY���̉ӏ��ɑ��'
    
    
    ''''''�����x�Z���T�[''''''
    '����'
    Call acceAnalysis
    
    Dim endLine As Long
    Dim i As Long
    i = 1
    
    '��������̍ŏI�s'
    endLine = Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow).End(xlUp).Row
    
    '�ŏI�̌����̍s������'
    While i <= 7
        If endLine <= Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow + i).End(xlUp).Row Then
            endLine = Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow + i).End(xlUp).Row
        End If
        i = i + 1
    Wend
    
    ''''''�O���t�쐬''''''
    '���ɃO���t������Έ�U�폜'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
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
'�̂̌��������߂�
'
Sub acceAnalysis()
    Dim x As Integer                    '�����x�Z���T�[_X��'
    Dim y As Integer                    '�����x�Z���T�[_Y��'
    Dim z As Integer                    '�����x�Z���T�[_Z��'
    Dim line As Long
    Dim x_abs As Integer
    Dim z_abs As Integer
    
    line = constInitDataLine
    
    While IsEmpty(Sheets(constDataSheetName).Cells(line, constAcceXRow)) = False
        x = Sheets(constDataSheetName).Cells(line, constAcceXRow).Value
        y = Sheets(constDataSheetName).Cells(line, constAcceYRow).Value
        z = Sheets(constDataSheetName).Cells(line, constAcceZRow).Value
        
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
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '�E(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '����'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '�E��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '�E(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
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
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '����(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    End If
                End If
            Else
                '����'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '����(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '����(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 5).Value = 2
                    End If
                End If
            End If
        End If
        line = line + 1
    Wend
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
'�O���t�쐬
'
Sub createGraph(ByVal endLine As Long)
'���т�/�ċz�̑傫��'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawSnoreRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow), Sheets(constDataSheetName).Cells(Rows.Count, constRawSnoreRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H7").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H7").Left
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
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow), Sheets(constDataSheetName).Cells(Rows.Count, constApneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H19").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H19").Left
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
            .ChartArea.Left = Sheets(constRetSheetName).Range("H31").Left
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
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow), Sheets(constDataSheetName).Cells(Rows.Count, constAcceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H43").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H43").Left
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























