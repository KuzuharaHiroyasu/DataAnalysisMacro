Attribute VB_Name = "Read"
'
'�f�[�^�ǂݍ���
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    Application.Calculation = xlManual
    
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
    
    Call acceAnalysis
    
    Application.Calculation = xlAutomatic

    If Not msg = "" Then
        msg = msg + "��ǂݍ��߂܂���ł����B"
    Else
        msg = "�������܂����B"
    End If
    
    MsgBox msg
End Sub

'
'�f�[�^�e�L�X�g�ǂݍ���
'
Public Function readText(ByVal fileName As String, ByVal inputRow As Long) As Boolean
    Dim a
    Dim inputLine As Long
    
    inputLine = 2
    
    a = Dir(fileName)
    If (a <> "") Then
        Open fileName For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            Sheets(constDataSheetName).Cells(inputLine, inputRow) = buf
            'DoEvents
            inputLine = inputLine + 1
        Loop
        Close #1
        readText = True
    Else
        readText = False
    End If
End Function

'
'�����x�Z���T�[�l�̐�Βl
'
Sub absoluteValue(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    x = Abs(x)
    y = Abs(y)
    z = Abs(z)
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
        DoEvents
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
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '�E��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '����'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '�E(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '�E��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '�E(�m)'
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
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '����(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '��(�m)'
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
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '��(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '����(�m)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 5).Value = 2
                    End If
                End If
            End If
        End If
        line = line + 1
    Wend
End Sub

