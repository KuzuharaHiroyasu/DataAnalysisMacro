Attribute VB_Name = "Read"
'
'�f�[�^�ǂݍ���
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    Application.Calculation = xlManual
    
    '�ċz��'
    If Not readText(ThisWorkbook.Path & "\raw_sum.txt", constRawRow) Then
        msg = "raw_sum.txt "
    End If
    
    '�S��������̌ċz��'
    If Not readText(ThisWorkbook.Path & "\raw_heartBeatRemov_sum.txt", constRawHBRemovRow) Then
        msg = "raw_heartBeatRemov_sum.txt "
    End If
    
    '���т���'
    If Not readText(ThisWorkbook.Path & "\rawsnore_sum.txt", constRawSnoreRow) Then
        msg = msg + "rawsnore_sum.txt "
    End If
    
    '���ċz���'
    If Not readText(ThisWorkbook.Path & "\apnea_sum.txt", constApneaStateRow) Then
        msg = msg + "apnea_sum.txt "
    End If
    
    '���т����'
    If Not readText(ThisWorkbook.Path & "\snore__sum.txt", constSnoreStateRow) Then
        msg = msg + "snore__sum.txt "
    End If
    
    '�t�H�g�Z���T�[�l'
    If Not readText(ThisWorkbook.Path & "\photoref_sum.txt", constPhotorefRow) Then
        msg = msg + "photoref_sum.txt "
    End If
    
    'X��'
    If Not readText(ThisWorkbook.Path & "\acce_x_sum.txt", constAcceXRow) Then
        msg = msg + "acce_x_sum.txt "
    End If
    
    'Y��'
    If Not readText(ThisWorkbook.Path & "\acce_y_sum.txt", constAcceYRow) Then
        msg = msg + "acce_y_sum.txt "
    End If
    
    'Z��'
    If Not readText(ThisWorkbook.Path & "\acce_z_sum.txt", constAcceZRow) Then
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
    Dim x As Long                    '�����x�Z���T�[_X��'
    Dim y As Long                    '�����x�Z���T�[_Y��'
    Dim z As Long                    '�����x�Z���T�[_Z��'
    Dim line As Long
    Dim x_abs As Integer
    Dim z_abs As Integer
    
    line = constInitDataLine
    
    While IsEmpty(Sheets(constDataSheetName).Cells(line, constAcceXRow)) = False
        DoEvents
        x = Sheets(constDataSheetName).Cells(line, constAcceXRow).Value
        y = Sheets(constDataSheetName).Cells(line, constAcceYRow).Value
        z = Sheets(constDataSheetName).Cells(line, constAcceZRow).Value
        
        If x > 200 Or y > 200 Or z > 200 Then
            '�G���[���(�C���M�����[�Ȓl)'
            Sheets(constDataSheetName).Range(Cells(line, constAcceXRow), Cells(line, constAcceZRow)).Delete
            GoTo Continue
        End If
        
        x_abs = Abs(x)
        z_abs = Abs(z)
        
        '�w�b�h���̏ꏊ�i�̂̌����ł͂Ȃ��j'
        If 0 <= x Then
            '��or��'
            If 0 <= z Then
                '��'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 6).Value = 1
            Else
                '��'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow).Value = 7
            End If
        Else
            '��or�E'
            If 0 <= z Then
                '��'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 4).Value = 3
            Else
                '�E'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 2).Value = 5
            End If
        End If
        line = line + 1
Continue:
    Wend
End Sub

