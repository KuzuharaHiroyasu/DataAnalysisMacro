Attribute VB_Name = "Read"
'
'�f�[�^�ǂݍ���
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    Dim startTime As String

    Dim fileName As String
    Dim sheetNameCSV As String
    Dim Path As String
    
    Application.Calculation = xlManual
    
    '�p�X�擾'
    Path = ThisWorkbook.Path + "\"
    
    'csv�t�@�C���������t�@�C�����擾'
    fileName = Dir(Path & "*.csv")
    
    Do While Len(fileName) > 0
        '�V�[�g���擾'
        sheetNameCSV = left(fileName, 14)
        
        '�R�s�[���̃V�[�g������t�@�C�����J��
        Workbooks.Open (Path + fileName)
         
        '�V�[�g���R�s�[(�V�����t�@�C���ɍ쐬)
        Workbooks(fileName).Worksheets(sheetNameCSV).Copy After:=ThisWorkbook.Sheets(1)
         
        '�R�s�[���t�@�C�������
        Workbooks(fileName).Close
        
        'csv����f�[�^�V�[�g�Ƀf�[�^���Z�b�g'
        dataSet (sheetNameCSV)
        
        '�J�n���ԋL��'
        startTime = setStartTime(sheetNameCSV)
        ThisWorkbook.Sheets(constRetSheetName).Range("B3").Value = startTime
        
        Application.DisplayAlerts = False ' ���b�Z�[�W���\��
    
        '�R�s�[����csv�t�@�C���̃V�[�g�폜'
        ThisWorkbook.Sheets(sheetNameCSV).Delete
        
        '�f�[�^���'
        Analysis.dataAnalysis
        
        '�Z�b�g�����f�[�^�N���A'
        Clear.dataClear
        Clear.retClear
        
        '����csv�t�@�C���������t�@�C�����擾'
        fileName = Dir()
    Loop
       
        Application.Calculation = xlAutomatic

'    If Not msg = "" Then
'        msg = buf + "��ǂݍ��߂܂���ł����B"
'    Else
'        msg = buf
'    End If
'
'    MsgBox msg
    MsgBox "�������܂����B"
    
    Worksheets(constCopySheetName).Activate ' �uSheet1�v�̃V�[�g���A�N�e�B�u
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

'
'�f�[�^�Z�b�g
'
Public Function dataSet(ByVal sheetNameCSV As String) As Boolean
    Dim cnt_csv_line As Long
    Dim cnt_csv_row_kokyu As Long
    Dim cnt_csv_row_acce As Long
    Dim cnt_dst_line As Long
    
    Set sh_dst = Sheets("�f�[�^")
    
    'csv�t�@�C���̃f�[�^�̊J�n�ʒu'
    cnt_csv_line = 4
    cnt_csv_row_kokyu = 4
    
    '�f�[�^���Z�b�g����J�n�s'
    cnt_dst_line = 2
    
    While IsEmpty(Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu)) = False
        '�f�[�^�Z�b�g'
        If cnt_csv_row_kokyu <= 6 Then
            '���т��A���ċz����Z�b�g'
            If Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 0 Then
                sh_dst.Range(sh_dst.Cells(cnt_dst_line, "E"), sh_dst.Cells(cnt_dst_line, "F")).Value = 0
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 1 Then
                sh_dst.Cells(cnt_dst_line, "E").Value = 0
                sh_dst.Cells(cnt_dst_line, "F").Value = 1
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_kokyu) = 2 Then
                sh_dst.Cells(cnt_dst_line, "E").Value = 2
                sh_dst.Cells(cnt_dst_line, "F").Value = 0
            End If
            
            cnt_csv_row_acce = cnt_csv_row_kokyu + 7
            '��̌����Z�b�g'
            If Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 0 Then
                '��'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 6).Value = 1
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 1 Then
                '��'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow).Value = 7
            ElseIf Worksheets(sheetNameCSV).Cells(cnt_csv_line, cnt_csv_row_acce) = 2 Then
                '�E'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 2).Value = 5
            Else
                '��'
                sh_dst.Cells(cnt_dst_line, constRetAcceStartRow + 4).Value = 3
            End If
        End If
        
        If cnt_csv_row_kokyu < 6 Then
            '�ċz��Ԃ̎��̒l��'
            cnt_csv_row_kokyu = cnt_csv_row_kokyu + 1
        Else
            '���̍s�̌ċz��Ԃ̍ŏ��̒l��'
            cnt_csv_row_kokyu = 4
            cnt_csv_line = cnt_csv_line + 1
        End If
        
        '�f�[�^���Z�b�g����s������'
        cnt_dst_line = cnt_dst_line + 1
    Wend
End Function

'
'�J�n���ԃZ�b�g
'
Public Function setStartTime(ByVal sheetNameCSV As String) As String
    Dim year As String
    Dim time As Date
    
    
    year = Worksheets(sheetNameCSV).Range("A3").Value
    time = Worksheets(sheetNameCSV).Range("C3").Value
    
    setStartTime = year + " " + time
    
End Function


