Attribute VB_Name = "Read"
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
    If Not readText(ThisWorkbook.Path & "\constSnore__sum.txt", 5) Then
        msg = msg + "constSnore__sum.txt "
    End If
    
    '���ċz���'
    If Not readText(ThisWorkbook.Path & "\constApnea_sum.txt", 6) Then
        msg = msg + "constApnea_sum.txt "
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


Private Function readText(ByVal fileName As String, ByVal inputRow As Long) As Boolean
    Dim a
    Dim inputLine As Long
    
    inputLine = 2
    
    a = Dir(fileName)
    If (a <> "") Then
        Open fileName For Input As #1
        Do Until EOF(1)
            Line Input #1, buf
            Sheets(constDataSheetName).Cells(inputLine, inputRow) = buf
            inputLine = inputLine + 1
        Loop
        Close #1
        readText = True
    Else
        readText = False
    End If
End Function

Sub absoluteValue(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    x = Abs(x)
    y = Abs(y)
    z = Abs(z)
End Sub
