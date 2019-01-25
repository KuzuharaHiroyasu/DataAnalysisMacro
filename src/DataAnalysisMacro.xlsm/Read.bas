Attribute VB_Name = "Read"
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    '呼吸音'
    If Not readText(ThisWorkbook.Path & "\raw_sum.txt", 2) Then
        msg = "raw_sum.txt "
    End If
    
    'いびき音'
    If Not readText(ThisWorkbook.Path & "\rawsnore_sum.txt", 3) Then
        msg = msg + "rawsnore_sum.txt "
    End If
    
    'いびき状態'
    If Not readText(ThisWorkbook.Path & "\snore__sum.txt", 5) Then
        msg = msg + "snore__sum.txt "
    End If
    
    '無呼吸状態'
    If Not readText(ThisWorkbook.Path & "\apnea_sum.txt", 6) Then
        msg = msg + "apnea_sum.txt "
    End If
    
    'X軸'
    If Not readText(ThisWorkbook.Path & "\acce_x_sum.txt", 7) Then
        msg = msg + "acce_x_sum.txt "
    End If
    
    'Y軸'
    If Not readText(ThisWorkbook.Path & "\acce_y_sum.txt", 8) Then
        msg = msg + "acce_y_sum.txt "
    End If
    
    'Z軸'
    If Not readText(ThisWorkbook.Path & "\acce_z_sum.txt", 9) Then
        msg = msg + "acce_z_sum.txt "
    End If
    
    If Not msg = "" Then
        msg = msg + "を読み込めませんでした。"
    Else
        msg = "完了しました。"
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
            Sheets(dataSheetName).Cells(inputLine, inputRow) = buf
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
