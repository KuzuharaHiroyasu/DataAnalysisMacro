Attribute VB_Name = "Read"
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    'åƒãzâπ'
    If Not readText(ThisWorkbook.Path & "\raw_sum.txt", 2) Then
        msg = "raw_sum.txt "
    End If
    
    'Ç¢Ç—Ç´âπ'
    If Not readText(ThisWorkbook.Path & "\rawsnore_sum.txt", 3) Then
        msg = msg + "rawsnore_sum.txt "
    End If
    
    'Ç¢Ç—Ç´èÛë‘'
    If Not readText(ThisWorkbook.Path & "\constSnore__sum.txt", 5) Then
        msg = msg + "constSnore__sum.txt "
    End If
    
    'ñ≥åƒãzèÛë‘'
    If Not readText(ThisWorkbook.Path & "\constApnea_sum.txt", 6) Then
        msg = msg + "constApnea_sum.txt "
    End If
    
    'Xé≤'
    If Not readText(ThisWorkbook.Path & "\acce_x_sum.txt", 7) Then
        msg = msg + "acce_x_sum.txt "
    End If
    
    'Yé≤'
    If Not readText(ThisWorkbook.Path & "\acce_y_sum.txt", 8) Then
        msg = msg + "acce_y_sum.txt "
    End If
    
    'Zé≤'
    If Not readText(ThisWorkbook.Path & "\acce_z_sum.txt", 9) Then
        msg = msg + "acce_z_sum.txt "
    End If
    
    If Not msg = "" Then
        msg = msg + "Çì«Ç›çûÇﬂÇ‹ÇπÇÒÇ≈ÇµÇΩÅB"
    Else
        msg = "äÆóπÇµÇ‹ÇµÇΩÅB"
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
