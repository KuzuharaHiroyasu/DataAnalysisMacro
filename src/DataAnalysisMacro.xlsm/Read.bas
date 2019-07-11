Attribute VB_Name = "Read"
'
'データ読み込み
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    Application.Calculation = xlManual
    
    '呼吸音'
    If Not readText(ThisWorkbook.Path & "\raw_sum.txt", constRawRow) Then
        msg = "raw_sum.txt "
    End If
    
    '心拍除去後の呼吸音'
    If Not readText(ThisWorkbook.Path & "\raw_heartBeatRemov_sum.txt", constRawHBRemovRow) Then
        msg = "raw_heartBeatRemov_sum.txt "
    End If
    
    'いびき音'
    If Not readText(ThisWorkbook.Path & "\rawsnore_sum.txt", constRawSnoreRow) Then
        msg = msg + "rawsnore_sum.txt "
    End If
    
    '無呼吸状態'
    If Not readText(ThisWorkbook.Path & "\apnea_sum.txt", constApneaStateRow) Then
        msg = msg + "apnea_sum.txt "
    End If
    
    'いびき状態'
    If Not readText(ThisWorkbook.Path & "\snore__sum.txt", constSnoreStateRow) Then
        msg = msg + "snore__sum.txt "
    End If
    
    'フォトセンサー値'
    If Not readText(ThisWorkbook.Path & "\photoref_sum.txt", constPhotorefRow) Then
        msg = msg + "photoref_sum.txt "
    End If
    
    'X軸'
    If Not readText(ThisWorkbook.Path & "\acce_x_sum.txt", constAcceXRow) Then
        msg = msg + "acce_x_sum.txt "
    End If
    
    'Y軸'
    If Not readText(ThisWorkbook.Path & "\acce_y_sum.txt", constAcceYRow) Then
        msg = msg + "acce_y_sum.txt "
    End If
    
    'Z軸'
    If Not readText(ThisWorkbook.Path & "\acce_z_sum.txt", constAcceZRow) Then
        msg = msg + "acce_z_sum.txt "
    End If
    
    Call acceAnalysis
    
    Application.Calculation = xlAutomatic

    If Not msg = "" Then
        msg = msg + "を読み込めませんでした。"
    Else
        msg = "完了しました。"
    End If
    
    MsgBox msg
End Sub

'
'データテキスト読み込み
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
'加速度センサー値の絶対値
'
Sub absoluteValue(ByVal x As Integer, ByVal y As Integer, ByVal z As Integer)
    x = Abs(x)
    y = Abs(y)
    z = Abs(z)
End Sub

'
'体の向きを決める
'
Sub acceAnalysis()
    Dim x As Long                    '加速度センサー_X軸'
    Dim y As Long                    '加速度センサー_Y軸'
    Dim z As Long                    '加速度センサー_Z軸'
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
            'エラー回避(イレギュラーな値)'
            Sheets(constDataSheetName).Range(Cells(line, constAcceXRow), Cells(line, constAcceZRow)).Delete
            GoTo Continue
        End If
        
        x_abs = Abs(x)
        z_abs = Abs(z)
        
        'ヘッド部の場所（体の向きではない）'
        If 0 <= x Then
            '上or左'
            If 0 <= z Then
                '左'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 6).Value = 1
            Else
                '上'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow).Value = 7
            End If
        Else
            '下or右'
            If 0 <= z Then
                '下'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 4).Value = 3
            Else
                '右'
                Sheets(constDataSheetName).Cells(line, constRetAcceStartRow + 2).Value = 5
            End If
        End If
        line = line + 1
Continue:
    Wend
End Sub

