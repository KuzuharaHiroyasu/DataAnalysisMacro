Attribute VB_Name = "Read"
'
'データ読み込み
'
Sub readData()
    Dim ret As Boolean
    Dim msg As String
    
    Application.Calculation = xlManual
    
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
    Dim x As Integer                    '加速度センサー_X軸'
    Dim y As Integer                    '加速度センサー_Y軸'
    Dim z As Integer                    '加速度センサー_Z軸'
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
        
        'ヘッド部の場所（体の向きではない）'
        If 0 <= x Then
            '右側'
            If 0 <= z Then
                '上側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '右上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '下側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '右(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '右下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '右(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '右(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    End If
                End If
            End If
        Else
            '左側'
            If 0 <= z Then
                '上側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '左(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '左上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '左(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '左(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    End If
                End If
            Else
                '下側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '左下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 5).Value = 2
                    End If
                End If
            End If
        End If
        line = line + 1
    Wend
End Sub

