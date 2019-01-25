Attribute VB_Name = "Module1"
'定数'
'データシート'
Const dataSheetName = "データ"  'シート名'
'行'
Const initDataLine = 2          'データの最初の行'
'列'
Const noRow = 1                 'Noの列(A列)'
Const rawRow = 2                '呼吸音の列(B列)'
Const rawSnoreRow = 3           'いびき音の列(C列)'
Const rawMovAvrRow = 4          '呼吸音の移動平均の列(D列)'
Const snoreStateRow = 5         'いびき判定結果の入った列(E列)'
Const apneaStateRow = 6         '無呼吸判定結果の入った列(F列)'
Const acceXRow = 7              '加速度(X)の入った列(G列)'
Const acceYRow = 8              '加速度(Y)の入った列(H列)'
Const acceZRow = 9              '加速度(Z)の入った列(I列)'
Const retAcceRow = 10           '向き(J列)'


'結果シート'
Const retSheetName = "結果"     'シート名'
'行'
Const initRetLine = 7           '結果が入力される最初の行'
'列'
Const retStartTimeRow = 2       '判定時刻の列(B列)'
Const retStopTimeRow = 3        '停止時刻の列(C列)'
Const retContinuTimeRow = 4     '継続時間の列(D列)'
Const retTypeRow = 5            '種別の列(E列)'
Const retStartFromStopTimeRow = 6   '前回停止から今回発生までの時間の列(F列)'
Const retRemarkRow = 7          '備考の列(F列)'
'入力文字'
Const snore = "いびき"
Const apnea = "無呼吸"

Sub dataAnalysis()
    Dim snoreState As Integer           'いびき判定状態'
    Dim apneaState As Integer           '無呼吸判定状態'
    Dim beforeSnoreState As Integer     '１つ前のいびき判定状態'
    Dim beforeApneaState As Integer     '１つ前の無呼吸判定状態'
    Dim time As Integer                 '経過時間(秒)'
    Dim snoreCnt As Integer             'いびき回数'
    Dim apneaCnt As Integer             '無呼吸回数'
    Dim startTime As Date               '開始時刻'
    Dim dataLine As Long                '現在解析中のデータの行'
    Dim retLine As Long                 '現在結果入力中の行'
    Dim no As Long                      'ナンバー'
    Dim remark As Long                  '備考用'
    
    '初期化'
    snoreState = 0
    apneaState = 0
    time = 0
    snoreCnt = 0
    apneaCnt = 0
    no = 1
    dataLine = initDataLine
    retLine = initRetLine
    
    '開始時刻設定'
    startTime = Sheets(retSheetName).Range("B3").Value
    
    '解析'
    While IsEmpty(Sheets(dataSheetName).Cells(dataLine, snoreStateRow)) = False
        Sheets(dataSheetName).Cells(dataLine, noRow).Value = no 'No挿入'
        beforeSnoreState = snoreState                       '１つ前のいびき判定状態を保存'
        beforeApneaState = apneaState                       '１つ前の無呼吸判定状態を保存'
        snoreState = Sheets(dataSheetName).Cells(dataLine, snoreStateRow).Value   'いびき状態取得'
        apneaState = Sheets(dataSheetName).Cells(dataLine, apneaStateRow).Value   '無呼吸状態取得'
        
        '呼吸の移動平均'
        If no >= 5 Then
            Sheets(dataSheetName).Cells(dataLine, rawMovAvrRow).Value = WorksheetFunction.Sum(Range(Sheets(dataSheetName).Cells(dataLine - 4, rawRow), Sheets(dataSheetName).Cells(dataLine, rawRow))) / 5
        Else
            Sheets(dataSheetName).Cells(dataLine, rawMovAvrRow).Value = "-"
        End If
        
        
        If snoreState = 1 Then
        'いびき判定あり'
            If beforeApneaState = 1 Or beforeApneaState = 2 Then
            '１つ前で無呼吸判定ありだった'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "から" & no
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
        
            If beforeSnoreState = 0 Then
            '１つ前はいびき判定なし'
                Call setStart(retLine, startTime, time, snore)
                snoreCnt = snoreCnt + 1
                remark = no
            End If
        ElseIf apneaState = 1 Or apneaState = 2 Then
        '無呼吸判定あり'
            If beforeSnoreState = 1 Then
            '１つ前でいびき判定ありだった'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "から" & no
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
        
            If beforeApneaState = 0 Then
            '１つ前は無呼吸判定なし'
                Call setStart(retLine, startTime, time, apnea)
                apneaCnt = apneaCnt + 1
                remark = no
            End If
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '１つ前で無呼吸判定あり、もしくはいびき判定ありだった'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "から" & no
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
        End If
        
        no = no + 1
        time = time + 10    '時間を10秒増やす'
        dataLine = dataLine + 1     '次の行のデータ02へ'
    Wend
    
    If IsEmpty(Sheets(retSheetName).Cells(retLine, retTypeRow).Value) = False Then
        '最後の判定の停止時刻など'
        Call setEnd(retLine, startTime, time)
        Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "から" & no
    End If
    
    '終了時刻'
    Sheets(retSheetName).Range("C3").Value = DateAdd("s", time, startTime)
    
    'データ取得時間'
    Sheets(retSheetName).Range("D3").Value = CStr(CDate(DateDiff("s", startTime, Sheets(retSheetName).Range("C3").Value) / 86400#))
    
    'いびき回数'
    Sheets(retSheetName).Range("E3").Value = snoreCnt
    
    '無呼吸回数'
    Sheets(retSheetName).Range("F3").Value = apneaCnt
    
    'グラフ削除(一度削除する)'
    If Sheets(retSheetName).ChartObjects.Count > 0 Then
        Sheets(retSheetName).ChartObjects.Delete
    End If
    
    'グラフ作成'
    'いびき/呼吸の大きさ'
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, rawRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, rawSnoreRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, rawRow), Sheets(dataSheetName).Cells(Rows.Count, rawSnoreRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H7").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H7").Left
            .SeriesCollection(1).Name = "=""いびき"""
            .SeriesCollection(2).Name = "=""呼吸音"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 1024
            .Axes(xlValue).MajorUnit = 256
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
        End With
    End If
    
    'いびき/呼吸の判定'
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, snoreStateRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, apneaStateRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, snoreStateRow), Sheets(dataSheetName).Cells(Rows.Count, apneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H19").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H19").Left
            .SeriesCollection(1).Name = "=""いびき"""
            .SeriesCollection(2).Name = "=""無呼吸"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 2
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
        End With
    End If
    
    '加速度センサー'
    '向き'
    Call acceAnalysis
    
    Dim endLine As Long               '向き判定の最終行'
    Dim i As Long
    i = 1
    endLine = Sheets(dataSheetName).Cells(Rows.Count, retAcceRow).End(xlUp).Row
    
    '最終の向きの行数検索'
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
            .SeriesCollection(1).Name = "=""上"""
            .SeriesCollection(2).Name = "=""右上"""
            .SeriesCollection(3).Name = "=""右"""
            .SeriesCollection(4).Name = "=""右下"""
            .SeriesCollection(5).Name = "=""下"""
            .SeriesCollection(6).Name = "=""左下"""
            .SeriesCollection(7).Name = "=""左"""
            .SeriesCollection(8).Name = "=""左上"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 7
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
        End With
    End If
    
    'センサー値'
    If IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceXRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceYRow)) = False And IsEmpty(Sheets(dataSheetName).Cells(initDataLine, acceZRow)) = False Then
        With Sheets(retSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(dataSheetName).Range(Sheets(dataSheetName).Cells(initDataLine, acceXRow), Sheets(dataSheetName).Cells(Rows.Count, acceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(retSheetName).Range("H43").Top
            .ChartArea.Left = Sheets(retSheetName).Range("H43").Left
            .SeriesCollection(1).Name = "=""Ｘ軸"""
            .SeriesCollection(2).Name = "=""Ｙ軸"""
            .SeriesCollection(3).Name = "=""Ｚ軸"""
            .Axes(xlValue).MinimumScale = -100
            .Axes(xlValue).MaximumScale = 100
            .Axes(xlValue).MajorUnit = 50
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
        End With
    End If
    
    MsgBox "完了しました。"
End Sub

Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
    Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value = DateAdd("s", time, startTime)   '開始時刻セット'
    Sheets(retSheetName).Cells(retLine, retStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    Sheets(retSheetName).Cells(retLine, retTypeRow).Value = kind                                '種別セット'
End Sub

Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Sheets(retSheetName).Cells(retLine, retStopTimeRow).Value = DateAdd("s", time, startTime)   '停止時刻セット'
    Sheets(retSheetName).Cells(retLine, retStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    Sheets(retSheetName).Cells(retLine, retContinuTimeRow).Value = Sheets(retSheetName).Cells(retLine, retStopTimeRow).Value - Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value   '継続時間'
    Sheets(retSheetName).Cells(retLine, retContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '継続時間書式設定'
    If retLine = initRetLine Then
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).Value = "-"
    Else
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).Value = Sheets(retSheetName).Cells(retLine, retStartTimeRow).Value - Sheets(retSheetName).Cells(retLine - 1, retStopTimeRow).Value '前回停止から今回発生までの時間'
        Sheets(retSheetName).Cells(retLine, retStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '前回停止から今回発生までの時間書式設定'
    End If
End Sub

Sub dataAndResultClear()
    Dim cnt As Long
    
    'グラフ削除'
    If Sheets(retSheetName).ChartObjects.Count > 0 Then
        Sheets(retSheetName).ChartObjects.Delete
    End If
    
    'データシート'
    cnt = 2 '2行目から'
    While IsEmpty(Sheets(dataSheetName).Cells(cnt, rawRow)) = False
        Sheets(dataSheetName).Rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    '結果シート'
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
    
    MsgBox "削除完了しました。"
End Sub

Sub acceAnalysis()
    Dim x As Integer                    '加速度センサー_X軸'
    Dim y As Integer                    '加速度センサー_Y軸'
    Dim z As Integer                    '加速度センサー_Z軸'
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
        
        'ヘッド部の場所（体の向きではない）'
        If 0 <= x Then
            '右側'
            If 0 <= z Then
                '上側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '右上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    Else
                        '上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '右上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow).Value = 7
                    Else
                        '右(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '下側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '右下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    Else
                        '下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '右下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    Else
                        '右(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 2).Value = 5
                    End If
                End If
            End If
        Else
            '左側'
            If 0 <= z Then
                '上側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '左上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    Else
                        '上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '左上(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    Else
                        '左(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 6).Value = 1
                    End If
                End If
            Else
                '下側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '左下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    Else
                        '下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '左下(確)'
                        Sheets(dataSheetName).Cells(line, retAcceRow + 4).Value = 3
                    Else
                        '左(確)'
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
