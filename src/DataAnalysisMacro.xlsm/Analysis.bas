Attribute VB_Name = "Analysis"


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
    no = 0
    dataLine = initDataLine
    retLine = initRetLine
    
    '開始時刻設定'
    startTime = Sheets(retSheetName).Range("B3").Value
    
    '解析'
    While IsEmpty(Sheets(dataSheetName).Cells(dataLine, snoreStateRow)) = False
        Sheets(dataSheetName).Cells(dataLine, noRow).Value = no 'ナンバー挿入'
        beforeSnoreState = snoreState                       '１つ前のいびき判定状態を保存'
        beforeApneaState = apneaState                       '１つ前の無呼吸判定状態を保存'
        snoreState = Sheets(dataSheetName).Cells(dataLine, snoreStateRow).Value   'いびき状態取得'
        apneaState = Sheets(dataSheetName).Cells(dataLine, apneaStateRow).Value   '無呼吸状態取得'
        
        '呼吸の移動平均'
        If no >= 4 Then
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
            
            'いびきのトータル時間'
            
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
            
            '無呼吸のトータル時間'
            
            
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '１つ前で無呼吸判定あり、もしくはいびき判定ありだった'
                Call setEnd(retLine, startTime, time)
                Sheets(retSheetName).Cells(retLine, retRemarkRow).Value = remark & "から" & no
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
            
            '通常呼吸のトータル時間'
            
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
    
    '状態別の体の向きの時間'
    
    'no×20+20' 'いびき1の行を探す⇒そこのno×20+20行目の向きを求めて加算、最後に該当の箇所に代入'
    
    
    
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


