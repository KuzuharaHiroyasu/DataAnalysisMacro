Attribute VB_Name = "Analysis"
'
'データ解析
'
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
    
    ''''''初期化''''''
    snoreState = 0
    apneaState = 0
    time = 0
    snoreCnt = 0
    apneaCnt = 0
    no = 0
    dataLine = constInitDataLine
    retLine = constInitRetLine
    
    ''''''開始時刻設定''''''
    startTime = Sheets(constRetSheetName).Range("B3").Value
    
    ''''''解析''''''
    While IsEmpty(Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow)) = False
        Sheets(constDataSheetName).Cells(dataLine, constNoRow).Value = no 'ナンバー挿入'
        beforeSnoreState = snoreState                       '１つ前のいびき判定状態を保存'
        beforeApneaState = apneaState                       '１つ前の無呼吸判定状態を保存'
        snoreState = Sheets(constDataSheetName).Cells(dataLine, constSnoreStateRow).Value   'いびき状態取得'
        apneaState = Sheets(constDataSheetName).Cells(dataLine, constApneaStateRow).Value   '無呼吸状態取得'
        
        '呼吸の移動平均'
        Call movAverage(dataLine)
        
        If snoreState = 1 Then
        'いびき判定あり'
            If beforeApneaState = 1 Or beforeApneaState = 2 Then
            '１つ前で無呼吸判定ありだった'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
        
            If beforeSnoreState = 0 Then
            '１つ前はいびき判定なし'
                Call setStart(retLine, startTime, time, constSnore)
                snoreCnt = snoreCnt + 1
                remark = no
            End If
            
            'いびきのトータル時間'
            
        ElseIf apneaState = 1 Or apneaState = 2 Then
        '無呼吸判定あり'
            If beforeSnoreState = 1 Then
            '１つ前でいびき判定ありだった'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
        
            If beforeApneaState = 0 Then
            '１つ前は無呼吸判定なし'
                Call setStart(retLine, startTime, time, constApnea)
                apneaCnt = apneaCnt + 1
                remark = no
            End If
            
            '無呼吸のトータル時間'
            
            
        Else
            If beforeApneaState = 1 Or beforeApneaState = 2 Or beforeSnoreState = 1 Then
            '１つ前で無呼吸判定あり、もしくはいびき判定ありだった'
                Call setRemarks(retLine, startTime, time, remark, no)
                retLine = retLine + 1     '結果入力を次の行へ'
            End If
            
            '通常呼吸のトータル時間'
            
        End If
        
        no = no + 1
        time = time + 10    '時間を10秒増やす'
        dataLine = dataLine + 1     '次の行のデータ02へ'
    Wend
    
    If IsEmpty(Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value) = False Then
        '最後の判定の停止時刻など'
        Call setRemarks(retLine, startTime, time, remark, no)
    End If
    
    
    ''''''各種データ記入''''''
    Call setData(time, startTime, snoreCnt, apneaCnt)
    
    '状態別の体の向きの時間'
    
    'no×20+20' 'いびき1の行を探す⇒そこのno×20+20行目の向きを求めて加算、最後に該当の箇所に代入'
    
    
    ''''''加速度センサー''''''
    '向き'
    Call acceAnalysis
    
    Dim endLine As Long
    Dim i As Long
    i = 1
    
    '向き判定の最終行'
    endLine = Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow).End(xlUp).Row
    
    '最終の向きの行数検索'
    While i <= 7
        If endLine <= Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow + i).End(xlUp).Row Then
            endLine = Sheets(constDataSheetName).Cells(Rows.Count, constRetAcceRow + i).End(xlUp).Row
        End If
        i = i + 1
    Wend
    
    ''''''グラフ作成''''''
    '既にグラフがあれば一旦削除'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    Call createGraph(endLine)
    
    MsgBox "完了しました。"
End Sub

'
'いびき・無呼吸の開始時刻セット
'
Sub setStart(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal kind As String)
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value = DateAdd("s", time, startTime)   '開始時刻セット'
    Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    Sheets(constRetSheetName).Cells(retLine, constRetTypeRow).Value = kind                                '種別セット'
End Sub

'
'いびき・無呼吸の終了時刻セット
'
Sub setEnd(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long)
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value = DateAdd("s", time, startTime)   '停止時刻セット'
    Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).NumberFormatLocal = "hh:mm:ss"         '時刻書式設定'
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStopTimeRow).Value - Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value   '継続時間'
    Sheets(constRetSheetName).Cells(retLine, constRetContinuTimeRow).NumberFormatLocal = "hh:mm:ss"      '継続時間書式設定'
    If retLine = constInitRetLine Then
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = "-"
    Else
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).Value = Sheets(constRetSheetName).Cells(retLine, constRetStartTimeRow).Value - Sheets(constRetSheetName).Cells(retLine - 1, constRetStopTimeRow).Value '前回停止から今回発生までの時間'
        Sheets(constRetSheetName).Cells(retLine, constRetStartFromStopTimeRow).NumberFormatLocal = "hh:mm:ss" '前回停止から今回発生までの時間書式設定'
    End If
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
                        '右上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '右上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow).Value = 7
                    Else
                        '右(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 1).Value = 6
                    End If
                End If
            Else
                '下側'
                If x_abs < z_abs Then
                    If (z_abs - x_abs) < 10 Then
                        '右下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 2).Value = 5
                    Else
                        '下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 3).Value = 4
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '右下(確)'
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
                        '左上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 6).Value = 1
                    Else
                        '上(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 7).Value = 0
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '左上(確)'
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
                        '左下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    End If
                Else
                    If (x_abs - z_abs) < 10 Then
                        '左下(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 4).Value = 3
                    Else
                        '左(確)'
                        Sheets(constDataSheetName).Cells(line, constRetAcceRow + 5).Value = 2
                    End If
                End If
            End If
        End If
        line = line + 1
    Wend
End Sub

'
'移動平均を求める
'
Sub movAverage(ByVal dataLine As Long)
    If no >= 4 Then
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = WorksheetFunction.Sum(Range(Sheets(constDataSheetName).Cells(dataLine - 4, constRawRow), Sheets(constDataSheetName).Cells(dataLine, constRawRow))) / 5
    Else
        Sheets(constDataSheetName).Cells(dataLine, constRawMovAvrRow).Value = "-"
    End If
End Sub

'
'備考欄に記入
'
Sub setRemarks(ByVal retLine As Long, ByVal startTime As Date, ByVal time As Long, ByVal remark As Long, ByVal no As Long)
    Call setEnd(retLine, startTime, time)
    Sheets(constRetSheetName).Cells(retLine, constRetRemarkRow).Value = remark & "から" & no
End Sub

'
'各種データをセットする
'
Sub setData(ByVal time As Long, ByVal startTime As Date, ByVal snoreCnt As Integer, ByVal apneaCnt As Integer)
    '終了時刻'
    Sheets(constRetSheetName).Range("C3").Value = DateAdd("s", time, startTime)
    
    'データ取得時間'
    Sheets(constRetSheetName).Range("D3").Value = CStr(CDate(DateDiff("s", startTime, Sheets(constRetSheetName).Range("C3").Value) / 86400#))
    
    'いびき回数'
    Sheets(constRetSheetName).Range("E3").Value = snoreCnt
    
    '無呼吸回数'
    Sheets(constRetSheetName).Range("F3").Value = apneaCnt
End Sub

'
'グラフ作成
'
Sub createGraph(ByVal endLine As Long)
'いびき/呼吸の大きさ'
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constRawSnoreRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constRawRow), Sheets(constDataSheetName).Cells(Rows.Count, constRawSnoreRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H7").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H7").Left
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
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constApneaStateRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constSnoreStateRow), Sheets(constDataSheetName).Cells(Rows.Count, constApneaStateRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H19").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H19").Left
            .SeriesCollection(1).Name = "=""いびき"""
            .SeriesCollection(2).Name = "=""無呼吸"""
            .Axes(xlValue).MinimumScale = 0
            .Axes(xlValue).MaximumScale = 2
            .Axes(xlValue).MajorUnit = 1
            .Axes(xlCategory).HasMajorGridlines = False
            .Axes(xlCategory).TickLabels.NumberFormatLocal = "G/標準"
        End With
    End If
    
    '体の向き'
    If endLine > 1 Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine - 1, constRetAcceRow), Sheets(constDataSheetName).Cells(endLine, 17))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H31").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H31").Left
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
    If IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceYRow)) = False And IsEmpty(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceZRow)) = False Then
        With Sheets(constRetSheetName).ChartObjects.Add(30, 50, 300, 200).Chart
            .ChartType = xlLine
            .SetSourceData Source:=Sheets(constDataSheetName).Range(Sheets(constDataSheetName).Cells(constInitDataLine, constAcceXRow), Sheets(constDataSheetName).Cells(Rows.Count, constAcceZRow).End(xlUp))
            .ChartArea.Width = 36000
            .ChartArea.Height = 150
            .ChartArea.Top = Sheets(constRetSheetName).Range("H43").Top
            .ChartArea.Left = Sheets(constRetSheetName).Range("H43").Left
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
End Sub























