Attribute VB_Name = "Clear"
Sub Clear()
    Dim cnt As Long
    
    'グラフ削除'
    If Sheets(constRetSheetName).ChartObjects.Count > 0 Then
        Sheets(constRetSheetName).ChartObjects.Delete
    End If
    
    'データシート'
    cnt = 2 '2行目から'
    While IsEmpty(Sheets(constDataSheetName).Cells(cnt, constRawRow)) = False
        Sheets(constDataSheetName).rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    '結果シート'
    Sheets(constRetSheetName).rows(3).Clear
'    Sheets(constRetSheetName).Range("C3").Clear
'    Sheets(constRetSheetName).Range("D3").Clear
'    Sheets(constRetSheetName).Range("E3").Clear
'    Sheets(constRetSheetName).Range("F3").Clear

    cnt = 44    '44行目から'
    While IsEmpty(Sheets(constRetSheetName).Cells(cnt, constRetStartTimeRow)) = False
        Sheets(constRetSheetName).rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    MsgBox "削除完了しました。"
End Sub

