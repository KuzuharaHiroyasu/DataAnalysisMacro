Attribute VB_Name = "Clear"
Sub Clear()
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

    cnt = 44    '44行目から'
    While IsEmpty(Sheets(retSheetName).Cells(cnt, retStartTimeRow)) = False
        Sheets(retSheetName).Rows(cnt).Clear
        cnt = cnt + 1
    Wend
    
    MsgBox "削除完了しました。"
End Sub

