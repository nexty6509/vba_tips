Attribute VB_Name = "Module1"
Sub ボタン1_Click()


    '指定列を削除
    cnt = Cells(Rows.Count, 1).End(xlUp).Row

'    Range("A2", "A3").Clear
    
'    Range(Cells(2, 1), Cells(3, 1)).Clear
    
'    Range(Cells(4, 1), Cells(6, 1)).Clear
    
    
    '2列目から最終列まで行削除
    Range(Cells(2, 1), Cells(cnt, 1)).Clear


    '2列目
    Range(Cells(2, 2), Cells(cnt, 2)).Clear


    '3列目
    cnt = Cells(Rows.Count, 3).End(xlUp).Row
    Range(Cells(2, 3), Cells(cnt, 3)).Clear


End Sub
