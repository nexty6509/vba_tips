Attribute VB_Name = "Module1"
Sub �{�^��1_Click()


    '�w�����폜
    cnt = Cells(Rows.Count, 1).End(xlUp).Row

'    Range("A2", "A3").Clear
    
'    Range(Cells(2, 1), Cells(3, 1)).Clear
    
'    Range(Cells(4, 1), Cells(6, 1)).Clear
    
    
    '2��ڂ���ŏI��܂ōs�폜
    Range(Cells(2, 1), Cells(cnt, 1)).Clear


    '2���
    Range(Cells(2, 2), Cells(cnt, 2)).Clear


    '3���
    cnt = Cells(Rows.Count, 3).End(xlUp).Row
    Range(Cells(2, 3), Cells(cnt, 3)).Clear


End Sub
