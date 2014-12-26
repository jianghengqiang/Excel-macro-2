Public Sub 汇总()
Dim i As Integer  '文件个数
Dim j As Integer
Dim m As Integer  '定义行偏移量变量
Dim k As Integer  '定义列偏移量变量
Dim Bookname As String
Dim Test1 As Boolean
Dim Test2 As Boolean
m = 0
For i = 2 To 6 Step 1
    j = 0
    Bookname = "0" & CStr(i) & ".xls"
    Workbooks.Open Bookname
    Test1 = True
    While (Test1)
       k = 0
       Test2 = True
       Debug.Print m
       While (Test2)
          Workbooks("All2.xlsm").Sheets(n).Range("A4").Offset(m, k).Value = _
          Workbooks(Bookname).Sheets(n).Range("A4").Offset(j, k).Value
                                          
          k = k + 1
          Test2 = Workbooks(Bookname).Sheets(n).Range("A4").Offset(j, k).Value <> ""
       Wend
       m = m + 1
       j = j + 1
       Test1 = Workbooks(Bookname).Sheets(n).Range("A4").Offset(j, 0).Value <> ""
    Wend
    Debug.Print m
    Workbooks(Bookname).Close
Next i

End Sub
