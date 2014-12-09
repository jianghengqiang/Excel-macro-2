Public Sub 汇总()
Dim i As Integer  '复合文件个数
Dim j As Integer '单独文件的个数
Dim m As Integer '定义行偏移量
Dim k As Integer '定义列偏移量

Dim num As Integer
Dim wb As Workbook
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim Test1 As Boolean
Dim Test2 As Boolean

Set wb = ThisWorkbook
num = wb.Worksheets.Count

Set ws2 = wb.Worksheets(num)
m = 0
For i = 1 To num - 3 Step 2
   j = 0
   Test1 = True
   
   While (Test1)
   k = 0
   Test2 = True
   
   While (Test2)
      ws2.Range("A3").Offset(m, k).Value = wb.Sheets(i).Range("A3").Offset(j, k).Value
      k = k + 1
      Test2 = wb.Sheets(i).Range("A3").Offset(j, k).Value <> ""
   Wend
   m = m + 1
   j = j + 1
   Test1 = wb.Sheets(i).Range("A3").Offset(j, 0).Value <> ""
   Wend
   
Next i

End Sub
