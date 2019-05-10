# VBA
VBA学习笔记


# 技巧

1. 求X列的行数
    
        irow = Cells(Rows.Count, X).End(xlUp).Row 

2. 字母转数字

        Function CWtoN(ByVal AB As String) As Long
            CWtoN = Range("a1:" & AB & "1").Cells.Count
        End Function

3. 数字转字母

        Function CNtoW(ByVal num As Long) As String
            CNtoW = Replace(Cells(1, num).Address(False, False), "1", "")
        End Function
4. 更新图表的数据

        Function updateChart(oChart, Data)
        With oChart
            Dim gWorkSheet As Excel.Worksheet

                
            Set gWorkSheet = .ChartData.Workbook.Worksheets("Sheet1")
            gWorkSheet.Range("a1:" & CNtoW(UBound(Data, 2)) & UBound(Data, 1)).Resize(32) = Data
            
            .refresh
            .ChartData.Activate
            .ChartData.Workbook.Close
        End With
        End Function

        