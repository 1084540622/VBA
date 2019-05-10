# VBA对象

    VBA中的对象其实就是我们操作的具有方法、属性的excel中支持的对象

### Excel中的几个常用对象表示方法

## 1. 工作簿
 
    Workbooks 代表工作簿集合，所有的工作簿,Workbooks(N)，表示已打开的第N个工作簿
    Workbooks ("工作簿名称")
    ActiveWorkbook 正在操作的工作簿
    ThisWorkBook '代码所在的工作簿
      
## 2. 工作表
    Sheets("工作表名称")
    Sheet1 表示第一个插入的工作表,Sheet2表示第二个插入的工作表....
    Sheets(n) 表示按排列顺序，第n个工作表
    ActiveSheet 表示活动工作表，光标所在工作表
    worksheet 也表示工作表，但不包括图表工作表、宏工作表等。

## 3. 单元格
    cells 所有单元格
    Range ("单元格地址")
    Cells(行数,列数)
    Activecell 正在选中或编辑的单元格
    Selection 正被选中或选取的单元格或单元格区域

##  4. 单元格填充色

    '数组也可以设置格式？
    '数组除了数字类型外，当然没有颜色、字体等格式，但是别忘了range对象可以表示多个连续或不连续的单元格区域
    '利用上述特点，我们就是要数组构造单元格地址串，然后批量对单元格进行格式设置。
    '注意，单元格地址串不能>255，所以如果单元格操作过多，我们还需要分次分批设置单元格格式
    
    Sub 填充颜色()
    Range("a2:d2,a7:d7,a10:d10").Interior.ColorIndex = 3
    End Sub