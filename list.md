# 数组

## 1. 求最

    Sub s()
    Dim arr1()
    
    arr1 = Array(1, 12, 4, 5, 19)
    
    MsgBox "1, 12, 4, 5, 19最大值" & Application.Max(arr1)
    MsgBox "1, 12, 4, 5, 19最小值:" & Application.Min(arr1)
    MsgBox "1, 12, 4, 5, 19第二大值：" & Application.Large(arr1, 2)
    MsgBox "1, 12, 4, 5, 19第二小值：" & Application.Small(arr1, 2)
    
    End Sub

## 2. 求和

    application.Sum

## 3. 统计个数

    Sub s1()
     
     Dim arr1, arr2(0 To 10), x
     arr1 = Array("a", "3", "", 4, 6)
     For x = 0 To 4
       arr2(x) = arr1(x)
     Next x
     
     MsgBox "数组1的数字个数：" & Application.Count(arr1)  '2
     
     MsgBox "数组2的已填充数值的个数" & Application.CountA(arr2) '11
     
     End Sub

## 4. 查找

    Sub s2()
      Dim arr
      On Error Resume Next
      arr = Array("a", "c", "b", "f", "d")
      MsgBox Application.Match("f", arr, 0)
     If Err.Number = 13 Then
        MsgBox "查找不到"
      End If
     End Sub
    
## 5. split函数

    '按分隔符把字符串截取成VBA数组,该数组是一维数组，编号从0开始
 
     'split(字符串,分隔符)
   
    Sub t1()
      Dim sr, arr
      sr = "A-BC-FGR-H"
      arr = VBA.Split(sr, "-")
      MsgBox Join(arr, ",")
    End Sub

##  6. Filter函数

     '按条件筛选符合条件的值组成一个新的数组

     'Filter(数组,筛选条件,是/否)
     
     '注：如果是（true）则返回包含的数组，如果否则返回非包含的数组
    Sub t2()
     Dim arr, arr1, arr2
     arr = Application.Transpose(Range("A2:A10"))
     arr1 = VBA.Filter(arr, "W", True)
     arr2 = VBA.Filter(arr, "W", False)
     Range("B2").Resize(UBound(arr1) + 1) = Application.Transpose(arr1)
     Range("C2").Resize(UBound(arr2) + 1) = Application.Transpose(arr2)
    End Sub

## 7. 数组的大小

    '数组是用编号排序的，那么如何获得一个数组的大小呢

    'Lbound(数组) 可以获取数组的最小下标(编号)
    'Ubound(数组) 可以获取数组的最大上标(编号)
    'Ubound(数组,1) 可以获得数组的行方面(第1维)最大上标
    'Ubound(数组,2) 可以获得数组的列方向(第2维)的最大上标

    ReDim Preserve arr() 可以声明一个动态大小的数组，而且可以保留原来的数值，就相当于厂房小了，可以改扩建增大，但是它只能
        '让最未维实现动态，如果是一维不存在最未维，只有一维

## 8. index函数

    '调用该工作表函数可以把二维数组的某一列或某一行截取出来，构成一个新的数组。
     ' Application.Index(二维数组,0,列数)) 返回二维数组
     ' Application.Index(二维数组,行数,0)) 返回一维数组
    Sub t3()
     Dim arr, arr1, arr2
      arr = Range("a2:d6")
      arr1 = Application.Index(arr, , 1)
      arr2 = Application.Index(arr, 4, 0)
      Stop
    End Sub

## 9. vlookup函数

    'Vlookup函数的第一个参数可以用VBA数组，返回的也是一个VBA数组
    Sub t4()
    Dim arr, arr1
      arr = Range("a2:d6")
      arr1 = Application.VLookup(Array("B", "C"), arr, 4, 0)
    End Sub

## 10. Sumif函数和Countif函数

    'Countif和sumif函数的第二个参数都可以使用数组，所以也可以返回一个VBA数组，如：
     Sub t5()
     Dim T
     T = Timer
       Dim arr
       arr = Application.SumIf(Range("a2:a10000"), Array("B", "C", "G", "R"), Range("B2:B10000"))
     MsgBox Timer - T
     Stop
     End Sub