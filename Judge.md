# VBA循环

## 1. 用于由数字规律
    Sub t2()
    Dim x As Integer
        For x = 2 To 18
            Range("d" & x) = Range("b" & x) * Range("c" & x)
        Next x
    End Sub

## 2. 用于无数字规律

    Sub t3()
    Dim rg As Range
        For Each rg In Range("d2:d18")
            rg = rg.Offset(0, -1) * rg.Offset(0, -2)
        Next rg
    End Sub

## 3. 无限循环 达到条件跳出

    Sub t4()
    Dim x As Integer
        x = 1
        Do
            x = x + 1
            Cells(x, 4) = Cells(x, 2) * Cells(x, 3)
        Loop Until x = 18
    End Sub

    Sub t5()
    Dim x As Integer
        x = 1
        Do While x < 18
            x = x + 1
            Cells(x, 4) = Cells(x, 2) * Cells(x, 3)
        Loop
    End Sub

# VBA判断

## 1. if判断

    Sub 判断2() '多条件判断
        If Range("a1").Value > 0 Then
            Range("b1") = "正数"
        ElseIf Range("a1") = 0 Then
            Range("b1") = "等于0"
        ElseIf Range("B1") <= 0 Then
            Range("b1") = "负数"
        End If
    End Sub

## 2. select判断

    Sub 判断2() '多条件判断
        Select Case Range("a1").Value
        Case Is > 0
            Range("b1") = "正数"
        Case Is = 0
            Range("b1") = "0"
        Case Else
            Range("b1") = "负数"
        End Select
    End Sub

## 3. IIF判断

    Sub 判断4()
        Range("a3") = IIf(Range("a1") <= 0, "负数或零", "负数")
    End Sub
