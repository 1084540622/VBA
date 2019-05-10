# END语句

## 1. END

作用：强制退出所有正在运行的程序。

## 2. Exit语句

    Sub e1()
     Dim x As Integer
        For x = 1 To 100
          Cells(1, 1) = x
          If x = 5 Then
            Exit Sub
          End If
         Next x
      Range("b1") = 100
     End Sub

# 分支语句

    'Goto语句,跳转到指定的地方
    Sub t1()
    Dim x As Integer
    Dim sr
    100:
        sr = Application.InputBox("请输入数字", "输入提示")
    If Len(sr) = 0 Or Len(sr) = 5 Then GoTo 100
    
    End Sub

    'gosub..return ,跳过去,再跳回来
    Sub t2()
    Dim x As Integer
        For x = 1 To 10
            If Cells(x, 1) Mod 2 = 0 Then GoSub 100
        Next x
    Exit Sub
    100:
        Cells(x, 1) = "偶数"
        Return          '跳到gosub 100 这一句
    End Sub

    'on error resume next '遇到错误,跳过继续执行下一句
    Sub t3()
    On Error Resume Next
    Dim x As Integer
        For x = 1 To 10
            Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
        Next x
    End Sub
    
    'on error goto  '出错时跳到指定的行数
    Sub t4()
    On Error GoTo 100
    Dim x As Integer
        For x = 1 To 10
            Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
        Next x
    Exit Sub
    100:
        MsgBox "在第" & x & "行出错了"
    End Sub
    
    'on error goto 0 '取消错误跳转
    Sub t5()
    On Error Resume Next
    Dim x As Integer
        For x = 1 To 10
            If x > 5 Then On Error GoTo 0
            Cells(x, 3) = Cells(x, 2) * Cells(x, 1)
        Next x
    Exit Sub

    End Sub