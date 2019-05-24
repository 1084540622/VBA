# 字典
    字典简介
    字典对象相当于一种联合数组，它是由具有唯一性的关键字（Key）和它的项（Item）联合组成
    VBA字典有6个方法Add , Keys, Items, Exists, Remove, RemoveAll
    VBA字典有4个属性Count , Key, Item, CompareMode

## 1. 创建字典
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

## 2. 添加
    dict.Add "A", 300

## 3. 统计数量
    n = dict.Count

## 4. 删除
    dict.Remove ("A")

## 5. 判断字典是否已存在
    dict.exists ("A")

## 6. 取关键字对应的值，注意在使用前需要判断是否存在key，否则dict中会多出一条记录
    value = dict.Item("A")

## 7. 修改关键字对应的值,如不存在则创建新的项目
    dict.Item("A") = 400

## 8. 对字典进行循环
    k = dict.keys
    v = dict.Items
    For i = 0 To dict.Count - 1
        key = k(i)
        Value = v(i)
        MsgBox key & Value
    Next

## 9. 清空字典
    dict.Removeall