# 函数


## 1. 测试函数

### 1. IsNumeric(x)     是否为数字 ,返回 Boolean 结果， True or False
### 2. IsDate(x)        是否是日期 ,返回 Boolean 结果， TrueorFalse
### 3. IsEmpty(x)       是否为 Empty, 返回 Boolean 结果， TrueorFalse
### 4. IsArray(x)       指出变量是否为一个数组
### 5. IsError(expression)      指出表达式是否为一个错误值
### 6. IsNull(expression)       指出表达式是否不包含任何有效数据
### 7. IsObject(identifier)     指出标识符是否表示对象变量


## 2. 数学函数

### 1. Sin(X) 、 Cos(X) 、Tan(X) 、Atan(x)	三角函数，单位为弧度
### 2. Log(x)	返回 x 的自然对数
### 3. Exp(x)	返回 ex
### 4. Abs(x)	返回绝对值
### 5. Int(number) 、Fix(number)	都返回参数的整数部分，区别： Int 将-8.4 转换成 -9，而 Fix 将-8.4 转换成 -8
### 6. Sgn(number)	返回一个 Variant(Integer) ，指出参数的正负号
### 7. Sqr(number)	返回一个 Double ，指定参数的平方根
### 8. VarType(varname)	返回一个 Integer ，指出变量的子类型
### 9. Rnd （ x）	返回 0-1 之间的单精度数据， x 为随机种子


## 3. 字符串函数

### 1. Trim(string)	去掉 string 左右两端空白
### 2. Ltrim(string)	去掉 string 左端空白
### 3. Rtrim(string)	去掉 string 右端空白
### 4. Len(string)	计算 string 长度
### 5. Left(string,x)	取 string 左段 x 个字符组成的字符串
### 6. Right(string,x)	取 string 右段 x 个字符组成的字符串
### 7. Mid(string,start,x)	取 string 从 start 位开始的 x 个字符组成的字符串
### 8. Ucase(string)	转换为大写
### 9. Lcase(string)	转换为小写
### 10. Space(x)	返回 x 个空白的字符串
### 11. Asc(string)	返回一个 integer ，代表字符串中首字母的字符代码
### 12. Chr(charcode)	返回 string, 其中包含有与指定的字符代码相关的字符


## 4. 转换函数


### 1. CBool(expression)	转换为 Boolean 型
### 2. CByte(expression)	转换为 Byte 型
### 3. CCur(expression)	转换为 Currency 型
### 4. CDate(expression)	转换为 Date 型
### 5. CDbl(expression)	转换为 Double 型
### 6. CDec(expression)	转换为 Decemal 型
### 7. CInt(expression)	转换为 Integer 型
### 8. CLng(expression)	转换为 Long 型
### 9. CSng(expression)	转换为 Single 型
### 10. CStr(expression)	转换为 String 型
### 11. CVar(expression)	转换为 Variant 型
### 12. Val(string)	转换为数据型
### 13. Str(number)	转换为 String


## 5. 时间函数


### 1. Now	返回一个 Variant(Date) ，根据计算机系统设置的日期和时间来指定日期和时间
### 2. Date	返回包含系统日期的 Variant(Date)
### 3. Time	返回一个指明当前系统时间的 Variant(Date)
### 4. Timer	返回一个 Single ，代表从午夜开始到现在经过的秒数
### 5. TimeSerial(hour,minute,second)	返回一个 Variant(Date) ，包含具有具体时、分、秒的时间
### 6. DateDiff(interval,date1,date2[,firstdayofweek[,firstweekofyear]])	返回 Variant(Long) 的值，表示两个指定日期间的时间间隔数目
### 7. Second(time)	其值为 0 到 59 之间的整数，表示一分钟之中的某个秒
### 8. Minute(time)	其值为 0 到 59 之间的整数，表示一小时中的某分钟
### 9. Hour(time)	返回一个 Variant(Integer) ，其值为 0 到 23 之间的整数，表示一天之中的某一钟点
### 10. Day(date)	返回一个 Variant(Integer) ，其值为 1 到 31 之间的整数，表示一个月中的某一日
### 11. Month(date)	返回一个 Variant(Integer) ，其值为 1 到 12 之间的整数，表示一年中的某月
### 12. Year(date)	返回 Variant(Integer) ，包含表示年份的整数
### 13. Weekday(date,[firstdayofweek])	返回一个 Variant(Integer) ，包含一个整数，代表某个日期是星期几
