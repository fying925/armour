---
title: EXCEL
abbrlink: 2f6440bb
date: 2023-03-29 09:06:02
tags: EXCEL
categories: EXCEL
---

### 1.第一项多个工作薄合并宏代码
```
Sub 合并当前工作簿下的所有工作表()
Application.ScreenUpdating = False
For j = 1 To Sheets.Count
			If Sheets(j).Name <> ActiveSheet.Name Then
							X = Range("A65536").End(xlUp).Row + 1
							Sheets(j).UsedRange.Copy Cells(X, 1)
			End If
Next
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "当前工作簿下的全部工作表已经合并完毕！", vbInformation, "提示"
End Sub
```

### 2.多个工作薄名字合并到一个工作簿宏代码
```
Sub 合并当前工作簿下的所有工作表()
Application.ScreenUpdating = False
For j = 1 To Sheets.Count
   If Sheets(j).Name <> ActiveSheet.Name Then
       X = Range("A65536").End(xlUp).Row + 1
       Sheets(j).UsedRange.Copy Cells(X, 1)
   End If
Next
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "当前工作簿下的全部工作表已经合并完毕！", vbInformation, "提示"
End Sub
Sub a()
For Each sh In Sheets
k = k + 1
Cells(k, 1) = sh.Name
Next
End Sub
```

### 3.多个工作表合并宏代码
```
Sub 合并当前目录下所有工作簿的全部工作表()
Dim MyPath, MyName, AWbName
Dim Wb As Workbook, WbN As String
Dim G As Long
Dim Num As Long
Dim BOX As String
Application.ScreenUpdating = False
MyPath = ActiveWorkbook.Path
MyName = Dir(MyPath & "\" & "*.xls")
AWbName = ActiveWorkbook.Name
Num = 0
Do While MyName <> ""
If MyName <> AWbName Then
Set Wb = Workbooks.Open(MyPath & "\" & MyName)
Num = Num + 1
With Workbooks(1).ActiveSheet
.Cells(.Range("B65536").End(xlUp).Row + 2, 1) = Left(MyName, Len(MyName) - 4)
For G = 1 To Sheets.Count
Wb.Sheets(G).UsedRange.Copy .Cells(.Range("B65536").End(xlUp).Row + 1, 1)
Next
WbN = WbN & Chr(13) & Wb.Name
Wb.Close False
End With
End If
MyName = Dir
Loop
Range("B1").Select
Application.ScreenUpdating = True
MsgBox "共合并了" & Num & "个工作薄下的全部工作表。如下：" & Chr(13) & WbN, vbInformation, "提示"
End Sub
```

### 4.劳务报酬个税反算
根据个税反算税前应发金额
```
=ROUND(IF(A1<=640,A1/0.2+800,IF(A1<=4000,A1/(0.8*0.2),IF(A1<=13000,(A1+2000)/(0.8*0.3),(A1+7000)/(0.8*0.4)))),2)
```
根据税后实发金额反算个税
```
=ROUND(IF(A1<=800,0,IF(A1<=3360,(A1-800)/4,IF(A1<=21000,0.16*A1/0.84,IF(A1<=49500,(0.24*A1-2000)/0.76,(0.32*A1-7000)/0.68)))),2)
```
劳务报酬个人所得税计算公式  
应纳税所得额 = 劳务报酬（少于4000元） - 800元  
应纳税所得额 = 劳务报酬（超过4000元） × （1 - 20%）  
应纳税额 = 应纳税所得额 × 适用税率 - 速算扣除数

### 5.根据班级代码编学号
```
=a2&TEXT(COUNTIF(a$2:a2,a2),"00")
```

### 6.根据班级名称匹配班级代码（f为要匹配班级代码列，b2查找班级名称，g为匹配班级名称）
```
=INDEX(F:F,MATCH(B2,G:G,0))
```

### 7.根据身份证号（18位）取性别
```
=IF(mod(mid(a2,17,1),2)=1,"男","女"
```

### 8.根据身份证号（18位）取年龄
```
=YEAR(NOW())-MID(A2,7,4)
```

### 9.excel 取某符号后的内容  
如：[070118001]数控1821     取数控1821
```
=SUBSTITUTE(MID(b2,FIND("]",b2,1)+1,100),")",,2)
```
