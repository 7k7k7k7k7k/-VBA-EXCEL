运用场景：
我现在有一个名为Sheet1的EXCEL工作表。这个工作表从C2单元格到C24单元格内容中都有内容（比如2000，2001依次排列到2023）。请复制C2单元格到C24单元格中的内容向下粘贴474次。例如复制名为Sheet1的EXCEL工作表中C2单元格到C24单元格内容粘贴在名为Sheet1的EXCEL工作表中的C25单元格到C47单元格。以此类推474次。请注意名为‘Sheet1的’工作表没有标题行。请注意名为‘Sheet1中的C列单元格内容是数值格式。使用VBA进行.我当时是为了构建给每一个企业个体构建2000年到2023年的数据，而企业实在太多需要构建就使用了VBA。
代码如下：
Sub 复制粘贴循环()
    Dim ws As Worksheet
    Dim i As Long, j As Long
    Dim cellValue As Variant

    ' 设置要操作的工作表
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' 获取C2到C24单元格的内容
    cellValue = ws.Range("C2:C24").Value

    ' 循环复制粘贴内容
    For i = 1 To 474
        ' 粘贴内容到对应的位置
        ws.Range("C" & (i - 1) * 23 + 25 & ":C" & i * 23 + 24).Value = cellValue
    Next i
End Sub
