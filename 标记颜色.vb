使用场景：如果EXCEL表格中数据很庞大，很多列很多行，对EXCEL的操作栏不熟悉的人可以使用VBA写出要求。
我现在有一个名为sheet1的EXCEL工作表。为了方便我查看数据，请把J列到AB列的中不为0的单元格全部标为黄色。请注意名为sheet1的EXCEL工作表第一行是标题行。请注意名为sheet1的EXCEL工作表J列到AB列的单元格格式是数值。使用VBA进行。（其中标记什么范围，标记什么颜色可以自行修改）
代码如下：
Sub 标记非零单元格为黄色()
    Dim ws As Worksheet
    Dim cell As Range
    Dim lastRow As Long, lastColumn As Long
    Dim rng As Range

    ' 设置要操作的工作表
    Set ws = ThisWorkbook.Sheets("no")

    ' 获取最后一行和最后一列
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    ' 循环遍历 J 列到 AB 列中的每个单元格
    For Each cell In ws.Range(ws.Cells(2, 10), ws.Cells(lastRow, lastColumn))
        ' 检查单元格是否不为 0
        If cell.Value <> 0 Then
            ' 将不为 0 的单元格添加到要标记的范围中
            If rng Is Nothing Then
                Set rng = cell
            Else
                Set rng = Union(rng, cell)
            End If
        End If
    Next cell

    ' 如果找到了要标记的单元格，则设置它们的背景色为黄色
    If Not rng Is Nothing Then
        rng.Interior.Color = RGB(255, 255, 0) ' 黄色
    End If
End Sub
