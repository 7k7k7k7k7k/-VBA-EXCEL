使用场景：
1：我现在有两个EXCEL工作表。
2：请遍历名为‘Sheet1’工作表中的A列单元格内容和B列单元格内容。
3：请遍历名为‘Sheet2’工作表中的C列单元格内容，G列单元格内容，M列单元格内容。

如果名为‘Sheet1’工作表中的A列单元格内容等于名为‘Sheet2’工作表中的C列单元格内容，并且，名为‘Sheet1’工作表中的B列单元格内容等于名为‘Sheet2’工作表中的G列单元格内容，则把名为‘Sheet2’工作表中的M列单元格内容内容复制粘贴到名为‘Sheet1’工作表中的D列单元格中去。
请注意：名为‘Sheet1’工作表中的A列单元格内容和B列单元格内容是數值型數據。
请注意：名为‘Sheet2’工作表中的C列单元格内容，G列单元格内容，M列单元格内容都是數值型數據。
请注意：名为‘Sheet1’工作表中的第一行是標題行。
请注意：名为‘Sheet2’工作表中的第一行是標題行。标题行不纳入比较范围。请使用VBA代码完成以上内容
VBA代码如下：

Sub CopyData()
    Dim wsSheet1 As Worksheet
    Dim wsSheet2 As Worksheet
    Dim lastRowSheet1 As Long
    Dim lastRowSheet2 As Long
    Dim i As Long, j As Long
    
    ' 设置工作表
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    Set wsSheet2 = ThisWorkbook.Sheets("Sheet2")
    
    ' 获取最后一行
    lastRowSheet1 = wsSheet1.Cells(wsSheet1.Rows.Count, "A").End(xlUp).Row
    lastRowSheet2 = wsSheet2.Cells(wsSheet2.Rows.Count, "C").End(xlUp).Row
    
    ' 循环遍历Sheet1工作表
    For i = 2 To lastRowSheet1 ' 从第2行开始，跳过标题行
        ' 循环遍历Sheet2工作表
        For j = 2 To lastRowSheet2 ' 从第2行开始，跳过标题行
            ' 如果Sheet1工作表中的A列和B列等于Sheet2工作表中的C列和G列
            If wsSheet1.Cells(i, "A").Value = wsSheet2.Cells(j, "C").Value And _
               wsSheet1.Cells(i, "B").Value = wsSheet2.Cells(j, "G").Value Then
                ' 将Sheet2工作表中的M列内容复制到Sheet1工作表中的D列
                wsSheet1.Cells(i, "D").Value = wsSheet2.Cells(j, "M").Value
                Exit For ' 匹配一次后跳出内部循环
            End If
        Next j
    Next i
End Sub
