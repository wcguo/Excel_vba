Sub 批量导出重复格式内容()
    Dim pptPre As Presentation
    Dim p, C As Long
    Dim n As Integer
    Dim myPath As String
    Dim appExcel As Object
    Dim myexcel As Object
    Dim mysheet As Object
    Dim rcount As Long
    
    Set pptPre = ActivePresentation
    
    myPath = "C:\Users\xl\media\" '图片位置
    
    Set appExcel = CreateObject("Excel.Application") '创建excel对象
    Set myexcel = appExcel.Workbooks.Open("C:\123.xlsx") '打开工作表
    Set mysheet = myexcel.sheets("Sheet1") '创建工作表对象
    rcount = mysheet.Cells(mysheet.Rows.Count, "A").End(3).Row '获取工作表最后一行行号
    
    For p = 4 To rcount '从第2行到最后一行

            n = p - 3

            
            For C = 2 To 3 '循环插入文本框
                With ActivePresentation.Slides(n)
                    With .Shapes.AddTextbox(msoTextOrientationHorizontal, 400, 80 + C * 5, 70, 50)     '文本框坐标及长宽
                         .TextFrame.TextRange.Font.Color = vbBlack '字体颜色
                         .TextFrame.TextRange.Font.Size = 18 '字号
                         .TextFrame.TextRange.Text = mysheet.Cells(p, C).Value '文本内容
                    End With
                End With
            Next C

    Next p
    
    myexcel.Close
    
    Set pptPre = Nothing
    Set appExcel = Nothing
    Set myexcel = Nothing
    Set mysheet = Nothing

End Sub
