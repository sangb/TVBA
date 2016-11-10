Attribute VB_Name = "模块2"
Private Sub 数据整理()
    Dim Tmp_Err As Range    '错误类型
    Dim Tmp_Std As Range    '标准值，即综资数据
    Dim Tmp_Mav As Range    '现网值，即网数数据
    Dim Tmp_Pdc As String  '产品号
    Dim Tmp_Old As Range    '存量数据
    Dim Tmp_New As Range    '新数据
    Dim Max_Row As Integer  '已用最大行
    Dim Tmp As Range    '临时range变量
    Dim Tmp2 As Range   '临时range变量
    Dim i As Integer    '临时整型变量
    Dim j As Integer    '临时整型变量
    Dim test As Range   '测试用的变量
    
    Set Tmp_Err = Worksheets("原始数据").Cells(2, 1)
    '清除数据
    Max_Row = Worksheets("问题数据").Range("A1").CurrentRegion.Rows.Count
    If Max_Row <> 1 Then
        Set Tmp = Worksheets("问题数据").Range(Worksheets("问题数据").Cells(2, 1), Worksheets("问题数据").Cells(Max_Row, 4))
    End If
    Tmp.Delete
    Let i = 2
    While Len(Tmp_Err.Value) <> 0
        Set Tmp_Std = Tmp_Err.Offset(2, 2)
        Set Tmp_Mav = Tmp_Err.Offset(3, 2)
        '读取剔除号码
        Max_Row = Worksheets("剔除号码").Range("A1").CurrentRegion.Rows.Count
        Set Tmp = Worksheets("剔除号码").Range(Worksheets("剔除号码").Cells(1, 1), Worksheets("剔除号码").Cells(Max_Row, 1))
        '读取在途工单号码
        Max_Row = Worksheets("已销户").Range("A1").CurrentRegion.Rows.Count
        Set Tmp2 = Worksheets("已销户").Range(Worksheets("已销户").Cells(1, 1), Worksheets("已销户").Cells(Max_Row, 1))
        '获取产品号
        j = InStrRev(Tmp_Std.Value, "_语音专线_")
        Tmp_Pdc = Mid(Tmp_Std.Value, j - 11, 11)
        '筛选数据
        Let Worksheets("问题数据").Cells(i, 1).Value = Tmp_Err.Value
        Let Worksheets("问题数据").Cells(i, 2).Value = Tmp_Std.Value
        Let Worksheets("问题数据").Cells(i, 3).Value = Tmp_Mav.Value
        Set test = Tmp.Find(Tmp_Pdc)
        If IsNumeric(Tmp.Find(Tmp_Pdc)) Then
            Let Worksheets("问题数据").Cells(i, 4).Value = "剔除号码"
        End If
        If IsNumeric(Tmp2.Find(Tmp_Pdc)) Then
            Let Worksheets("问题数据").Cells(i, 4).Value = "在途销户"
        End If
        i = i + 1
        '下一条数据
        Set Tmp_Err = Tmp_Err.Offset(6, 0)
    Wend
    '检查存量数据是否已消除
    Max_Row = Worksheets("存量数据").Range("A1").CurrentRegion.Rows.Count
    Set Tmp_Old = Worksheets("存量数据").Range(Worksheets("存量数据").Cells(2, 6), Worksheets("存量数据").Cells(Max_Row, 6))
    Max_Row = Worksheets("问题数据").Range("A1").CurrentRegion.Rows.Count
    Set Tmp_New = Worksheets("问题数据").Range(Worksheets("问题数据").Cells(2, 2), Worksheets("问题数据").Cells(Max_Row, 2))
    For Each Tmp In Tmp_Old
        If Tmp_New.Find(Tmp) Is Nothing Then
            Let Tmp.Offset(0, 8).Value = "已消除"
        End If
    Next
    '检查新增数据及重复报错数据
    For Each Tmp In Tmp_New
        Set Tmp2 = Tmp_Old.Find(Tmp)
        If Tmp2 Is Nothing Then
            Let Tmp.Offset(0, 3).Value = "是"
        Else
            If Tmp2.Offset(0, 8).Value = "已消除" Then
                Let Tmp.Offset(0, 4).Value = "是"
            End If
        End If
    Next
    '格式整理
    With Range(Worksheets("问题数据").Cells(2, 1), Worksheets("问题数据").Cells(Max_Row, 1))
        .HorizontalAlignment = xlCenter
    End With
End Sub

