Attribute VB_Name = "ģ��2"
Private Sub ��������()
    Dim Tmp_Err As Range    '��������
    Dim Tmp_Std As Range    '��׼ֵ������������
    Dim Tmp_Mav As Range    '����ֵ������������
    Dim Tmp_Pdc As String  '��Ʒ��
    Dim Tmp_Old As Range    '��������
    Dim Tmp_New As Range    '������
    Dim Max_Row As Integer  '���������
    Dim Tmp As Range    '��ʱrange����
    Dim Tmp2 As Range   '��ʱrange����
    Dim i As Integer    '��ʱ���ͱ���
    Dim j As Integer    '��ʱ���ͱ���
    Dim test As Range   '�����õı���
    
    Set Tmp_Err = Worksheets("ԭʼ����").Cells(2, 1)
    '�������
    Max_Row = Worksheets("��������").Range("A1").CurrentRegion.Rows.Count
    If Max_Row <> 1 Then
        Set Tmp = Worksheets("��������").Range(Worksheets("��������").Cells(2, 1), Worksheets("��������").Cells(Max_Row, 4))
    End If
    Tmp.Delete
    Let i = 2
    While Len(Tmp_Err.Value) <> 0
        Set Tmp_Std = Tmp_Err.Offset(2, 2)
        Set Tmp_Mav = Tmp_Err.Offset(3, 2)
        '��ȡ�޳�����
        Max_Row = Worksheets("�޳�����").Range("A1").CurrentRegion.Rows.Count
        Set Tmp = Worksheets("�޳�����").Range(Worksheets("�޳�����").Cells(1, 1), Worksheets("�޳�����").Cells(Max_Row, 1))
        '��ȡ��;��������
        Max_Row = Worksheets("������").Range("A1").CurrentRegion.Rows.Count
        Set Tmp2 = Worksheets("������").Range(Worksheets("������").Cells(1, 1), Worksheets("������").Cells(Max_Row, 1))
        '��ȡ��Ʒ��
        j = InStrRev(Tmp_Std.Value, "_����ר��_")
        Tmp_Pdc = Mid(Tmp_Std.Value, j - 11, 11)
        'ɸѡ����
        Let Worksheets("��������").Cells(i, 1).Value = Tmp_Err.Value
        Let Worksheets("��������").Cells(i, 2).Value = Tmp_Std.Value
        Let Worksheets("��������").Cells(i, 3).Value = Tmp_Mav.Value
        Set test = Tmp.Find(Tmp_Pdc)
        If IsNumeric(Tmp.Find(Tmp_Pdc)) Then
            Let Worksheets("��������").Cells(i, 4).Value = "�޳�����"
        End If
        If IsNumeric(Tmp2.Find(Tmp_Pdc)) Then
            Let Worksheets("��������").Cells(i, 4).Value = "��;����"
        End If
        i = i + 1
        '��һ������
        Set Tmp_Err = Tmp_Err.Offset(6, 0)
    Wend
    '�����������Ƿ�������
    Max_Row = Worksheets("��������").Range("A1").CurrentRegion.Rows.Count
    Set Tmp_Old = Worksheets("��������").Range(Worksheets("��������").Cells(2, 6), Worksheets("��������").Cells(Max_Row, 6))
    Max_Row = Worksheets("��������").Range("A1").CurrentRegion.Rows.Count
    Set Tmp_New = Worksheets("��������").Range(Worksheets("��������").Cells(2, 2), Worksheets("��������").Cells(Max_Row, 2))
    For Each Tmp In Tmp_Old
        If Tmp_New.Find(Tmp) Is Nothing Then
            Let Tmp.Offset(0, 8).Value = "������"
        End If
    Next
    '����������ݼ��ظ���������
    For Each Tmp In Tmp_New
        Set Tmp2 = Tmp_Old.Find(Tmp)
        If Tmp2 Is Nothing Then
            Let Tmp.Offset(0, 3).Value = "��"
        Else
            If Tmp2.Offset(0, 8).Value = "������" Then
                Let Tmp.Offset(0, 4).Value = "��"
            End If
        End If
    Next
    '��ʽ����
    With Range(Worksheets("��������").Cells(2, 1), Worksheets("��������").Cells(Max_Row, 1))
        .HorizontalAlignment = xlCenter
    End With
End Sub

