Attribute VB_Name = "table_transpose"
'������, ��������������� �������, � ������� ��������� ������. �������� ������� �������� ��� ���������; �����������������-
'���������� � ����� ���������, ����� ������� ������.

'�������� �. 3645
'23.04.2012


'���������, �������� ���������� �� ������������
Public Type cellInfo
    rs As Integer   'rowspan
    cs As Integer   'colspan
    w As Integer    'width
    shiftH As Integer '�������� �� ����������� - ����� ��� ���������� ��������������
    shiftV As Integer '���������� - �� ���������
End Type

'���������, �������� ������� �����
Public Type indexes
    i As Integer    '������ �� ���������
    j As Integer    '������ �� �����������
End Type


Sub getinfo(ByRef tinfo() As cellInfo, ByRef rinfo() As Integer, t As Table)
'tinfo - ������, �������� ������ �� ������������ ������
'rinfo - ������, ���������� ���������� ����� � �������
't - ��������� �� �������� �������
'����� ��������� �������, �������� ������ �� ������������ �� �����
    For Each el In rinfo  '��������� ������ (����� ����������� �����)
        el = 0
    Next el
    
    For Each aCell In t.Range.cells   '�������� ��� �������, ������� ���������� � �� �������
        aCell.Select
        i = aCell.rowIndex
        j = aCell.ColumnIndex
        tinfo(i, j).rs = (Selection.Information(wdEndOfRangeRowNumber) - Selection.Information(wdStartOfRangeRowNumber))  '���������� ������ ����� -1
        tinfo(i, j).w = aCell.width         ' ������ ������ (��������� ��� ������������ ����������� �������� cs
        tinfo(i, j).cs = 0                  '���������� ������ �������� (����� ���������� �����)
        rinfo(i) = rinfo(i) + 1             '���������� ����� � ������ (��� ������ ������)
        tinfo(i, j).shiftH = 0              '�������� �� ����������� - ����� ��� ���������� ��������������, ����� ����������� �����
        tinfo(i, j).shiftV = 0              '���������� - �� ���������
    Next aCell
    
    Call mergingResearch(tinfo, rinfo, t)   '��������� ������ ������� �� ������������ - cs, shiftH � shiftV
    
End Sub

Function getPerfIndex(ByRef rinfo() As Integer, maxColsAm As Integer)
'rinfo - ������, ���������� ���������� ����� � �������
'maxColsAm - ���������� �������� � ������� = ���������� ����� � "���������" ������
'������� ���������� ������ "���������" ������
    getPerfIndex = 1
    For i = 1 To UBound(rinfo)          '������������� ��� ������ �������
        If rinfo(i) = maxColsAm Then    '-���� ���������� ����� � ������ ����� ���������� �������� � �������,
            getPerfIndex = i            '������ ��� - "���������" ������, ���������� �� ������
            Exit For
        End If
    Next i
End Function

Sub mergingResearch(ByRef tinfo() As cellInfo, ByRef rinfo() As Integer, t As Table)
'tinfo - ������, �������� ������ �� ������������ ������
'rinfo - ������, ���������� ���������� ����� � �������
't - ��������� �� �������� �������
'����� ��������� ������ tinfo ������� �� ������������ - cs, shiftH � shiftV

    Dim i, j, perfIndex, w, k, shift As Integer
    'i,j,k - ���������
    'perfIndex - ������ "���������" ������
    'w - ������ ��� ������������ ��������� width ����� ��������� ������ -- ��� ��������� � ������� ��������������� ������
    'shift - �������� MSWord-������� ������ ������������ �����������, ����������� � ������ ������� �����������

    perfIndex = getPerfIndex(rinfo, t.Columns.count)    '���� ������ "���������" ������

    For Each c In t.Range.cells                 '�������� ��� ������ � �������
        shift = 0           '�������� ��������
        If (c.rowIndex <> perfIndex) Then           '"���������" ������ �� ������������� - � ��� ��� �����������
            i = c.rowIndex                          '���������� �������
            j = c.ColumnIndex
            k = j
            w = tinfo(perfIndex, k).w               '������� ������ ������, ����������� ���������������, ����������� � "���������" ������
            While (tinfo(i, j).w - w > 5)           '���� ������ �� ��������� (5 - ���������� ����������� ����� ����� ������)
                w = w + tinfo(perfIndex, k).w       '����������� �������� cs (colspan) ��������������� �����������
                tinfo(i, j).cs = tinfo(i, j).cs + 1 '��� ���� ���������� �������� ��������� ������ w �� ������ ��������� ������ "���������" ������
                k = k + 1
                shift = shift + 1
            Wend

            If tinfo(i, j).cs <> 0 Then                     '� ������, ���� ������ ���� ����������, ������� ��������� �������� ���.��������
                For z = i To (i + tinfo(i, j).rs)           '��� ���� �����, ��� ������� ����� �������� �� ����� �����������
                    For k = (j + 1) To UBound(tinfo, 2)
                        tinfo(z, k).shiftH = tinfo(z, k).shiftH + shift
                    Next k
                Next z
            End If
        End If
    Next c
    

    
    For i = 1 To UBound(tinfo, 1)           '����� ����, ��� ��� ����������� ��������, ������� ��������� ���������� � ������������ ������
            For j = 1 To UBound(tinfo, 2)   '�������� ����� - ��� ����������� ��� ����������� ���������� ��������������. �� ������� �������,
                If tinfo(i, j).rs <> 0 Then '������ ��������� �����������
                    For z = i + tinfo(i, j).rs + 1 To UBound(tinfo, 1)  '���� ������ ����� ������������ �����������, ������� ��������� �������� ����.��������
                        For k = j + tinfo(i, j).shiftH To j + tinfo(i, j).cs    '��� ���� �����, ��� ������� ����� �������� �� ����� �����������
                            tinfo(z, k).shiftV = tinfo(z, k).shiftV + tinfo(i, j).rs
                        Next k
                    Next z
                End If
            Next j
    Next i

End Sub

Sub applyMerges(ByRef tinfo() As cellInfo, newT As Table)
'tinfo - ������, �������� ������ �� ������������ ������
'newT - ��������� �� ����� �������
'����� ���������� ������ ����� ������� �� ��������� ������, ���������� � tinfo, �� ���� ������������ �������
    For i = 1 To UBound(tinfo, 2)                       '�������� ������, �������� ���������� �� ������������
            For j = 1 To UBound(tinfo, 1)
                If tinfo(j, i).w <> 0 Then              '�, � ������, ���� ��������������� ������� ������ ���������� � ������������ ������
                    If (tinfo(j, i).cs + tinfo(j, i).rs) <> 0 Then  '���������� �����������, ���������� ����������������� ������ ����� �������
                        newT.Cell(i + tinfo(j, i).shiftH, j - tinfo(j, i).shiftV).Merge newT.Cell(i + tinfo(j, i).shiftH + tinfo(j, i).cs, j + tinfo(j, i).rs)
                    End If
                End If
            Next j
    Next i
End Sub

Sub fillTable(t As Table, newT As Table, ByRef tinfo() As cellInfo)
't - ��������� �� �������� �������
'newT - ��������� �� ����� �������
'tinfo - ������, �������� ������ �� ������������ ������
'����� ��������� ����� ������� ������� �� �������� �������
    Dim i, j, m, n, k As Integer             '���������
    Dim cinfo() As indexes         '������, �������� ������ �� �������� ����� - ����� ��� ��������� ����� ���� foreach
    ReDim cinfo(1 To t.Rows.count * t.Columns.count)
    
    k = 1
    For Each c In t.Range.cells             '�������� ��� ������ �������� �������, ��������� �� �������
        cinfo(k).i = c.rowIndex
        cinfo(k).j = c.ColumnIndex
        k = k + 1                           '� �� ����������
    Next c
        
    While k <> 1            '����� ���� - �������� �� �������� ������� "�������� foreach'�� - �� ������ ������ ������ � ����� �������
        k = k - 1
        
        i = cinfo(k).i                      '���������� �������, �� ��������� ������� �������� ���������� ������ �������� �������
        j = cinfo(k).j                      '� ������ ����� �������, �� ������� ������� ������� �� �����������������
        m = j + tinfo(i, j).shiftH          ' m � n - ��������� �������, ����������� �� ������ � ����������������� �������
        n = i - tinfo(i, j).shiftV
        
        '�������� �������� �� �������� ������� - � �������� � ���������������:
        t.Cell(i, j).Range.Copy
        newT.Cell(m, n).Shading.BackgroundPatternColorIndex = t.Cell(i, j).Shading.BackgroundPatternColorIndex
        If t.Cell(i, j).Shading.BackgroundPatternColor <> -1 Then
            newT.Cell(m, n).Shading.BackgroundPatternColor = t.Cell(i, j).Shading.BackgroundPatternColor
        End If
        newT.Cell(m, n).Borders = t.Cell(i, j).Borders
        newT.Cell(m, n).Borders.OutsideLineStyle = t.Cell(i, j).Borders.OutsideLineStyle
        newT.Cell(m, n).Borders.OutsideColorIndex = t.Cell(i, j).Borders.OutsideColorIndex
        newT.Cell(m, n).Borders.OutsideColor = t.Cell(i, j).Borders.OutsideColor
        newT.Cell(m, n).Borders.Enable = t.Cell(i, j).Borders.Enable
        newT.Cell(m, n).Borders.DistanceFromTop = t.Cell(i, j).Borders.DistanceFromTop
        newT.Cell(m, n).Borders.DistanceFromRight = t.Cell(i, j).Borders.DistanceFromRight
        newT.Cell(m, n).Borders.DistanceFromLeft = t.Cell(i, j).Borders.DistanceFromLeft
        newT.Cell(m, n).Borders.DistanceFromBottom = t.Cell(i, j).Borders.DistanceFromBottom
        newT.Cell(m, n).Range.PasteAndFormat (wdPasteDefault)
    Wend
End Sub

Sub createTransponedTable(ByRef tinfo() As cellInfo, t As Table)
'tinfo - ������, �������� ������ �� ������������ ������
't - ��������� �� �������� �������
'����� ������� ����� �������, ��������� �� ������ � �������� ������� �� �������� �������, ������������ ��

    Selection.EndKey Unit:=wdStory  '����� ������� ����������� � ����� ���������,
    Selection.TypeParagraph         '����� ������� ������

    '������� ����� �������, ��������� ������ ������
    ActiveDocument.tables.Add Range:=Selection.Range, NumRows:=t.Columns.count, NumColumns:= _
        t.Rows.count, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed
    Dim newT As Table   '��������� �� ����� �������
    Set newT = Selection.tables(1)
    With newT
        .Style = t.Style '"����� �������"
        .ApplyStyleHeadingRows = t.ApplyStyleFirstColumn
        .ApplyStyleLastRow = t.ApplyStyleLastColumn
        .ApplyStyleFirstColumn = t.ApplyStyleHeadingRows
        .ApplyStyleLastColumn = t.ApplyStyleLastRow
        .Borders.OutsideLineStyle = t.Borders.OutsideLineStyle
        .Borders.InsideLineStyle = t.Borders.InsideLineStyle
    End With

   Call applyMerges(tinfo, newT)    '���������� ������
   Call fillTable(t, newT, tinfo)   '��������� ������� ���������� (�����, �����)
End Sub

Sub table_transpose()
'������, ��������������� �������, � ������� ��������� ������. �������� ������� �������� ��� ���������; �����������������-
'���������� � ����� ���������, ����� ������� ������.
    Dim t As Table                  '������������� �������, � �������
    Set t = Selection.tables(1)     '��������� ������

    Dim tinfo() As cellInfo         '������, �������� ������ �� ������������ ������
    ReDim tinfo(1 To t.Rows.count, 1 To t.Columns.count)
    Dim rinfo() As Integer          '������, ���������� ���������� ����� � �������
    ReDim rinfo(1 To t.Rows.count)
    Call getinfo(tinfo, rinfo, t)   '�������� ���������� �� ������������ ����� �������� �������
    Call createTransponedTable(tinfo, t)    '�, �� ��������� ���������� ����������, ������� ����� �������, �������� ��
                                            '����������� �� �������� �������

End Sub

