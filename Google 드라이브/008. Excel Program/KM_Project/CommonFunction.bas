Attribute VB_Name = "CommonFunction"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ���� ���� : ���� ��ο� ���ϸ� ���� ���ָ� ��.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function openFile(filePath As Variant, fileName As Variant)

    Workbooks.Open filePath + fileName
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ���� �ݱ� : ���� ��ο� ���ϸ� ���� ���ָ� ��.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function closeFile(fileName As Variant)

    With Workbooks(fileName)
    .Save
    .Close
    
    End With
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ����� : ���� ��ο� ���ϸ� ���� ���ָ� ��.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function createEmptyFile(filePath As Variant, fileName As Variant)
    
    Workbooks.Add
    
    With ActiveWorkbook
        
        .SaveAs filePath + fileName
        .Close
    
    End With
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ����: ���� ��ο� ���ϸ� ���� ���ָ� ��.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function copyFile(srcShtName As Variant, srcRange As Variant)
    
    ActiveWorkbook.Worksheets(srcShtName).Range(srcRange).Copy

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ����: area 1,2,3,���ῡ ���� �������� �ٸ���
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function copyAreaFile(srcShtName As Variant, srcRange As Variant, area As Integer)
    
    Select Case area
        
        Case 1
            With ActiveWorkbook
                With .Worksheets(srcShtName)
                    With .Range(.Range(srcRange), .Range(srcRange).End(xlDown).End(xlToRight).End(xlToRight).End(xlToRight))
                    .Copy
                    End With
                End With
            End With
        Case 2
            With ActiveWorkbook
                With .Worksheets(srcShtName)
                    With .Range(srcRange).CurrentRegion
                    .Copy
                    End With
                End With
            End With

        Case Else
            
        
    End Select

    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ���� �ٿ��ֱ� : ���� ��ο� ���ϸ� ���� ���ָ� ��.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function pasteFile(tarShtName As Variant, tarRange As Variant)
    
    ActiveWorkbook.Worksheets(tarShtName).Activate
    With ActiveSheet
        .Range(tarRange).Select
        .Paste
    End With

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ���� �ٿ��ֱ� : area 1,2,3,���ῡ ���� �������� �ٸ���
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function pasteAreaFile(tarShtName As Variant, tarRange As Variant, area As Integer)
    
'    ActiveWorkbook.Worksheets(tarShtName).Activate
'    With ActiveSheet
'        .Range(tarRange).Select
'        .Paste
'    End With

    Select Case area
        
        Case 1
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                .Range(tarRange).Select
                .Paste
            End With
        Case 2
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                .Range("A65536").End(xlUp).Offset(0, 0).Select
                .Paste
            End With

        Case Else
            
        
    End Select

    
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' �� ������� ��Ʈ���� ������ ����� : ��Ʈ�� ������
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function createSheetsFile(shtNames As Variant, cntOfSheet As Integer)

    Dim i As Integer
    
    Workbooks.Add
    Worksheets.Add Count:=cntOfSheet - ActiveWorkbook.Worksheets.Count

    For i = 1 To cntOfSheet

    With Worksheets(i)
        .Name = shtNames(i - 1)
        .Range("a1").Value = " A " & shtNames(i - 1) & " ���� ����"
        .Range("a1").Font.Bold = True
    End With
    
    Next i
    

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' �հ� ���ϱ� : ���� ��� ������ ���� ���� ���� ����.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function sumRange1(tarShtName As Variant, tarRange1 As Variant, tarRange2 As Variant)

    ActiveWorkbook.Worksheets(tarShtName).Activate

    With ActiveSheet
        mySum = WorksheetFunction.Sum(.Range(.Range(tarRange1), .Range(tarRange2)))
    End With

        Range(tarRange1).Offset(-1, 4).Value = "�� ��"
        Range(tarRange1).Offset(0, 4).Value = Format(mySum, "#,##0")
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���ǿ� ���� ���� ��ĥ�ϱ�
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function conditionColor(tarShtName As Variant, tarRange As Variant)

    ActiveWorkbook.Worksheets(tarShtName).Activate
    ActiveSheet.Range(tarRange).Select
    
    
    Do
        If ActiveCell.Value < 20 Then
            ActiveCell.Interior.ColorIndex = 6
        Else
            ActiveCell.Interior.ColorIndex = 0
        End If
    ActiveCell.Offset(1, 0).Activate
    Loop Until IsEmpty(ActiveCell)

End Function
