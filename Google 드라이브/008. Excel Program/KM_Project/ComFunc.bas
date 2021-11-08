Attribute VB_Name = "ComFunc"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���� ���� �÷� ���� ���� : ���ϰ� �÷���
'----------------------------------------------------------------------------------------------------------------------------------------------------------

Function getColumCount(tarShtName As Variant, tarRange As Variant) As Integer
    
    ActiveWorkbook.Worksheets(tarShtName).Range(tarRange).Select
    Range(Selection, Selection.End(xlDown)).Select
    getColumCount = Selection.count


End Function



'----------------------------------------------------------------------------------------------------------------------------------------------------------
' ���ǿ� ���� ������ ã�ƿ��� 3 ���� ���� �ڵ�
' 1. excel open
' 2. target sheet keyword ��������
' 32. target sheet keyword ��������
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function searchBuDataWithExcepection(srcShtName As Variant, tarShtName As Variant, srcRange As Variant, tarRange As Variant)
    '---------------------------------------------------------------------------------
    ' ��Ʈ ����
    '---------------------------------------------------------------------------------
    Dim searchRng As Range
    Dim keyword As String
    Dim i As Integer
    i = 0
    
    
    Set myRng = ActiveSheet.Range(srcRange).Find(What:=countryName, Lookat:=xlWhole, LookIn:=xlValues)
    Do
        ActiveWorkbook.Worksheets(tarShtName).Activate
        ActiveSheet.Range(tarRange).Offset(i, 0).Select
        keyword = ActiveCell.Value

        ActiveWorkbook.Worksheets(srcShtName).Activate
        ActiveSheet.Range(srcRange).Select
        Set searchRng = ActiveSheet.Range(srcRange).Find(What:=keyword, Lookat:=xlWhole, LookIn:=xlValues)

        ActiveWorkbook.Worksheets(tarShtName).Activate
        ActiveSheet.Range(tarRange).Select

    '---------------------------------------------------------------------------------
    ' �� �Ǵ�
    '---------------------------------------------------------------------------------
        If searchRng Is Nothing Then
            ActiveSheet.Range(tarRange).Offset(i, 5) = "�� ����"
        Else
    
    '---------------------------------------------------------------------------------
    ' �� ���� �� �ٿ��ֱ�
    '---------------------------------------------------------------------------------
            searchRng.Offset(0, 1).Resize(1, 10).Copy Destination:=ActiveSheet.Range(tarRange).Offset(i, 5)
        End If

        i = i + 1
        ActiveWorkbook.Worksheets(tarShtName).Activate
        ActiveSheet.Range(tarRange).Offset(i, 0).Select
        
        
    Loop Until IsEmpty(ActiveCell)


End Function
