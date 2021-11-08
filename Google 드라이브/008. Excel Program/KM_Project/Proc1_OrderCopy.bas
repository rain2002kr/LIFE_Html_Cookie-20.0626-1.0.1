Attribute VB_Name = "Proc1_OrderCopy"
Sub Proc_OrderCopy()
        
    Dim myDir As String, myFile As String
    myFile = "PO-2022012KM"
    
    ' �ݺ� ����
    Call FileCopysA(myFile)
    Call OrderRaw2
    Call OrderRaw3
    Call OrderRaw4
    ' �ݺ� �� ����
    
    ' orderRaw -> orderMake
    Call OrderRaw5
    
    ' orderMake -> SUM
    Call OrderRaw6
    
    ' SUM P, Month, NO �߰�
    Call OrderRaw7
    
End Sub

Sub ExtractFiles()

    Dim myDir As String, myFile As String
    Dim i As Integer

    myDir = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    myFile = Dir(myDir & "P*.xlsx")

    i = 1

    Do While myFile <> ""
        With ActiveSheet.Range("A1")
            .Offset(i, 0) = myFile
            .Offset(i, 1) = FileLen(myDir & myFile)
            .Offset(i, 2) = FileDateTime(myDir & myFile)
        End With
        
      Call FileCopysA(myFile)
'      Call OrderRaw2
'      Call OrderRaw3
'      Call OrderRaw4
'      Call OrderRaw5
'      Call OrderRaw6
'      Call OrderRaw7
      
    myFile = Dir
        i = i + 1
    Loop
    

'    With Range("A1")
'        .Offset(0, 0).Value = "�����̸�"
'        .Offset(0, 1).Value = "ũ��"
'        .Offset(0, 2).Value = "��¥/�ð�"
'        .CurrentRegion.AutoFormat
'    End With


End Sub

Sub FileCopys()
    '--------------------------------------------
    ' 0. ���� ����
    '--------------------------------------------
    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Const TARGET_COPY As Integer = 1
    Const REGION_COPY As Integer = 2

    '--------------------------------------------
    ' 1. �ҽ� ���� ����
    '--------------------------------------------
    srcFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    SrcFileName = "PO-2022012KM"
    
    
    srcShtName = "���ּ�"
    srcRange = "A29"
    
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 2. �ҽ� ���� ����
    '--------------------------------------------
    Call CopyAreaFile(srcShtName, srcRange, TARGET_COPY)
    
    '--------------------------------------------
    ' 3. ��� ���� ���� Ȥ�� �����
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "A2"
    
    Call OpenFile(tarFilePath, tarFileName)

    '--------------------------------------------
    ' 4. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    Call PasteFile(tarShtName, tarRange)
    
    
    '--------------------------------------------
    ' 5. �ҽ� ���� ����
    '--------------------------------------------
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    srcRange = "C9:C15"
    Call CopyFile(srcShtName, srcRange)
    
    '--------------------------------------------
    ' 7. ��� ���� ���� Ȥ�� �����
    '--------------------------------------------
    Call OpenFile(tarFilePath, tarFileName)
    '--------------------------------------------
    ' 8. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    tarRange = "N2"
    Call PasteFile(tarShtName, tarRange)
    
    '--------------------------------------------
    ' 2. �÷� ����
    '--------------------------------------------
    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. ��¥�� ���÷��� ����
    '--------------------------------------------
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)

    
        
    '--------------------------------------------
    ' 100. �ҽ� ���� �ݱ�
    '--------------------------------------------
    Call CloseFile(SrcFileName)
    Call CloseFile(tarFileName)

End Sub

Sub FileCopysA(SrcFileNameIn As Variant)
    '--------------------------------------------
    ' 0. ���� ����
    '--------------------------------------------
    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Const TARGET_COPY As Integer = 1
    Const REGION_COPY As Integer = 2

    '--------------------------------------------
    ' 1. �ҽ� ���� ����
    '--------------------------------------------
    srcFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    If (SrcFileNameIn = "") Then
        SrcFileName = "PO-2022012KM"
        Else
    
        SrcFileName = SrcFileNameIn
    End If
    
    
    srcShtName = "���ּ�"
    srcRange = "A29"
    
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 2. �ҽ� ���� ����
    '--------------------------------------------
    Call CopyAreaFile(srcShtName, srcRange, TARGET_COPY)
    
    '--------------------------------------------
    ' 3. ��� ���� ���� Ȥ�� �����
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "A2"
    
    Call OpenFile(tarFilePath, tarFileName)

    '--------------------------------------------
    ' 4. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    Call PasteFile(tarShtName, tarRange)
    
    
    '--------------------------------------------
    ' 5. �ҽ� ���� ����
    '--------------------------------------------
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    srcRange = "C9:C15"
    Call CopyFile(srcShtName, srcRange)
    
    '--------------------------------------------
    ' 7. ��� ���� ���� Ȥ�� �����
    '--------------------------------------------
    Call OpenFile(tarFilePath, tarFileName)
    '--------------------------------------------
    ' 8. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    tarRange = "N2"
    Call PasteFile(tarShtName, tarRange)
    
    '--------------------------------------------
    ' 2. �÷� ����
    '--------------------------------------------
    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. ��¥�� ���÷��� ����
    '--------------------------------------------
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)

    
        
    '--------------------------------------------
    ' 100. �ҽ� ���� �ݱ�
    '--------------------------------------------
    Call CloseFile(SrcFileName)
    Call CloseFile(tarFileName)

End Sub

'--------------------------------------------
' 0. orderRaw2 ����
'
'--------------------------------------------

Sub OrderRaw2()
    '--------------------------------------------
    ' 0. ���� ����
    '--------------------------------------------
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Dim copyCount As Variant
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    
    
    Call OpenFile(tarFilePath, tarFileName)
    copyCount = getColumCount(tarShtName, "A2")
    
    
    'Worksheets("OrderRaw").Range("A1").Value = copyCount
    
    
    '--------------------------------------------
    ' 1. �÷��߰� ����
    '--------------------------------------------
    tarShtName = "OrderRaw"
    tarRange = "A"
    
    For i = 1 To 8
        Call InsertCell(tarShtName, tarRange, 1)
    Next i
    '--------------------------------------------
    ' 2. �� ���� �� ����
    '--------------------------------------------
    
    For i = 0 To copyCount - 1
        Range("A2").Offset(i, 0).Value = Range("Q4").Value
        Range("B2").Offset(i, 0).Value = Range("Q8").Value
        Range("C2").Offset(i, 0).Value = Range("Q2").Value
        Range("D2").Offset(i, 0).Value = "KM Engineering Co Ltd"
        Range("E2").Offset(i, 0).Value = Range("P2").Value
        Range("F2").Offset(i, 0).Value = Range("Q5").Value
        Range("G2").Offset(i, 0).Value = Range("Q6").Value
        Range("H2").Offset(i, 0).Value = "STI"
        
    Next i
    
    
    '--------------------------------------------
    ' 3. �÷�����
    '--------------------------------------------
    tarRange = "M:M,P:Q"
    
    Call DeleteCell(tarShtName, tarRange, 1)
    
    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    'srcRange = "A1"
    'tarShtName = "Template"

    'Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    'tarShtName = "orderRaw"

    'Call PasteAreaFile(tarShtName, srcRange, 1)

    Call AutoColumFitSize
    

End Sub


Sub OrderRaw3()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    
    Call OpenFile(tarFilePath, tarFileName)
    
    
    srcShtName = "Data"
    tarShtName = "OrderRaw"
    srcRange = "B2:K131"
    tarRange = "J2"

    retVal = searchBuDataWithExcepection(srcShtName, tarShtName, srcRange, tarRange)
    
    
    '--------------------------------------------
    ' 3. �÷�����
    '--------------------------------------------
    tarRange = "T:V"
    
    Call DeleteCell(tarShtName, tarRange, 1)

    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    'srcRange = "A1"
    'tarShtName = "Template"

    'Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    'tarShtName = "orderRaw"

    'Call PasteAreaFile(tarShtName, srcRange, 1)

    Call AutoColumFitSize





End Sub


Sub OrderRaw4()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Dim copyCount As Variant
    
    
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    tarShtName = "OrderRaw"
    
    Call OpenFile(tarFilePath, tarFileName)
    
    copyCount = getColumCount(tarShtName, "A2")
    
    srcRange = "N2:O2"
    
    
    Dim FA, PA, MC, TOTAL As Integer
    
    
    For i = 0 To copyCount

        If (Range("O2").Offset(i, 0).Value = "FA") Then
            FA = FA + Range("N2").Offset(i, 0).Value
            
        ElseIf (Range("O2").Offset(i, 0).Value = "PA") Then
            PA = PA + Range("N2").Offset(i, 0).Value
        
        Else
            MC = MC + Range("N2").Offset(i, 0).Value

        End If


    Next i
    

    Range("U2").Value = FA
    Range("V2").Value = PA
    Range("W2").Value = MC
    Range("X2").Value = FA + PA + MC
    
    
    Columns("U:X").Select
    Selection.Cut
    Columns("I:I").Select
    Selection.Insert Shift:=xlToRight
    
    Columns("I:L").Select
    Selection.NumberFormatLocal = "\#,##0_);(\#,##0)"
    
    
    
    'Call AutoColumFitSize

End Sub

Sub OrderRaw5()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    
    Call OpenFile(tarFilePath, tarFileName)


    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    tarShtName = "OrderRaw"
    srcRange = "A2"
    Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. ��� ���� �ٿ� �ֱ�
    '--------------------------------------------
    tarShtName = "orderMake"
    

    Call PasteAreaFile(tarShtName, srcRange, 2)
    
    Call AutoColumFitSize
    
    
    '--------------------------------------------
    ' 8. orderRaw ��Ʈ ���� ����
    '--------------------------------------------
    tarShtName = "OrderRaw"
    ActiveWorkbook.Worksheets(tarShtName).Activate
    
    Cells.Select
    Range("W6").Activate
    Selection.Delete Shift:=xlUp
    
    
    



End Sub



Sub OrderRaw6()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    '--------------------------------------------
    ' 6. �ҽ� ���� ����
    '--------------------------------------------
    srcShtName = "orderMake"
    tarShtName = "SUM"
    
    srcRange = "A1"
    
    Dim count As Variant
    
    Call OpenFile(tarFilePath, tarFileName)
    
    ActiveWorkbook.Worksheets("orderMake").Activate
    ActiveSheet.Range("A1").Select
      
    
        Range(Selection, Selection.End(xlDown)).Select
    count = Selection.count
    
    
    
    Dim cnt As Integer
    cnt = 1
    
    For i = 0 To count
        If Worksheets(srcShtName).Range("I1").Offset(i, 0) <> "" Then
            srcRange = "A1:L1"
            tarRange = "A1:L1"
            
            Worksheets(srcShtName).Range(srcRange).Offset(i, 0).Copy Destination:=Worksheets(tarShtName).Range(tarRange).Offset(cnt, 3)
            
            ' ADD Number COUNT
            Worksheets(tarShtName).Range("C2").Offset(cnt - 1, 0) = cnt
            
            cnt = cnt + 1
        
        End If

    Next i

End Sub


Sub OrderRaw7()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    '--------------------------------------------
    ' 8. NO �߰�, P , MONTH �߰�
    '--------------------------------------------
    srcShtName = "SUM"
    srcRange = "A1"
    
    Dim count As Variant
    Dim monthValue, pValue As Variant
    
    Call OpenFile(tarFilePath, tarFileName)
    
    ActiveWorkbook.Worksheets("SUM").Activate
    
    
    
     Dim cntLoop As Integer
     
     Do While Worksheets(srcShtName).Range("D2").Offset(cntLoop, 0) <> ""
            
            monthValue = month(Worksheets(srcShtName).Range("D2").Offset(cntLoop, 0).Value)
            
            
            If (monthValue > 9) Then
            
                pValue = monthValue - 9
                
            Else
                pValue = monthValue + 3
                        
            End If
            
            Worksheets(srcShtName).Range("B2").Offset(cntLoop, 0) = "P" & pValue
            Worksheets(srcShtName).Range("A2").Offset(cntLoop, 0) = monthValue
        
        cntLoop = cntLoop + 1

    Loop
    
    
   
    

End Sub






Sub FileModifyDelete()
    '--------------------------------------------
    ' 0. ���� ����
    '--------------------------------------------
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String

    '--------------------------------------------
    ' 1. ��� ���� ����
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google ����̺�\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "E:F,H:J"
    
    Call OpenFile(srcFilePath, SrcFileName)
    '--------------------------------------------
    ' 2. �÷� ����
    '--------------------------------------------
    ' Range(tarRange).Select
    ' Selection.Delete Shift:=xlToLeft

    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. ��¥�� ���÷��� ����
    '--------------------------------------------
    ' With ActiveSheet.Range("I2:I10")
    '         .Select
    '         .NumberFormatLocal = "yyyy-mm-dd"
    ' End With
    
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)
    
    
    '--------------------------------------------
    ' 8. ���� �ݱ�
    '--------------------------------------------
   ' retVal = closeFile(tarFileName)


End Sub





