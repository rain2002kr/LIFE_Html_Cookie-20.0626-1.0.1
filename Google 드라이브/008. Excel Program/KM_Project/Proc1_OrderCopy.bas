Attribute VB_Name = "Proc1_OrderCopy"
Sub Proc_OrderCopy()
        
    Dim myDir As String, myFile As String
    myFile = "PO-2022012KM"
    
    ' 반복 실행
    Call FileCopysA(myFile)
    Call OrderRaw2
    Call OrderRaw3
    Call OrderRaw4
    ' 반복 끝 실행
    
    ' orderRaw -> orderMake
    Call OrderRaw5
    
    ' orderMake -> SUM
    Call OrderRaw6
    
    ' SUM P, Month, NO 추가
    Call OrderRaw7
    
End Sub

Sub ExtractFiles()

    Dim myDir As String, myFile As String
    Dim i As Integer

    myDir = "C:\Users\rain2\Google 드라이브\TEST\KM\"
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
'        .Offset(0, 0).Value = "파일이름"
'        .Offset(0, 1).Value = "크기"
'        .Offset(0, 2).Value = "날짜/시간"
'        .CurrentRegion.AutoFormat
'    End With


End Sub

Sub FileCopys()
    '--------------------------------------------
    ' 0. 변수 지정
    '--------------------------------------------
    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Const TARGET_COPY As Integer = 1
    Const REGION_COPY As Integer = 2

    '--------------------------------------------
    ' 1. 소스 파일 열기
    '--------------------------------------------
    srcFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    SrcFileName = "PO-2022012KM"
    
    
    srcShtName = "발주서"
    srcRange = "A29"
    
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 2. 소스 파일 복사
    '--------------------------------------------
    Call CopyAreaFile(srcShtName, srcRange, TARGET_COPY)
    
    '--------------------------------------------
    ' 3. 대상 파일 열기 혹은 만들기
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "A2"
    
    Call OpenFile(tarFilePath, tarFileName)

    '--------------------------------------------
    ' 4. 대상 파일 붙여 넣기
    '--------------------------------------------
    Call PasteFile(tarShtName, tarRange)
    
    
    '--------------------------------------------
    ' 5. 소스 파일 열기
    '--------------------------------------------
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 6. 소스 파일 복사
    '--------------------------------------------
    srcRange = "C9:C15"
    Call CopyFile(srcShtName, srcRange)
    
    '--------------------------------------------
    ' 7. 대상 파일 열기 혹은 만들기
    '--------------------------------------------
    Call OpenFile(tarFilePath, tarFileName)
    '--------------------------------------------
    ' 8. 대상 파일 붙여 넣기
    '--------------------------------------------
    tarRange = "N2"
    Call PasteFile(tarShtName, tarRange)
    
    '--------------------------------------------
    ' 2. 컬럼 삭제
    '--------------------------------------------
    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. 날짜값 디스플레이 변경
    '--------------------------------------------
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)

    
        
    '--------------------------------------------
    ' 100. 소스 파일 닫기
    '--------------------------------------------
    Call CloseFile(SrcFileName)
    Call CloseFile(tarFileName)

End Sub

Sub FileCopysA(SrcFileNameIn As Variant)
    '--------------------------------------------
    ' 0. 변수 지정
    '--------------------------------------------
    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Const TARGET_COPY As Integer = 1
    Const REGION_COPY As Integer = 2

    '--------------------------------------------
    ' 1. 소스 파일 열기
    '--------------------------------------------
    srcFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    If (SrcFileNameIn = "") Then
        SrcFileName = "PO-2022012KM"
        Else
    
        SrcFileName = SrcFileNameIn
    End If
    
    
    srcShtName = "발주서"
    srcRange = "A29"
    
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 2. 소스 파일 복사
    '--------------------------------------------
    Call CopyAreaFile(srcShtName, srcRange, TARGET_COPY)
    
    '--------------------------------------------
    ' 3. 대상 파일 열기 혹은 만들기
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "A2"
    
    Call OpenFile(tarFilePath, tarFileName)

    '--------------------------------------------
    ' 4. 대상 파일 붙여 넣기
    '--------------------------------------------
    Call PasteFile(tarShtName, tarRange)
    
    
    '--------------------------------------------
    ' 5. 소스 파일 열기
    '--------------------------------------------
    Call OpenFile(srcFilePath, SrcFileName)
    
    '--------------------------------------------
    ' 6. 소스 파일 복사
    '--------------------------------------------
    srcRange = "C9:C15"
    Call CopyFile(srcShtName, srcRange)
    
    '--------------------------------------------
    ' 7. 대상 파일 열기 혹은 만들기
    '--------------------------------------------
    Call OpenFile(tarFilePath, tarFileName)
    '--------------------------------------------
    ' 8. 대상 파일 붙여 넣기
    '--------------------------------------------
    tarRange = "N2"
    Call PasteFile(tarShtName, tarRange)
    
    '--------------------------------------------
    ' 2. 컬럼 삭제
    '--------------------------------------------
    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. 날짜값 디스플레이 변경
    '--------------------------------------------
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)

    
        
    '--------------------------------------------
    ' 100. 소스 파일 닫기
    '--------------------------------------------
    Call CloseFile(SrcFileName)
    Call CloseFile(tarFileName)

End Sub

'--------------------------------------------
' 0. orderRaw2 설명
'
'--------------------------------------------

Sub OrderRaw2()
    '--------------------------------------------
    ' 0. 변수 지정
    '--------------------------------------------
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Dim copyCount As Variant
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    
    
    Call OpenFile(tarFilePath, tarFileName)
    copyCount = getColumCount(tarShtName, "A2")
    
    
    'Worksheets("OrderRaw").Range("A1").Value = copyCount
    
    
    '--------------------------------------------
    ' 1. 컬럼추가 지정
    '--------------------------------------------
    tarShtName = "OrderRaw"
    tarRange = "A"
    
    For i = 1 To 8
        Call InsertCell(tarShtName, tarRange, 1)
    Next i
    '--------------------------------------------
    ' 2. 값 세팅 및 복사
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
    ' 3. 컬럼삭제
    '--------------------------------------------
    tarRange = "M:M,P:Q"
    
    Call DeleteCell(tarShtName, tarRange, 1)
    
    '--------------------------------------------
    ' 6. 소스 파일 복사
    '--------------------------------------------
    'srcRange = "A1"
    'tarShtName = "Template"

    'Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. 대상 파일 붙여 넣기
    '--------------------------------------------
    'tarShtName = "orderRaw"

    'Call PasteAreaFile(tarShtName, srcRange, 1)

    Call AutoColumFitSize
    

End Sub


Sub OrderRaw3()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    
    Call OpenFile(tarFilePath, tarFileName)
    
    
    srcShtName = "Data"
    tarShtName = "OrderRaw"
    srcRange = "B2:K131"
    tarRange = "J2"

    retVal = searchBuDataWithExcepection(srcShtName, tarShtName, srcRange, tarRange)
    
    
    '--------------------------------------------
    ' 3. 컬럼삭제
    '--------------------------------------------
    tarRange = "T:V"
    
    Call DeleteCell(tarShtName, tarRange, 1)

    '--------------------------------------------
    ' 6. 소스 파일 복사
    '--------------------------------------------
    'srcRange = "A1"
    'tarShtName = "Template"

    'Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. 대상 파일 붙여 넣기
    '--------------------------------------------
    'tarShtName = "orderRaw"

    'Call PasteAreaFile(tarShtName, srcRange, 1)

    Call AutoColumFitSize





End Sub


Sub OrderRaw4()

    Dim srcFilePath, SrcFileName, srcShtName, srcRange As String
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String
    Dim copyCount As Variant
    
    
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
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
    
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    
    Call OpenFile(tarFilePath, tarFileName)


    '--------------------------------------------
    ' 6. 소스 파일 복사
    '--------------------------------------------
    tarShtName = "OrderRaw"
    srcRange = "A2"
    Call CopyAreaFile(tarShtName, srcRange, 2)
    '--------------------------------------------
    ' 7. 대상 파일 붙여 넣기
    '--------------------------------------------
    tarShtName = "orderMake"
    

    Call PasteAreaFile(tarShtName, srcRange, 2)
    
    Call AutoColumFitSize
    
    
    '--------------------------------------------
    ' 8. orderRaw 시트 내용 삭제
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
    
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    '--------------------------------------------
    ' 6. 소스 파일 복사
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
    
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    
    '--------------------------------------------
    ' 8. NO 추가, P , MONTH 추가
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
    ' 0. 변수 지정
    '--------------------------------------------
    Dim tarFilePath, tarFileName, tarShtName, tarRange As String

    '--------------------------------------------
    ' 1. 대상 파일 열기
    '--------------------------------------------
    tarFilePath = "C:\Users\rain2\Google 드라이브\TEST\KM\"
    tarFileName = "KM_SUM.xlsx"
    tarShtName = "OrderRaw"
    tarRange = "E:F,H:J"
    
    Call OpenFile(srcFilePath, SrcFileName)
    '--------------------------------------------
    ' 2. 컬럼 삭제
    '--------------------------------------------
    ' Range(tarRange).Select
    ' Selection.Delete Shift:=xlToLeft

    tarRange = "E:F,H:J"
    Const DEL_COLUM As Integer = 1
    
    Call DeleteCell(tarShtName, tarRange, DEL_COLUM)

    Call AutoColumFitSize
    
    '--------------------------------------------
    ' 7. 날짜값 디스플레이 변경
    '--------------------------------------------
    ' With ActiveSheet.Range("I2:I10")
    '         .Select
    '         .NumberFormatLocal = "yyyy-mm-dd"
    ' End With
    
    tarRange = "I2:I10"
    Call ChangeProperty(tarShtName, tarRange, 1)
    
    
    '--------------------------------------------
    ' 8. 파일 닫기
    '--------------------------------------------
   ' retVal = closeFile(tarFileName)


End Sub





