Attribute VB_Name = "ComSub"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 기존 파일 열기 : 파일 경로와 파일명만 지정 해주면 됨.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub OpenFile(filePath As Variant, fileName As Variant)

    Workbooks.Open filePath + fileName
    
End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 기존 파일 닫기 : 파일 경로와 파일명만 지정 해주면 됨.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CloseFile(fileName As Variant)

    With Workbooks(fileName)
    .Save
    .Close
    End With
    
End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 만들기 : 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CreateEmptyFile(filePath As Variant, fileName As Variant)
    
    Workbooks.Add
    
    With ActiveWorkbook
        
        .SaveAs filePath + fileName
        .Close
    
    End With
    
End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 복사: 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CopyFile(srcShtName As Variant, srcRange As Variant)
    
    ActiveWorkbook.Worksheets(srcShtName).Range(srcRange).Copy

End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 복사: area 1,2,3,종료에 따라 지정영역 다르게
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CopyAreaFile(srcShtName As Variant, srcRange As Variant, type_code As Integer)
    Const TARGET_COPY as Integer = 1
    Const REGION_COPY as Integer = 2
    
    Select Case type_code
        
        Case TARGET_COPY
            With ActiveWorkbook
                With .Worksheets(srcShtName)
                    With .Range(.Range(srcRange), .Range(srcRange).End(xlDown).End(xlToRight).End(xlToRight).End(xlToRight))
                    .Copy
                    End With
                End With
            End With

        Case REGION_COPY
            With ActiveWorkbook
                With .Worksheets(srcShtName)
                    With .Range(srcRange).CurrentRegion
                    .Copy
                    End With
                End With
            End With

        Case Else
            
        
    End Select

    
End Sub


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 복사 파일 붙여넣기 : 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub PasteFile(tarShtName As Variant, tarRange As Variant)
    
    ActiveWorkbook.Worksheets(tarShtName).Activate
    With ActiveSheet
        .Range(tarRange).Select
        .Paste
    End With

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 복사 파일 붙여넣기 : area 1,2,3,종료에 따라 지정영역 다르게
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub PasteAreaFile(tarShtName As Variant, tarRange As Variant, type_code As Integer)
    
    Const CUR_POS_PASTE as Integer = 1
    Const NEXT_POS_PASTE as Integer = 2
    
    Select Case type_code
        
        Case CUR_POS_PASTE
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                .Range(tarRange).Select
                .Paste
            End With

        Case NEXT_POS_PASTE
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                .Range("A65536").End(xlUp).Offset(0, 0).Select
                .Paste
            End With
        
        Case Else
        
    End Select

    
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 각 지사명이 시트명인 새파일 만들기 : 시트명 여러개
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub CreateSheetsFile(shtNames As Variant, cntOfSheet As Integer)

    Dim i As Integer
    
    Workbooks.Add
    Worksheets.Add Count:=cntOfSheet - ActiveWorkbook.Worksheets.Count

    For i = 1 To cntOfSheet

    With Worksheets(i)
        .Name = shtNames(i - 1)
        .Range("a1").Value = " A " & shtNames(i - 1) & " 매출 실적"
        .Range("a1").Font.Bold = True
    End With
    
    Next i
    

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 컬럼사이즈 자동조정 : 
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub AutoColumFitSize()

    '--------------------------------------------
    ' 2.1 컬럼 넓이 조정
    '--------------------------------------------
    Columns.ColumnWidth = 30
    
    '--------------------------------------------
    ' 3. 셀 속성 변경 : 폰트 설정
    '--------------------------------------------
    Cells.Select
    With Selection.Font
        .Name = "Arial"
        .Size = 10
        .ColorIndex = xlAutomatic
        .Bold = False
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With

    '--------------------------------------------
    ' 4. 셀 속성 변경 : 색상 지우기
    '--------------------------------------------
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    '--------------------------------------------
    ' 5. 컬럼 및 로 자동 위치정렬
    '--------------------------------------------
    With Cells
        .Select
        .EntireColumn.AutoFit
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        .EntireRow.AutoFit
    End With
    
    '--------------------------------------------
    ' 6. 보더 값 다 지우기
    '--------------------------------------------
    With Selection
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With

End Sub

----------------------------------------------------------------------------------------------------------------------------------------------------------
' 컬럼삭제 : 
' type_code 1 :  왼쪽으로 삭제
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub DeleteCell(tarShtName As Variant, tarRange As Variant, type_code as Variant)

    'Range(tarRange).Select
    'Selection.Delete Shift:=xlToLeft
    
    Const DEL_COLUM as Integer = 1

    Select Case type_code
        
        Case DEL_COLUM
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet.Range(tarRange)
                .Select
                
            End With
            Selection.Delete Shift:=xlToLeft
            
        Case Else
        
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 합계 구하기 : 
' type_code 1 : 시작 행과 마지막 행을 합계값을 쓴다.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub GetSumValue(tarShtName As Variant, scrRange1 As Variant, scrRange2 As Variant, type_code as Variant)

    Const SRC_Range_x2 as Integer = 1
    Const SUM_TITLE as String = "합 계"
    Const TAR_Range_ROW as Integer = -1
    Const TAR_Range_COLUM as Integer = 4

    
    Select Case type_code
        
        Case SRC_Range_x2
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                mySum = WorksheetFunction.Sum(.Range(.Range(tarRange1), .Range(tarRange2)))
            End With
            Range(tarRange1).Offset(TAR_Range_ROW, TAR_Range_COLUM).Value = SUM_TITLE
            Range(tarRange1).Offset(0, TAR_Range_COLUM).Value = Format(mySum, "#,##0")
    
        Case Else
        
    End Select
        
End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 조건에 따른 셀에 색칠하기 : 
' type_code 1 : 조건에 따라 컬러 색칠하기 
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub ConditionFillColor(tarShtName As Variant, tarRange As Variant, type_code as Variant)

    Const U_TYPE as Integer = 1
    
    Const CON_VALUE as Integer = 20
    Const TRUE_COLOR as Integer = 6
    Const FALSE_COLOR as Integer = 0
    
    Select Case type_code
        
        Case U_TYPE

            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
            
                .Range(tarRange).Select
                
            End With
            Do
                If ActiveCell.Value < CON_VALUE Then
                    ActiveCell.Interior.ColorIndex = TRUE_COLOR
                Else
                    ActiveCell.Interior.ColorIndex = FALSE_COLOR
                End If
                ActiveCell.Offset(1, 0).Activate

            Loop Until IsEmpty(ActiveCell)
                
        Case Else
        
    End Select



End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 날짜값 디스플레이 변경
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub ChangeProperty(tarShtName As Variant, tarRange As Variant, type_code as Variant)
    
    Const U_TYPE_DATE as Integer = 1
    Const DATE_FORMAT as String = "yyyy-mm-dd"
    
    Select Case type_code
        
        Case U_TYPE_DATE
            ActiveWorkbook.Worksheets(tarShtName).Activate
            With ActiveSheet
                .Range(tarRange).Select
                .NumberFormatLocal = DATE_FORMAT
            End With
    
        Case Else
        
    End Select

End Sub

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' fill 카피 : 
' type_code 1 : 조건에 따라 컬러 색칠하기 
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Sub FillCopy(srcShtName As Variant, srcRange As Variant, tarShtName As Variant, tarRange As Variant, type_code as Variant)

    Const U_TYPE as Integer = 1
    
    Const CON_VALUE as Integer = 20
    Const TRUE_COLOR as Integer = 6
    Const FALSE_COLOR as Integer = 0
    
    Select Case type_code
        
        Case U_TYPE

            ActiveWorkbook.Worksheets(srcShtName).Activate
            With ActiveSheet
            
                .Range(srcRange).Select
                
            End With
            Do
                If ActiveCell.Value < CON_VALUE Then
                    ActiveCell.Interior.ColorIndex = TRUE_COLOR
                Else
                    ActiveCell.Interior.ColorIndex = FALSE_COLOR
                End If
                ActiveCell.Offset(1, 0).Activate

            Loop Until IsEmpty(ActiveCell)
                
        Case Else
        
    End Select



End Sub
    