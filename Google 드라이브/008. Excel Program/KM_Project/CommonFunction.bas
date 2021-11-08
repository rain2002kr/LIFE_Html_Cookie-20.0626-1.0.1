Attribute VB_Name = "CommonFunction"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 기존 파일 열기 : 파일 경로와 파일명만 지정 해주면 됨.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function openFile(filePath As Variant, fileName As Variant)

    Workbooks.Open filePath + fileName
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 기존 파일 닫기 : 파일 경로와 파일명만 지정 해주면 됨.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function closeFile(fileName As Variant)

    With Workbooks(fileName)
    .Save
    .Close
    
    End With
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 만들기 : 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function createEmptyFile(filePath As Variant, fileName As Variant)
    
    Workbooks.Add
    
    With ActiveWorkbook
        
        .SaveAs filePath + fileName
        .Close
    
    End With
    
End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 복사: 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function copyFile(srcShtName As Variant, srcRange As Variant)
    
    ActiveWorkbook.Worksheets(srcShtName).Range(srcRange).Copy

End Function


'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 파일 복사: area 1,2,3,종료에 따라 지정영역 다르게
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
' 복사 파일 붙여넣기 : 파일 경로와 파일명만 지정 해주면 됨.'.Worksheets(shtName).Range(rangeArea).Value = "Create New File"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function pasteFile(tarShtName As Variant, tarRange As Variant)
    
    ActiveWorkbook.Worksheets(tarShtName).Activate
    With ActiveSheet
        .Range(tarRange).Select
        .Paste
    End With

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 복사 파일 붙여넣기 : area 1,2,3,종료에 따라 지정영역 다르게
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
' 각 지사명이 시트명인 새파일 만들기 : 시트명 여러개
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function createSheetsFile(shtNames As Variant, cntOfSheet As Integer)

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
    

End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 합계 구하기 : 시작 행과 마지막 행을 더한 값을 쓴다.
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function sumRange1(tarShtName As Variant, tarRange1 As Variant, tarRange2 As Variant)

    ActiveWorkbook.Worksheets(tarShtName).Activate

    With ActiveSheet
        mySum = WorksheetFunction.Sum(.Range(.Range(tarRange1), .Range(tarRange2)))
    End With

        Range(tarRange1).Offset(-1, 4).Value = "합 계"
        Range(tarRange1).Offset(0, 4).Value = Format(mySum, "#,##0")
        
End Function

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 조건에 따른 셀에 색칠하기
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
