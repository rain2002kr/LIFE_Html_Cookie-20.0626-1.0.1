Attribute VB_Name = "ComFunc"
'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 선택 영역 컬럼 숫자 세기 : 리턴값 컬럼수
'----------------------------------------------------------------------------------------------------------------------------------------------------------

Function getColumCount(tarShtName As Variant, tarRange As Variant) As Integer
    
    ActiveWorkbook.Worksheets(tarShtName).Range(tarRange).Select
    Range(Selection, Selection.End(xlDown)).Select
    getColumCount = Selection.count


End Function



'----------------------------------------------------------------------------------------------------------------------------------------------------------
' 조건에 따른 데이터 찾아오기 3 에러 방지 코드
' 1. excel open
' 2. target sheet keyword 가져오기
' 32. target sheet keyword 가져오기
'----------------------------------------------------------------------------------------------------------------------------------------------------------
Function searchBuDataWithExcepection(srcShtName As Variant, tarShtName As Variant, srcRange As Variant, tarRange As Variant)
    '---------------------------------------------------------------------------------
    ' 시트 선택
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
    ' 값 판단
    '---------------------------------------------------------------------------------
        If searchRng Is Nothing Then
            ActiveSheet.Range(tarRange).Offset(i, 5) = "값 없음"
        Else
    
    '---------------------------------------------------------------------------------
    ' 값 복사 및 붙여넣기
    '---------------------------------------------------------------------------------
            searchRng.Offset(0, 1).Resize(1, 10).Copy Destination:=ActiveSheet.Range(tarRange).Offset(i, 5)
        End If

        i = i + 1
        ActiveWorkbook.Worksheets(tarShtName).Activate
        ActiveSheet.Range(tarRange).Offset(i, 0).Select
        
        
    Loop Until IsEmpty(ActiveCell)


End Function
