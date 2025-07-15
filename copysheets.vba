Sub sheet_creation_click()
    Dim selectedRange As Range
    Dim sheetNames As Variant
    Dim sheetCount
    
    Dim cellValue As String
    Dim baseName As String
    Dim newName As String
    Dim i As Integer
    Dim ws As Worksheet
    Dim exists As Boolean

    Set selectedRange = Selection
    If selectedRange.Count = 1 Then
        'Cell 1개 선택
        sheetCount = selectedRange.Count

        ReDim sheetNames(1, 1) As Variant
        sheetNames(1, 1) = selectedRange.Value
    Else
        'Cell 복수개 선택
        sheetNames = selectedRange.Value
        sheetCount = UBound(sheetNames) - LBound(sheetNames) + 1
    End If

    For x = 1 To sheetCount
    
        baseName = sheetNames(x, 1)
        newName = baseName
        'MsgBox newName
        i = 1

        ' 이름 중복 방지
        Do
            exists = False
            For Each ws In Worksheets
                If ws.Name = newName Then
                    exists = True
                    Exit For
                End If
            Next ws
            If exists Then
                newName = baseName & "_" & i
                i = i + 1
            End If
        Loop While exists
    
        ' 시트는 생성되나 오류메시지 표시됨
        'Worksheets("template").Copy After:=Worksheets(Worksheets.Count)
        'ActiveSheet.Name = newName
        
        ' 시트는 생성되나 오류메시지 표시됨
        'Sheets("template").Copy After:=Sheets(Sheets.Count - 2)
        'ActiveSheet.Name = newName
        
        Sheets.Add After:=Sheets(Sheets.Count)
        Sheets("template").Cells.Copy Destination:=ActiveSheet.Cells
        ActiveSheet.Name = newName
        '시트 색 지정
        ActiveSheet.Tab.Color = RGB(255, 255, 0)
    Next x
End Sub
