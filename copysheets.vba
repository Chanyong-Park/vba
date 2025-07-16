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
    
    Dim rowIndex As Integer
    Dim classified As String
    Dim uri As String
    Dim desc As String

    Set apiSheet = Worksheets("API목록")
    Set selectedRange = Selection
    rowIndex = selectedRange.Row
    
    If selectedRange.Count = 1 Then
        sheetCount = selectedRange.Count

        ReDim sheetNames(1, 1) As Variant
        sheetNames(1, 1) = selectedRange.Value
    Else
        'Cell 복수개 선택
        sheetNames = selectedRange.Value
        sheetCount = UBound(sheetNames) - LBound(sheetNames) + 1
    End If

    For x = 1 To sheetCount
        
        classified = apiSheet.Cells(rowIndex + x - 1, 1).Value
        uri = apiSheet.Cells(rowIndex + x - 1, 2).Value
        desc = apiSheet.Cells(rowIndex + x - 1, 5).Value
    
        'MsgBox classified & " " & uri & " " & desc

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
        
        ActiveSheet.Cells(4, 2).Value = classified
        ActiveSheet.Cells(4, 5).Value = uri
        ActiveSheet.Cells(4, 12).Value = desc
        
    Next x
End Sub
