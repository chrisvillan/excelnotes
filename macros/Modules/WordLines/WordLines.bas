Attribute VB_Name = "WordLines"
Sub UseWordLines()
    'add please close all word doc validatioin
    Dim frm As New UserForm_WordLines
    
    With frm
        .checkbox_keeporiginal.Value = True
        .Show vbModeless
    End With
    
End Sub
Private Function GetWordLine(line As Integer, doc As Word.Document) As String

    Dim wordVar As Variant
    wordVar = GetWordVar(line, doc)
    
    GetWordLine = doc.ActiveWindow.Panes(wordVar(1)).Pages(wordVar(2)).Rectangles(wordVar(3)).Lines(wordVar(4)).Range.Text
End Function

Private Sub ReplaceLines(str As String, line As Integer, doc As Word.Document)

    Dim wordVar As Variant
    
    wordVar = GetWordVar(line, doc)
    
    doc.ActiveWindow.Panes(wordVar(1)).Pages(wordVar(2)).Rectangles(wordVar(3)).Lines(wordVar(4)).Range.Text = str & vbNewLine
    wb.ActiveSheet.Cells(wordVar(5), 2).Value = str
End Sub


Private Sub DeleteLines(line As Integer, doc As Word.Document, keepblank As Boolean)

    Dim wordVar As Variant
    
    wordVar = GetWord(line, doc)
    
    If keepblank = True Then
    
        doc.ActiveWindow.Panes(wordVar(1)).Pages(wordVar(2)).Rectangles(wordVar(3)).Lines(wordVar(4)).Range.Delete
    Else
        doc.ActiveWindow.Panes(wordVar(1)).Pages(wordVar(2)).Rectangles(wordVar(3)).Lines(wordVar(4)).Range.Text = vbNewLine
    End If
    
End Sub


Private Function GetWordVar(line As Integer, doc As Word.Document) As Variant
    Dim wordVar As Variant
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    ReDim wordVar(5)
    
    
    For i = 2 To wb.ActiveSheet.Range("C2").End(xlDown).row
        If wb.ActiveSheet.Cells(i, 3).Value = line Then
            wordVar(1) = wb.ActiveSheet.Cells(i, 4).Value
            wordVar(2) = wb.ActiveSheet.Cells(i, 5).Value
            wordVar(3) = wb.ActiveSheet.Cells(i, 6).Value
            wordVar(4) = wb.ActiveSheet.Cells(i, 7).Value
            wordVar(5) = i
        End If
    Next
    
    GetWordVar = wordVar
End Function

Private Sub PrintAllLines(doc As Word.Document)
    Dim line_text As String
    Dim row_excel As Integer
    Dim row As Integer
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    
    row_excel = 2
    Dim count_pane As Integer
    Dim count_page As Integer
    Dim count_rect As Integer
    Dim count_line As Integer
    Dim word_count_line As Integer
    
    
    count_pane = 0
    count_page = 0
    count_rect = 0
    count_line = 0
    word_count_line = 0
    
    'add headers
    wb.ActiveSheet.Range("A1").Value = "Word Row"
    wb.ActiveSheet.Range("A1").Font.Bold = True
    wb.ActiveSheet.Range("B1").Value = "Line Text"
    wb.ActiveSheet.Range("B1").Font.Bold = True
    wb.ActiveSheet.Range("C1").Value = "Word Row"
    wb.ActiveSheet.Range("C1").Font.Bold = True
    wb.ActiveSheet.Range("D1").Value = "Pane"
    wb.ActiveSheet.Range("D1").Font.Bold = True
    wb.ActiveSheet.Range("E1").Value = "Page"
    wb.ActiveSheet.Range("E1").Font.Bold = True
    wb.ActiveSheet.Range("F1").Value = "Rect"
    wb.ActiveSheet.Range("F1").Font.Bold = True
    wb.ActiveSheet.Range("G1").Value = "Line"
    wb.ActiveSheet.Range("G1").Font.Bold = True
    wb.ActiveSheet.Range("H1").Value = "Original Line Text"
    wb.ActiveSheet.Range("H1").Font.Bold = True
    
    
    For word_pane = 1 To doc.ActiveWindow.Panes.Count
        count_pane = count_pane + 1
        For word_page = 1 To doc.ActiveWindow.Panes(word_pane).Pages.Count
            count_page = count_page + 1
            For word_rect = 1 To doc.ActiveWindow.Panes(word_pane).Pages(word_page).Rectangles.Count
                count_rect = count_rect + 1
                For word_line = 1 To doc.ActiveWindow.Panes(word_pane).Pages(word_page).Rectangles(word_rect).Lines.Count
                    count_line = count_line + 1
                    line_text = doc.ActiveWindow.Panes(word_pane).Pages(word_page).Rectangles(word_rect).Lines(word_line).Range.Text
                    word_count_line = word_count_line + 1
                    
                    wb.ActiveSheet.Cells(row_excel, 1).Value = word_count_line
                    wb.ActiveSheet.Cells(row_excel, 2).Value = line_text
                    
                    wb.ActiveSheet.Cells(row_excel, 7).Value = word_line
                    wb.ActiveSheet.Cells(row_excel, 8).Value = line_text
                    
                    wb.ActiveSheet.Cells(row_excel, 4).Value = word_pane
                    wb.ActiveSheet.Cells(row_excel, 3).Value = word_count_line
                    wb.ActiveSheet.Cells(row_excel, 5).Value = word_page
                    wb.ActiveSheet.Cells(row_excel, 6).Value = word_rect
                    
                    
                    row_excel = row_excel + 1
                Next word_line
            Next word_rect
        Next word_page
    Next word_pane
        
    Cells.Columns.AutoFit

End Sub

