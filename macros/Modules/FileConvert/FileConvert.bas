Attribute VB_Name = "FileConvert"
Sub UseConvert()
    'add please close all word doc validatioin
    Dim frm As New UserForm_FileConvert
    
    With frm
        .checkbox_word.Value = True
        .Show vbModeless
    End With

End Sub


Private Sub PDF_To_Word(pdf_path As String, output_path As String)

Application.ScreenUpdating = False
Application.DisplayAlerts = False
Application.DisplayStatusBar = True

Dim fso As New FileSystemObject
Dim fo As Folder
Dim f As File


Set fo = fso.GetFolder(pdf_path)

Dim wa As Object
Dim doc As Object

Set wa = CreateObject("word.application")
wa.Visible = True

Dim file_Count As Integer

wa.Visible = False
For Each f In fo.Files
    'skips ghost files
    If Left(f.Name, 1) <> "~" Then
        Application.StatusBar = "Converting - " & file_Count + 1 & "/" & fo.Files.Count
        Set doc = wa.Documents.Open(f.path)
        doc.SaveAs (output_path & "\" & Replace(f.Name, ".pdf", ".docx"))
        doc.Close False
        file_Count = file_Count + 1
    End If
Next f

wa.Quit

MsgBox "All PDF files have been converted in to word", vbInformation
Application.StatusBar = ""

End Sub


Private Sub Word_To_PDF(word_path As String, output_path As String)
 
Application.ScreenUpdating = False
Application.DisplayStatusBar = True

Dim fso As New FileSystemObject
Dim fo As Folder
Dim f As File

Dim wb As Workbook
Dim n As Integer

Dim wordapp As New Word.Application
Dim worddoc As Word.Document

Set fo = fso.GetFolder(word_path)

For Each f In fo.Files
    'skips ghost files
    If Left(f.Name, 1) <> "~" Then
        n = n + 1
        Application.StatusBar = "Processing..." & n & "/" & fo.Files.Count
        
        Set worddoc = wordapp.Documents.Open(f.path)
        
        worddoc.ExportAsFixedFormat output_path & Application.PathSeparator & VBA.Replace(f.Name, ".docx", ".pdf"), wdExportFormatPDF
        worddoc.Close False
    End If
Next

Application.StatusBar = ""
wa.Quit

MsgBox "Process Completed"



End Sub
'
'Sub Excel_To_PDF()
'
'Application.ScreenUpdating = False
'Application.DisplayStatusBar = True
'Dim sh As Worksheet
'Set sh = ThisWorkbook.Sheets("Sheet1")
'
'Dim fso As New FileSystemObject
'Dim fo As Folder
'Dim f As File
'
'Dim wb As Workbook
'
'Dim n As Integer
'
'Set fo = fso.GetFolder(sh.Range("E13").Value)
'
'For Each f In fo.Files
'    VBA.DoEvents
'    n = n + 1
'    Application.StatusBar = "Processing..." & n & "/" & fo.Files.Count
'    Set wb = Workbooks.Open(f.Path)
'    wb.ExportAsFixedFormat xlTypePDF, sh.Range("E14").Value & Application.PathSeparator & VBA.Replace(f.Name, ".xlsx", ".pdf")
'    wb.Close False
'Next
'Application.StatusBar = ""
'MsgBox "Process Completed"
'
'
'End Sub

