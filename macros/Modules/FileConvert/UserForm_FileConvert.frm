VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_FileConvert 
   Caption         =   "UserForm1"
   ClientHeight    =   3420
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "UserForm_FileConvert.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm_FileConvert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Cancelled As Boolean

Public Property Get Cancelled() As Variant
    Cancelled = m_Cancelled
End Property

Private Sub checkbox_pdf_Click()
    If checkbox_pdf.Value = True Then
        textbox_pdfpath.Enabled = False
        textbox_pdfpath.Visible = False
        checkbox_word.Value = False
    Else
        textbox_pdfpath.Enabled = True
        textbox_pdfpath.Visible = True
    End If
End Sub

Private Sub checkbox_word_Click()
    If checkbox_word.Value = True Then
        textbox_wordpath.Enabled = False
        textbox_wordpath.Visible = False
        checkbox_pdf.Value = False
    Else
        textbox_wordpath.Enabled = True
        textbox_wordpath.Visible = True
    End If
End Sub

Private Sub fselect_output_Click()

    Dim folderpath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .InitialFileName = ActiveWorkbook.path
        
        If .Show = -1 Then ' if OK is pressed
            folderpath = .SelectedItems(1)
        End If
    End With
    
    textbox_outputpath.Text = folderpath
    
End Sub

Private Sub fselect_pdf_Click()
    Dim folderpath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .InitialFileName = ActiveWorkbook.path
        
        If .Show = -1 Then ' if OK is pressed
            folderpath = .SelectedItems(1)
        End If
    End With
    
    textbox_pdfpath.Text = folderpath
End Sub

Private Sub fselect_word_Click()

    Dim folderpath As String
    
    With Application.FileDialog(msoFileDialogFolderPicker)
    
        .InitialFileName = ActiveWorkbook.path
        
        If .Show = -1 Then ' if OK is pressed
            folderpath = .SelectedItems(1)
        End If
    End With
    
    textbox_wordpath.Text = folderpath
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub


Private Sub button_convert_Click()
    ConvertFile
End Sub

Private Sub ConvertFile()
    If Cancelled = False Then
        If checkbox_word = True Then
            Application.Run "FileConvert.PDF_To_Word", textbox_pdfpath, textbox_outputpath
        ElseIf checkbox_pdf = True Then
            Application.Run "FileConvert.Word_To_PDF", textbox_wordpath, textbox_outputpath
        Else
            MsgBox ("No options checked")
        End If
    End If
End Sub

Private Sub button_cancel_Click()
    Hide
    m_Cancelled = True
End Sub

Public Property Get pdfpath() As String
    pdfpath = pdfpath_text.Value
    
End Property
