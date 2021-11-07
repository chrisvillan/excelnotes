VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_WordLines 
   Caption         =   "Word Lines Macro"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7935
   OleObjectBlob   =   "UserForm_WordLines.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm_WordLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private m_Cancelled As Boolean

Public Property Get Cancelled() As Variant
    Cancelled = m_Cancelled
End Property
Private Sub button_cancel_Click()
    Hide
    m_Cancelled = True
End Sub

Private Sub textbox_text_Change()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
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

Private Sub button_printlines_Click()

    Dim myfiles As Files
    Dim f As File
    Dim filepath As String
    
    Set myfiles = GetFiles
    
    For Each f In myfiles
        filepath = f.path
        Exit For
    Next
    
    Dim wordapp As New Word.Application
    Dim doc As Word.Document
    Set doc = wordapp.Documents.Open(FileName:=filepath)
    wordapp.Visible = False
    Application.Run "WordLines.PrintAllLines", doc
    
    wordapp.Quit
    
End Sub

Private Sub button_replace_Click()
    
    Dim myfiles As Files
    Dim f As File
    Dim fso As New FileSystemObject
    Dim fo As Folder
    Dim sfo As Folder
    Dim fexists As Boolean
    Dim fcount As Integer
    Dim fendcount As Integer
    Dim wordapp As New Word.Application
    Dim doc As Word.Document
    
    Set myfiles = GetFiles
    
    fexists = False
    fcount = 1
    fendcount = 0
    For Each f In myfiles
        'skips ghost files ("~")
        If Left(f.Name, 1) <> "~" Then
            fendcount = fendcount + 1
        End If
    Next
    
    label_progress.Font.Size = 14
    
    For Each f In myfiles
        'skips ghost files ("~")
        If Left(f.Name, 1) <> "~" Then
            Set doc = wordapp.Documents.Open(FileName:=f.path)
            wordapp.Visible = False
            
            label_progress.Caption = "Processing: " & fcount & " out of " & fendcount
            
            Application.Run "WordLines.ReplaceLines", textbox_text.Text, CStr(textbox_linetarget.Text), doc
            
            If checkbox_keeporiginal.Value = True Then
                Set fo = fso.GetFolder(textbox_wordpath.Text)
                For Each sfo In fo.SubFolders
                    If sfo.Name = "Replaced" Then
                        fexists = True
                    Else
                        fexists = False
                    End If
                Next
                
                If fexists = False Then
                    MkDir textbox_wordpath.Text & "\Replaced"
                End If
                
                doc.SaveAs (textbox_wordpath.Text & "\Replaced\" & f.Name)
            Else
                doc.SaveAs (textbox_wordpath.Text & "\" & f.Name)
            End If
            doc.Close
            fcount = fcount + 1
        End If
    
    Next
    wordapp.Quit
    
    label_progress.Caption = "Complete"
    
End Sub

Private Sub button_viewdoc_Click()

End Sub

Private Sub checkbox_insert_Click()

End Sub
Private Sub button_delete_Click()

End Sub

Private Sub button_search_Click()

    Dim myfiles As Files
    Dim f As File
    Dim filepath As String
    Set myfiles = GetFiles
    
    Dim wordapp As New Word.Application
    Dim doc As Word.Document
    
    For Each f In myfiles
        filepath = f.path
        Exit For
    Next
    
    Set doc = wordapp.Documents.Open(FileName:=filepath)
    wordapp.Visible = False
    
    'insert linesearch text validation
    textbox_searchedline.Text = Application.Run("WordLines.GetWordLine", CInt(textbox_linesearch.Text), doc)
    textbox_searchedline.Text = Replace(textbox_searchedline.Text, vbNewLine, "")
    
Exit_Handler:
    wordapp.Quit
    
End Sub

Private Function GetFiles() As Files

    Dim fso As New FileSystemObject
    Dim fo As Folder
    Dim f As File
    Dim temparr() As String
    
    Set fo = fso.GetFolder(textbox_wordpath.Text)
    
    Set GetFiles = fo.Files

End Function
