Attribute VB_Name = "PrintArr"
Private Sub test()
    Dim arr_1d() As Variant
    Dim arr_2d() As Variant
    
    
    
    Dim headers_1d() As Variant
    Dim headers_2d() As Variant
    ReDim arr_1d(1)
    ReDim headers_1d(0)
    headers_1d(0) = "1stCol"
    arr_1d(0) = "1d_Alex"
    arr_1d(1) = "1d_Brandon"
    
    
    ReDim arr_2d(1, 2)
    ReDim headers_2d(2)
    headers_2d(0) = "1stCol"
    headers_2d(1) = "2ndCol"
    headers_2d(2) = "3rdCol"
    arr_2d(0, 0) = "2d_Alex"
    arr_2d(0, 1) = "2d_Brandon"
    arr_2d(1, 0) = "2d_Chris"
    arr_2d(1, 1) = "2d_David"
    arr_2d(0, 2) = "2d_Emily"
    arr_2d(1, 2) = "2d_Fin"
    
    'PrintArr arr_1d, headers_1d
    PrintArr arr_2d, headers_2d

End Sub

Private Sub PrintArr(arr As Variant, Optional headers As Variant)
    Dim colspace As String
    Dim str_row As String
    If IsMissing(headers) = True Then
        PrintHeader arr
    Else
        PrintHeader arr, headers
    End If
    '1D
    If GetDim(arr) = 1 Then
        colspace = CSpc("", ArrMaxLen(arr))
        For i = LBound(arr, 1) To UBound(arr, 1)
            Debug.Print ("[" & GetLeadZero(UBound(arr, 1), CInt(i)) & i & "]| " & AlignWord(CStr(arr(i)), colspace, "Left") & " |")
        Next i
    ElseIf GetDim(arr) = 2 Then
        For i = LBound(arr, 1) To UBound(arr, 1)
            str_row = ""
            For j = LBound(arr, 2) To UBound(arr, 2)
                colspace = CSpc("", ArrMaxLen(arr, CInt(j)))
                If j = LBound(arr, 2) Then
                    str_row = str_row & "[" & GetLeadZero(UBound(arr, 1), CInt(i)) & i & "]| "
                End If
                If j = UBound(arr, 2) Then
                    str_row = str_row & AlignWord(CStr(arr(i, j)), colspace, "Left") & " |"
                Else
                    str_row = str_row & AlignWord(CStr(arr(i, j)), colspace, "Left") & " | "
                End If
            Next j
            Debug.Print (str_row)
        Next i
    End If
End Sub

Private Sub PrintHeader(arr As Variant, Optional headers As Variant)
    '[###]| xxxxx |
    '[leadzero + i]| ArrMaxLen | ArrMaxLen | ….
    On Error Resume Next
    Dim indexspace As String
    Dim colspace As String
    Dim leadzero As String
    Dim row_header As String
    Dim row_index As String
    Dim row_line As String
    'gets index max space
    indexspace = CSpc("[" & GetLeadZero(UBound(arr)) & "]")
    row_header = ""
    row_index = ""
    row_line = ""
    '1D
    If GetDim(arr) = 1 Then
        colspace = CSpc("", ArrMaxLen(arr))
        row_header = indexspace & "| " & AlignWord(CStr(headers(0)), colspace, "Center") & " |"
        row_index = indexspace & "| " & AlignWord("[0]", colspace, "Center") & " |"
    '2D
    ElseIf GetDim(arr) = 2 Then
        For j = LBound(arr, 2) To UBound(arr, 2)
            colspace = CSpc("", ArrMaxLen(arr, CInt(j)))
            leadzero = GetLeadZero(UBound(arr, 2), CInt(j))
            If j = LBound(arr, 2) Then
                row_header = row_header & indexspace & "| "
                row_index = row_index & indexspace & "| "
            End If
            row_header = row_header & AlignWord(CStr(headers(j)), colspace, "Center")
            row_index = row_index & AlignWord("[" & leadzero & j & "]", colspace, "Center")
            If j = UBound(arr, 2) Then
                row_header = row_header & " |"
                row_index = row_index & " |"
            Else
                row_header = row_header & " | "
                row_index = row_index & " | "
            End If
        Next j
    End If
    
    For k = 0 To Len(row_index) - Len(indexspace) - 1
        row_line = row_line & "-"
    Next k
    row_line = indexspace & row_line
    
    If IsMissing(headers) = True Then
        Debug.Print (row_index)
        Debug.Print (row_line)
    Else
        Debug.Print (row_header)
        Debug.Print (row_index)
        Debug.Print (row_line)
    End If
    
End Sub
Private Function GetLeadZero(max As Integer, Optional i As Integer = -1) As String
    Dim leadzero As String
    If i = -1 Then
        If max < 10 Then
            leadzero = "0"
        ElseIf max < 100 And max > 10 Then
            leadzero = "00"
        ElseIf max < 1000 And max > 100 Then
            leadzero = "000"
        End If
    Else
        If max < 10 Then
            leadzero = ""
        ElseIf max < 100 And i < 10 Then
            leadzero = "0"
        ElseIf max < 100 And i >= 10 Then
            leadzero = ""
        ElseIf max < 1000 And i < 100 And i < 10 Then
            leadzero = "00"
        ElseIf max < 1000 And i < 100 And i >= 10 Then
            leadzero = "0"
        Else
            leadzero = ""
        End If
    End If
    GetLeadZero = leadzero
End Function
Private Function AlignWord(str As String, strspace As String, align As String) As String
    'strspace = "----------" (10 spaces)
    'str1 = "Head 1" (6 let)
    'result1 = "--Head 1--" (2spc + 6let + 2spc = 10 space)
    'str2 = "Head1" (5 let)
    'result2 = "--Head1---" (2 spc + 5let + 3spc = 10 spcs)

    'strspace = "-----------" (11 spaces)
    'str1 = "Head 1" (6 let)
    'result1 = "--Head 1---" (2spc + 6let + 3spc = 11 space)
    'str2 = "Head1" (5 let)
    'result2 = "---Head1---" (3 spc + 5let + 3spc = 11 spcs)

    Dim begspace As String
    Dim endspace As String
    Dim spacespill As Integer
    Dim strtemp As String
   
    spacespill = Len(strspace) - Len(str)
    'if spacespill is even
    If spacespill Mod 2 = 0 Then
        begspace = CSpc("", CInt(spacespill / 2))
        endspace = CSpc("", CInt(spacespill / 2))
    'if spacespill is odd
    Else
        begspace = CSpc("", CInt(spacespill / 2))
        endspace = CSpc("", spacespill - Len(begspace))
    End If
    
    If align = "Center" Then
        AlignWord = begspace & str & endspace
    ElseIf align = "Left" Then
        AlignWord = str & begspace & endspace
    ElseIf align = "Right" Then
        AlignWord = begspace & endspace & str
    Else
        AlignWord = "Error Align"
    End If
End Function



Private Function CSpc(str As String, Optional count As Integer = -1) As String
    Dim strtemp As String

    strtemp = ""
    If str <> "" Then
        For k = 0 To Len(str) - 1
            strtemp = strtemp & " "
        Next k
    End If

    If count <> -1 Then
        For k = 0 To count - 1
            strtemp = strtemp & " "
        Next k
    End If
    
    CSpc = strtemp
End Function
Private Function ArrMaxLen(arr As Variant, Optional col As Integer = -1) As Integer
    Dim maxlen As Integer
    '1D
    If col = -1 Then
        maxlen = Len(arr(LBound(arr, 1)))
        For i = LBound(arr, 1) To UBound(arr, 1)
            If Len(arr(i)) > maxlen Then maxlen = Len(arr(i))
        Next i
    
    '2D
    Else
        maxlen = Len(arr(LBound(arr, 1), col))
        For i = LBound(arr, 1) To UBound(arr, 1)
            If Len(arr(i, col)) > maxlen Then maxlen = Len(arr(i, col))
        Next i
    End If
    
    ArrMaxLen = maxlen
End Function

Private Function GetDim(arr As Variant) As Integer
    Dim onedim As Boolean
    Dim twodim As Boolean
        
    On Error GoTo ErrorHandler
    If UBound(arr, 1) > -1 Then onedim = True
    If UBound(arr, 2) > -1 Then twodim = True
    If onedim = True And twodim = True Then
        GetDim = 2
    End If
    
ErrorHandler:
    '1D array
    If onedim = True And twodim = False Then
        GetDim = 1
    End If
    

End Function

