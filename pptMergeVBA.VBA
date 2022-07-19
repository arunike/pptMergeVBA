'   Run this script inside the same folder of powerpoint that requires merging.
'   Do no attempt to insert this file to itself.
Sub pptMergeVBA()
    '   Variables
    Dim vArray() As String
    Dim x As Long
    Dim strInput As String

    ' Take user input
    Keyword = InputBox("Enter search keyword: ", "Keyword search", "Enter your input text HERE")

    '   Filter by keyword
    EnumerateFiles ActivePresentation.Path & "\", Keyword + "*.PPTX", vArray

    '   For loop to combine the files together
    With ActivePresentation
        For x = 1 To UBound(vArray)
            If Len(vArray(x)) > 0 Then
                .Slides.InsertFromFile vArray(x), .Slides.Count
            End If
        Next
    End With
End Sub

'   Collect all files matching the specifications into vArray, an array of strings
Sub EnumerateFiles(ByVal Directory As String, _
    ByVal FileSpec As String, _
    ByRef vArray As Variant)
    
    '   Variables
    Dim Temp As String
    ReDim vArray(1 To 1)

    '   Set current inputted fle to a temp variable
    Temp = Dir$(Directory & FileSpec)
    
    '   Do while loop to reallocate storage space for array
    Do While Len(Temp) > 0
        If Temp <> ActivePresentation.Name Then
            ReDim Preserve vArray(1 To UBound(vArray) + 1)
            vArray(UBound(vArray)) = Directory & Temp
        End If
        Temp = Dir$
    Loop
End Sub