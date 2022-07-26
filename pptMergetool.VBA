'   Run this script inside the same folder of powerpoint that requires merging.
'   Do no attempt to insert this file to itself.
Sub InsertAllSlides()
    
    Dim vArray() As String
    Dim x As Long
    Dim strInput As String
    Dim fileName As String
    Dim saveName As String
    Dim timeStamp As String
    Dim NewPres As Presentation
    
    Set NewPres = Presentations.Add
    
    fileName = InputBox("Enter a name for the Powerpoint: ", "Powerpoint create", "Enter your input text HERE")
    FileLocation = InputBox("Enter a location where you wish to the this file: ", "Powerpoint save location", "Enter your input text HERE")
    saveName = fileName & ".pptx"
    
    Path = FileLocation
    NewPres.SaveAs Path & saveName
    
    Keyword = InputBox("Enter a search keyword: ", "Keyword search", "Enter your input text HERE")
    RunLocation = InputBox("Enter the lococation to search: ", "Powerpoint search location", "Enter your input text HERE")
    
    EnumerateFiles RunLocation & "\", Keyword + "*.PPTX", vArray
        
    With ActivePresentation
        For x = 1 To UBound(vArray)
            If Len(vArray(x)) > 0 Then
                .Slides.InsertFromFile vArray(x), .Slides.count
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
