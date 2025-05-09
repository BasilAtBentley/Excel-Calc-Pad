Attribute VB_Name = "Module1"
Function RemoveBlanks(inputArray As Variant) As Variant

    Dim base As Long
    base = LBound(inputArray)

    Dim result() As String
    ReDim result(base To UBound(inputArray))

    Dim countOfNonBlanks As Long
    Dim i As Long
    Dim myElement As String

    For i = base To UBound(inputArray)
        myElement = inputArray(i)
        If myElement <> vbNullString Then
            result(base + countOfNonBlanks) = myElement
            countOfNonBlanks = countOfNonBlanks + 1
        End If
    Next i
    If countOfNonBlanks = 0 Then
        ReDim result(base To base)
    Else
        ReDim Preserve result(base To base + countOfNonBlanks - 1)
    End If

    RemoveBlanks = result
    
End Function

Function RemoveDups(inputArray As Variant) As Variant

    Dim nFirst As Long, nLast As Long, i As Long
    Dim item As String
    
    Dim arrTemp() As String
    Dim Coll As New Collection

    'Get First and Last Array Positions
    nFirst = LBound(inputArray)
    nLast = UBound(inputArray)
    ReDim arrTemp(nFirst To nLast)

    'Convert Array to String
    For i = nFirst To nLast
        arrTemp(i) = CStr(inputArray(i))
    Next i
    
    'Populate Temporary Collection
    On Error Resume Next
    For i = nFirst To nLast
        Coll.Add arrTemp(i), arrTemp(i)
    Next i
    Err.Clear
    On Error GoTo 0

    'Resize Array
    nLast = Coll.Count + nFirst - 1
    ReDim arrTemp(nFirst To nLast)
    
    'Populate Array
    For i = nFirst To nLast
        arrTemp(i) = Coll(i - nFirst + 1)
    Next i
    
    'Output Array
    RemoveDups = arrTemp

End Function

Function CreatePeriodArray() As Variant

    Dim result As String
    Dim characters As Integer
    Dim myArray() As Variant
    
    'Get Full String in Cell
    result = Range(ActiveCell.Address).Value
    
    'Count Characters in Cell
    characters = Range(ActiveCell.Address).characters.Count
    ReDim myArray(characters)

    'Get Index of Periods
    For i = 1 To characters - 1
        myArray(i - 1) = InStr(i, result, ".")
    Next i
    
    'Output Array
    CreatePeriodArray = myArray

End Function

Function CreateCircumflexArray() As Variant

    Dim result As String
    Dim characters As Integer
    Dim myArray() As Variant
    
    'Get Full String in Cell
    result = Range(ActiveCell.Address).Value
    
    'Count Characters in Cell
    characters = Range(ActiveCell.Address).characters.Count
    ReDim myArray(characters)

    'Get Index of Circumflex
    For i = 1 To characters - 1
        myArray(i - 1) = InStr(i, result, "^")
    Next i
    
    'Output Array
    CreateCircumflexArray = myArray

End Function

Function CreateCommaArray() As Variant

    Dim result As String
    Dim characters As Integer
    Dim myArray() As Variant
    
    'Get Full String in Cell
    result = Range(ActiveCell.Address).Value
    
    'Count Characters in Cell
    characters = Range(ActiveCell.Address).characters.Count
    ReDim myArray(characters)

    'Get Index of Comma
    For i = 1 To characters - 1
        myArray(i - 1) = InStr(i, result, ",")
    Next i
    
    'Output Array
    CreateCommaArray = myArray

End Function

Function SubScript()
    
    'Declare variables
    Dim result As String
    Dim var As Variant
    Dim arr As Variant
    Dim myData As DataObject
    Dim counter As Integer
    
    'Set variables
    Set myData = New DataObject
    myArray = CreatePeriodArray()
    myArray = RemoveDups(myArray)
    myArray = RemoveBlanks(myArray)
    counter1 = 0
    counter2 = 0

    'Remove period
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            Range(ActiveCell.Address).characters(arr - counter1, 1).Delete
            counter1 = counter1 + 1
        End If
    Next arr
    
    'Add subscripts
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            If counter2 = 0 Then
                'On first period
                Range(ActiveCell.Address).characters(arr, 1).Font.SubScript = True
                counter2 = 1
            Else
                'On other periods
                Range(ActiveCell.Address).characters(arr - counter2, 1).Font.SubScript = True
                counter2 = counter2 + 1
            End If
        End If
    Next arr
    
End Function

Function SuperScript()
    
    'Declare variables
    Dim result As String
    Dim var As Variant
    Dim arr As Variant
    Dim myData As DataObject
    Dim counter As Integer
    
    'Set variables
    Set myData = New DataObject
    myArray = CreateCircumflexArray()
    myArray = RemoveDups(myArray)
    myArray = RemoveBlanks(myArray)
    counter1 = 0
    counter2 = 0

    'Remove circumflex
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            Range(ActiveCell.Address).characters(arr - counter1, 1).Delete
            counter1 = counter1 + 1
        End If
    Next arr
    
    'Add superscripts
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            If counter2 = 0 Then
                'On first circumflex
                Range(ActiveCell.Address).characters(arr, 1).Font.SuperScript = True
                counter2 = 1
            Else
                'On other circumflexes
                Range(ActiveCell.Address).characters(arr - counter2, 1).Font.SuperScript = True
                counter2 = counter2 + 1
            End If
        End If
    Next arr
    
End Function

Function GreekConverter()
   
    'Declare variables
    Dim result As String
    Dim var As Variant
    Dim arr As Variant
    Dim myData As DataObject
    Dim counter As Integer
    
    'Set variables
    Set myData = New DataObject
    myArray = CreateCommaArray()
    myArray = RemoveDups(myArray)
    myArray = RemoveBlanks(myArray)
    counter1 = 0
    counter2 = 0

    'Remove Comma
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            Range(ActiveCell.Address).characters(arr - counter1, 1).Delete
            counter1 = counter1 + 1
        End If
    Next arr
    
    'Add GreekLetters
    For Each arr In myArray
    result = Range(ActiveCell.Address).Value
        If arr > 0 Then
            If counter2 = 0 Then
                'On first comma
                Range(ActiveCell.Address).characters(arr, 1).Font.Name = "Symbol"

                counter2 = 1
            Else
                'On other circumflexes
                Range(ActiveCell.Address).characters(arr - counter2, 1).Font.Name = "Symbol"

                counter2 = counter2 + 1
            End If
        End If
    Next arr
    
End Function

Function NormalConverter()
Range(ActiveCell.Address).Style = "Normal"
End Function

Sub Indices()
Attribute Indices.VB_ProcData.VB_Invoke_Func = "h\n14"

Call SuperScript
Call SubScript

End Sub

Sub GreekLetters()
Attribute GreekLetters.VB_ProcData.VB_Invoke_Func = "g\n14"

Call GreekConverter

'Return Default to Calibri
ThisWorkbook.Styles("Normal").Font.Name = "Calibri"

End Sub

Sub NormalLetters()
Attribute NormalLetters.VB_ProcData.VB_Invoke_Func = "j\n14"

Call NormalConverter

End Sub



