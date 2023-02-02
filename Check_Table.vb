Sub Check_Table()
Dim LastRow As Long, LastCol As Long, LastCityRow As Long, LastIrishRow As Long, LastAbbrRow As Long, LastCodeRow As Long
Dim i As Long, j As Long, k As Long, m As Integer, n As Integer, p As Integer, S As Integer, Error As Long, str As String, StrR As String, Nm As String, Pth As String, Section As Variant
Dim STime As Double, TTime As Double, CityCount As Long, CitiesSize As Integer, City As String, Char As String, Sht As Worksheet, yAddr As String
Dim Str1 As String, Col As Long, StateAbbr As Variant, StateName As Variant, MsgStr As String, PrintHeaderStr As String, IrishLastName As Boolean, AbbrFirstName As Boolean, Debugg As Boolean
Dim IrishSearchCount As Long, AbbrCount As Long, ErrorType As Long, TotalCount As Long, StrTab As String, Firstltr() As Byte, CapCities As String, Repeat As Boolean
Dim Highlighted() As Boolean, Excluded As Boolean, Dummy As Boolean, IsAState As Boolean, HighCells As Long
'Arrays are dim-ed below

Debugg = False

MsgBox "Processing is usually less than 15 seconds per 1000 records." & vbCrLf & _
        "1. Rows with missing addresses will be highlighted." & vbCrLf & _
        "2. Addresses that don't start with a number or ""PO BOX "" will be highlighted." & vbCrLf & _
        "3. Addresses containing repeated items, incl city and zip, will be highlighted." & vbCrLf & _
        "4. Possible duplicate rows will be highlighted." & vbCrLf & _
        "5. Leading double caps and all caps in names will be highlighted." & vbCrLf & _
        "6. Names and addresses will be checked for non-English characters." & vbCrLf & _
        "7. Missing city and zip will be highlighted." & vbCrLf & _
        "8. Cities will be checked for all caps." & vbCrLf & _
        "9. Cities not in the district(see Cities tab) will be highlighted if missing." & vbCrLf & _
        "10. CBID column values will have a hyperlink to the CB profile." & vbCrLf & _
        "11. Highlighted values will have a hyperlink to the CB profile's relevant tab."

STime = Timer
SortCities  'see sub below

Set Sht = ActiveWorkbook.Sheets("Table")
LastRow = Sht.Range("A1").CurrentRegion.Rows.Count
LastCityRow = Sheets("Cities").Range("A:A").Find(What:="*", searchdirection:=xlPrevious, SearchOrder:=xlByRows).row
LastIrishRow = Sheets("Lookup").Range("I:I").Find(What:="*", searchdirection:=xlPrevious, SearchOrder:=xlByRows).row
LastAbbrRow = Sheets("Lookup").Range("G:G").Find(What:="*", searchdirection:=xlPrevious, SearchOrder:=xlByRows).row
LastCodeRow = Sheets("Interpretation").Range("A:A").Find(What:="*", searchdirection:=xlPrevious, SearchOrder:=xlByRows).row


'fill SubStr with profile tab names
ReDim SubStr(11)
SubStr = Array("", "", "", "", "", "edit", "edit", "address", "address", "address")

'will show listobject as NOT including the header row
ActiveWindow.Zoom = 80
Pth = "https://dashboard.conventionofstates.com/admin/people/"
Application.StatusBar = ""

Nm = "WXYZ"
Range("A1").CurrentRegion.Select    'includes header
Selection.Columns.AutoFit
Selection.EntireColumn.Hidden = False
Selection.ClearFormats  'removes all fill colors
Columns(12).Clear       'clears column 12 of reasons for row highlights
'Must sort data upfront to facilitate finding duplicates.  Will need to sort again later as well.
'Sht.ListObjects(Nm).Sort.SortFields.Clear
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[Last Name]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[First Name]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[Street Address]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'With Sht.ListObjects(Nm).Sort
'    .Header = xlYes
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
'End With

Sht.Sort.SortFields.Clear
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 7), Cells(LastRow, 7)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'first name
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 6), Cells(LastRow, 6)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'lastname
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 8), Cells(LastRow, 8)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'street address
With Sht.Sort
    .SetRange Range(Cells(1, 1), Cells(LastRow, 12))    'include error code column
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Application.WindowState = xlNormal




'Left justify and center justify certain columns
With Columns("C:D")
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = True
    .ColumnWidth = 7
End With
'With Range("C1:D1")
'    .WrapText = True
'End With
With Columns("E:J")
    .HorizontalAlignment = xlLeft
    .VerticalAlignment = xlBottom
    .WrapText = False
End With
'ReDim Peti(12)
'A variant holding the array is far more flexible than a declared array of variants!
'https://stackoverflow.com/questions/37689847/creating-an-array-from-a-range-in-vba

Sheets("Cities").Select
Cities = Application.Transpose(Range(Cells(2, 1), Cells(LastCityRow, 1)))   'from a column. Cities() is indexed to 1, not zero
Sheets("Lookup").Visible = True
Sheets("Lookup").Select
Irish = Application.Transpose(Range(Cells(2, 9), Cells(LastIrishRow, 9)))   'from a column, I:I,  > 600 elements
Abbr = Application.Transpose(Range(Cells(2, 7), Cells(LastAbbrRow, 7)))     'from a column, G:G
StateAbbr = Application.Transpose(Range(Cells(2, 2), Cells(52, 2)))         'from column 2
StateName = Application.Transpose(Range(Cells(2, 1), Cells(52, 1)))         'from column 1
Sheets("Lookup").Visible = False
Sheets("Interpretation").Select
Codes = Application.Transpose(Range(Cells(2, 1), Cells(LastCodeRow, 1)))    'from a column
ReDim CodeCounts(UBound(Codes)) 'can have a value for CodeCounds(0), but not used
Sht.Select
Peti = Application.Transpose(Application.Transpose(Range(Cells(1, 1), Cells(1, 12))))    'So PriorPeti isn't empty on first comparison. From a row (needs 2 transpose statements)
ReDim PriorPeti(5, UBound(Peti)) 'used to hold several record prior to current one
ReDim PriorPetis(UBound(Peti))   'used to hold one Specific record from prior records


On Error Resume Next
CitiesSize = UBound(Cities)
If Err <> 0 Then    'Err# 13 if Cities is a scalar, i.e. only one value in what would otherwise be an array.
    CitiesSize = 1
    'Debug.Print "Error = " & Err & " Error Description = " & Err.Description
End If
Err = 0

'Check for Cities in col A that don't have an initial capital letter
For j = 1 To CitiesSize
    If Left(Cities(j), 1) <> Left(UCase(Cities(j)), 1) Then
        CapCities = CapCities & Cities(j) & ", "
    End If
Next
If CapCities <> "" Then
    MsgBox "Not all cities in Column A of the Cities tab have an initial capital letter." & vbCrLf & _
            "Capitalize the first letter of each part of the city name and re-try." & vbCrLf & vbCrLf & _
            "Errors: " & CapCities
    Exit Sub
End If

'***************************************************************************************************************************
'***************************************************************************************************************************
For i = 2 To LastRow    'i is row counter
    'Application.StatusBar = "Processing Names and Address on row " & i
'    Application.ScreenUpdating = True
'    If i = 70 Then Application.ScreenUpdating = False
'
    If i = 3 Then
        i = i  'for a debugging breakpoint
    End If

    ReDim Highlighted(12)

    'setup prior rows to later check against current row for duplicates
    'newest row in PriorPeti() is (1,x), oldest is (5,x)
    n = IIf(i - 2 < 5, i - 2, 5)
    If i > 2 Then
        For m = n To 1 Step -1
            For k = 1 To UBound(Peti)
                PriorPeti(m, k) = PriorPeti(m - 1, k)
            Next
        Next
        For k = 1 To UBound(Peti)
            PriorPeti(1, k) = Peti(k)   'old Peti from last row
        Next
    End If


    'new Peti *****************************  Peti is short for Petition row
    Peti = Application.Transpose(Application.Transpose(Range(Cells(i, 1), Cells(i, 12))))
    'Peti variant-with-an-array holds all the values in a row
    'Sheets("Archive").Range(Cells(i, 1), Cells(i, 12)) = Peti

    'Check for error cells and blank cells
    For Col = 1 To 10
        Cells(i, Col).Interior.ColorIndex = 0   'reset all highlighting to no fill
        Cells(i, Col).Hyperlinks.Delete         'delete all hyperlinks
        If IsError(Peti(Col)) Then
            Peti(Col) = "DataError" 'value is chosen so as to not trigger other errors checked below
            Cells(i, Col) = "DataError"
                Debug.Print ("Data error in row " & i & ", column " & Col)
            ErrorType = 14  ' " dataErr"
            HighLight i, Col, Peti, ErrorType, Highlighted
            Highlighted(Col) = True
        ElseIf Left(Peti(Col), 1) = "#" Or Peti(Col) = "DataError" Then
            Peti(Col) = "DataError" 'value is chosen so as to not trigger other errors checked below
            Cells(i, Col) = "DataError"
            ErrorType = 14  ' " dataErr"
            HighLight i, Col, Peti, ErrorType, Highlighted
            Highlighted(Col) = True
        ElseIf VarType(Peti(Col)) = vbBoolean Then
            Cells(i, Col).Value = "'" & Peti(Col)   'leading apostrophe critical to changing data type to string!!!
        ElseIf Trim(Peti(Col)) = "" Then
            ErrorType = 6   '"empty"
            HighLight i, Col, Peti, ErrorType, Highlighted
        End If
        Peti(Col) = Trim(Peti(Col))
    Next

    'Delete periods in Peti and name and address cells
    For Col = 6 To 8
        If InStr(Peti(Col), ".") > 0 Then
            Peti(Col) = Replace(Peti(Col), ".", "")
            Cells(i, Col) = Peti(Col)
        End If
    Next

    'Check for missing or badly formed zip (1)
    k = NmbrCount(Peti(10)) 'it's OK if hyphen is missing
    If k <> 5 And k <> 9 Then
        ErrorType = 12  '"zip"
        HighLight i, 10, Peti, ErrorType, Highlighted
    End If

    If i = 16 Then
        i = i  'for a debugging breakpoint
    End If

    'Addresses *******************************************************************
    'Highlight missing or non-numeric street number. "PO Box " is OK.
    'Do not highlight street numbers that start with N,S,E,W followed by a number or a space and a number.
    'Highlight any part of address that is alpha characters only with all caps or no caps

    Firstltr() = Peti(8)    'produce byte array from string
    'In one state, street numbers can start with N,S,E, or W followed by the number.
    'We look at second character after NSEW because first character might be a space.
    'Byte decimal ASCII for N,S,E,W: 78,83,69,87.  Odd indexes of Byte array are zeros.  ASCII numbers are 48 through 57  decimal.
    If (Firstltr(0) = 78 Or Firstltr(0) = 83 Or Firstltr(0) = 69 Or Firstltr(0) = 87) And Firstltr(4) > 47 And Firstltr(4) < 58 Then
        'MsgBox "First Letter = " & Firstltr(0) & ",  Second Char = " & Firstltr(2)
        If Firstltr(2) = 32 Then    'space
            Peti(8) = "0" & Right(Peti(8), Len(Peti(8)) - 2)    'delete space
        Else
            Peti(8) = "0" & Right(Peti(8), Len(Peti(8)) - 1)
        End If
    End If
    If (StrComp(Left(CStr(Peti(8)), 2), "HC", vbBinaryCompare) <> 0) And (StrComp(Left(CStr(Peti(8)), 2), "RR", vbBinaryCompare) <> 0) Then
        If (Not IsNumeric(Left(Peti(8), 1))) And (StrComp(Left(CStr(Peti(8)), 7), "PO Box ", vbBinaryCompare) <> 0) Then   'OK if numeric or with specific characters at start
            'could do binary compare with "PO Box " such as StrComp(Left(Str, 7), "PO Box ", vbBinaryCompare) = 0
            'HC (Highway Contract) is common in Oklahoma, RR (rural route) is used in a few places.
            ErrorType = 10  ' " st#po"
            HighLight i, 8, Peti, ErrorType, Highlighted
            Peti(8) = Replace(Peti(8), "PO Box ", "0")
        End If
    Else
        Peti(8) = Replace(Peti(8), "HC ", "0")
        Peti(8) = Replace(Peti(8), "HC", "0")
        Peti(8) = Replace(Peti(8), "RR ", "0")
        Peti(8) = Replace(Peti(8), "RR", "0")
    End If
    Peti(8) = Replace(Peti(8), "# ", "0")
    Peti(8) = Replace(Peti(8), "#", "0")
    If Debugg Then Debug.Print vbCrLf & "************************************" & vbCrLf & Peti(8)
    'Highlight addresses (complete address) with all caps or all lower case
    Str1 = OnlyLowC(Peti(8))            'byte function (see code below)
    If Len(Str1) > 0 And Len(OnlyLowC(Str1)) < 1 Then
        If Str1 <> "th" And Str1 <> "US" Then
            ErrorType = 1  ' " CAPS"
            HighLight i, 8, Peti, ErrorType, Highlighted
            If Debugg Then Debug.Print "Row=" & i & " no k" & " Err=CAPS"
        End If
    ElseIf Len(Peti(8)) > 0 And Len(OnlyCAPS(Peti(8))) < 1 Then
        ErrorType = 2  ' " lower_case"
        HighLight i, 8, Peti, ErrorType, Highlighted
        If Debugg Then Debug.Print "Row=" & i & " no k" & " Err=lower_case"
    End If
    'Highlight address if it contains " ," (space comma) or " ." (space period) or double space
    If InStr(Peti(8), " ,") > 0 Or InStr(Peti(8), " .") > 0 Or InStr(Peti(8), "  ") > 0 Then
        ErrorType = 4  ' " dblspace"
        HighLight i, 8, Peti, ErrorType, Highlighted
        If Debugg Then Debug.Print "Row=" & i & " no k" & " Err=dblspace"
    End If
    'Highlight addresses that includes city or first five of zip
    Str1 = Peti(2) & "   " & Peti(9) & "   " & Peti(10)
    'Note that a city often has a street, dr, blvd, etc named after the city.  DupeCity addresses this.
    If DupeCity(Peti(8), Peti(9)) Or (InStr(Peti(8), Left(Peti(10), 5)) > 0) Then
        ErrorType = 7  ' " incl_cityzip"
        HighLight i, 8, Peti, ErrorType, Highlighted
        If Debugg Then Debug.Print "Row=" & i & " no k" & " Err=incl_cityzip"
    End If

    'Address  Sections **************************************************************
    'Prepare Address for section by section analysis.  For example, "123 Main St" has three sections separated by spaces.
    'MsgBox Peti(8)
    yAddr = CStr(Peti(8))
    'yAddr = NoPunct(Trim(yAddr))
    yAddr = Replace(yAddr, ",", " ")
    yAddr = Replace(yAddr, "   ", " ")
    yAddr = Replace(yAddr, "  ", " ")
    yAddr = Replace(yAddr, "# ", "#")
    If Len(yAddr) > 0 Then
        'MsgBox "Before split address: " & yAddr
        Section = Split(yAddr, " ")     'first element is index zero
        For k = 0 To UBound(Section)
            str = str & k & "|" & Section(k) & "|" & vbCrLf
        Next
        'MsgBox "After split address: " & vbCrLf & Str

    If i = 16 Then
        i = i    'for debugging at a particular row
    End If

        For k = 1 To UBound(Section)    'skip st #, Section has zero index
'            Excluded = Section(k) <> "N" And Section(k) <> "S" And Section(k) <> "E" And Section(k) <> "W"
'            Excluded = Excluded And Section(k) <> "NE" And Section(k) <> "NW" And Section(k) <> "SE" And Section(k) <> "SW"
            Excluded = Not IsDirections(Section(k))
            'If Not Excluded Then MsgBox "Row=" & i & " k=" & k & " Section(k)=" & Section(k) & "  Excluded: IsDirections = " & IsDirections(Section(k))
            Excluded = Excluded And Section(k) <> "th" And Section(k) <> "PO" And Section(k) <> "US" And NmbrCount(Section(k)) = 0

            If Excluded And Len(Section(k)) > 1 Then
                If Not IsNumeric(Left(Section(k), 1)) Then
                    'section is either all caps or no caps?
                    'MsgBox k & vbCrLf & Section(k) & vbCrLf & OnlyCAPS(Section(k)) & vbCrLf & OnlyLowC(Section(k))
                    If Len(OnlyLowC(Section(k))) < 1 Then   '
                        IsAState = False
                        If Len(Section(k)) = 2 Then
                            'Check to see if these 2 chars are a state abbreviation
                            For S = 1 To 51
                                If Section(k) = StateAbbr(S) Then
                                    IsAState = True
                                    Exit For
                                End If
                            Next
                        End If
                        If Not IsAState Then
                            ErrorType = 1  ' " CAPS"
                            HighLight i, 8, Peti, ErrorType, Highlighted
                            If Debugg Then Debug.Print "Row=" & i & " k=" & k & " Err=CAPS"
                        End If
                    ElseIf Len(OnlyCAPS(Section(k))) < 1 Then
                        ErrorType = 2  ' " lower_case"
                        HighLight i, 8, Peti, ErrorType, Highlighted
                        If Debugg Then Debug.Print "Row=" & i & " k=" & k & " Err=lower_case"
                    Else
                        StrR = Right(Left(Section(k), 2), 1)
                        If Len(OnlyCAPS(StrR)) = 1 Then
                            ErrorType = 3  ' " 2nd_CAP"
                            HighLight i, 8, Peti, ErrorType, Highlighted
                            If Debugg Then Debug.Print "Row=" & i & " k=" & k & " Err=2nd_CAP"
                        End If
'                        For s = 1 To 51 'Check if section(k) is a full, proper case state name
'                            If Section(k) = StateName(s) Then
'                                IsAState = True
'                                Exit For
'                            End If
'                        Next
                    End If
                    'city or partial zip in address?
                    'debate whether to look at k = 1, 2, or later...
                    If k > 2 And ((LCase(Section(k)) = LCase(Peti(9)) Or Left(Section(k), 4) = Left(Peti(10), 4)) Or LCase(Section(k)) = LCase(Peti(2))) Then   'in most cases (K>1), ignores street # and  street name
                        ErrorType = 7  ' " incl_cityzip"  or state
                        HighLight i, 8, Peti, ErrorType, Highlighted
                        If Debugg Then Debug.Print "Row=" & i & " k=" & k & " Err=incl_cityzip"
                    End If

                    Repeat = False
                    'any section element repeats?
                    For S = 0 To k
                        If k <> S And LCase(Section(k)) = LCase(Section(S)) Then Repeat = True
                    Next
                    If Repeat Then
                        ErrorType = 13  ' " repeat"
                        HighLight i, 8, Peti, ErrorType, Highlighted
                        If Debugg Then Debug.Print "Row=" & i & " k=" & k & " Err=repeat"
                    End If
                End If
            End If
        Next
    End If

    'Check for repeated elements in address
'    If AddressRepetition(Peti(8)) Then
'        ErrorType = 13  ' " repeat"
'        HighLight i, 8, Peti, ErrorType, Highlighted
'    End If

    If i = 3 Then
        i = i    'for debugging at a particular row
    End If

    'Check for possible duplicate rows among the previous 5 rows
    'Peti(6) is first name. Check for match on first character only
    'Peti(7) is last name. Check for match on entire word
    'Peti(8) is street address.  Check for match on first 8 characters
    If i > 2 Then
        For m = 1 To n
            If Left(LCase(Peti(6)), 1) = Left(LCase(PriorPeti(m, 6)), 1) And LCase(Peti(7)) = LCase(PriorPeti(m, 7)) And Left(LCase(Peti(8)), 8) = Left(LCase(PriorPeti(m, 8)), 8) Then 'And Left(LCase(Peti(10)), 5) = Left(LCase(PriorPeti(m, 10)), 5) Then
                ErrorType = 5   ' " duplicate"
                HighLight i, 7, Peti, ErrorType, Highlighted
                For k = 0 To UBound(Peti)
                    PriorPetis(k) = PriorPeti(m, k)
                Next
                If Debugg Then Debug.Print PriorPetis(1) & ", " & PriorPetis(2)
                k = i - m   'row of other record in duplicate pair
                'PriorPetiS = Application.Transpose(Application.Transpose(PriorPeti(m)))
                HighLight k, 7, PriorPetis, ErrorType, Highlighted
                'Debug.Print "Row " & i & ", " & m & " " & Peti(6) & ", " & Peti(7) & ", " & Peti(8) & vbCrLf & String(Len(i), " ") & "       " & PriorPeti(m, 6) & ", " & PriorPeti(m, 7) & ", " & PriorPeti(m, 8)
                Exit For
            End If
        Next
    End If

    If i = 21 Then
        i = i    'for debugging at a particular row
    End If
    IrishLastName = False
    AbbrFirstName = False
    For Col = 6 To 7
        If i = 21 And Col = 7 Then
            Col = 7 'debugging
        End If
        str = Peti(Col)
        'first or second name
        If UCase(str) = "TRUE" Then
            'Peti(Col) = WorksheetFunction.Proper(Str)
        End If
        StrR = Right(Left(str, 2), 1)  'second letter
        If Len(Peti(Col)) > 1 And Len(OnlyLowC(Peti(Col))) < 1 Then  'there should be lower case letters
            If Col = 6 And Len(Peti(6)) = 2 Then
                For k = 1 To UBound(Abbr)  'k is an array index, not a row number
                    If Peti(6) = Abbr(k) Then
                        AbbrFirstName = True
                        AbbrCount = AbbrCount + 1
                        Exit For
                    End If
                Next
            End If
            If Not AbbrFirstName Then   'first names are done first
                ErrorType = 1  ' " CAPS"
                HighLight i, Col, Peti, ErrorType, Highlighted
            End If
        ElseIf Len(Peti(Col)) > 0 And Len(OnlyCAPS(Peti(Col))) < 1 Then  'there should be upper case letters
            ErrorType = 2  ' " lower_case"
            HighLight i, Col, Peti, ErrorType, Highlighted
        ElseIf StrR <> " " And StrR <> "." Then                          'there should not be two initial capital letters
            If Len(Peti(Col)) > 1 And Len(Left(OnlyCAPS(Peti(Col)), 1)) = 1 And Len(OnlyCAPS(StrR)) = 1 Then
                If Col = 7 And LCase(Left(Peti(Col), 1)) = "o" Then
                    For k = 1 To UBound(Irish)  'k is an array index, not a row number
                        If LCase(Peti(Col)) = LCase(Irish(k)) Then
                            IrishLastName = True
                            IrishSearchCount = IrishSearchCount + 1
                            Exit For
                        End If
                    Next
                End If
                If Not IrishLastName Then   'first names are done first
                    ErrorType = 3  ' " 2nd_CAP"
                    HighLight i, Col, Peti, ErrorType, Highlighted
                End If
            End If
        End If
    Next
    'last names must have at least two characters
    If Len(Peti(7)) < 2 Then
        ErrorType = 11  ' " <2chars"
        HighLight i, 7, Peti, ErrorType, Highlighted
    End If

    'Check for non-English characters in names - allow "&", spaces, dashes, apostrophes (O'Neill), NOT COMMAS
    For Col = 6 To 7
        If Peti(Col) <> "" And Len(MarkForeignName(Peti(Col))) > 0 Then
            ErrorType = 9  ' " non-english"
            HighLight i, Col, Peti, ErrorType, Highlighted
        End If
    Next

    'Check for non-English characters in addresses - allow numbers, commas, spaces, dashes, #
    If Peti(8) <> "" And Len(MarkForeignAddr(Peti(8))) > 0 Then
        ErrorType = 9   ' " non-english"
        HighLight i, Col, Peti, ErrorType, Highlighted
    End If

    If i = 982 Then
        i = i    'for debugging at a particular row
    End If

    'Is City in the "Good City" column (A) on Cities sheet?
    City = Peti(9)
    CityCount = 0
    'j is the index of the Cities array, not rows on the Cities tab.
    'CitiesSize is set above the i (Excel row) For loop
    For j = 1 To CitiesSize
        If Peti(9) = Cities(j) Then
            CityCount = CityCount + 1
            GoTo CityCounted
        End If
    Next
CityCounted:
    If CityCount = 0 Or City = "" Or Len(OnlyLowC(City)) < 1 Or Len(OnlyCAPS(City)) < 1 Then
        ErrorType = 8   ' " invalid_city"
        HighLight i, 9, Peti, ErrorType, Highlighted
    End If
    'Cells(i, 12) = Peti(12)

    'Add hyperlink to Cb Id (CB ID) column *****************************
    ActiveSheet.Hyperlinks.Add Anchor:=Sht.Cells(i, 5), Address:=Pth & Peti(5) 'keep cell text
Next    'next i (Table row)

'Sort Again - primary purpose is to put people with last name of "True" into alphabetical order.
Sht.Sort.SortFields.Clear
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 7), Cells(LastRow, 7)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'first name
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 6), Cells(LastRow, 6)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'lastname
Sht.Sort.SortFields.Add2 Key:=Range(Cells(1, 8), Cells(LastRow, 8)) _
    , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal      'street address
With Sht.Sort
    .SetRange Range(Cells(1, 1), Cells(LastRow, 12))    'include error code column
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With
Application.WindowState = xlNormal

k = 0
For i = 2 To LastRow
    If Cells(i, 1).Interior.ColorIndex = 6 Then k = k + 1
Next
HighCells = k



'Do some formatting  ***************************************
'Sort Columns in listobject
'Sht.ListObjects(Nm).Sort.SortFields.Clear
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[Last Name]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[First Name]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'Sht.ListObjects(Nm).Sort. _
'    SortFields.Add2 Key:=Range(Nm & "[Street Address]"), SortOn:=xlSortOnValues, _
'    Order:=xlAscending, DataOption:=xlSortNormal
'With Sht.ListObjects(Nm).Sort
'    .Header = xlYes
'    .MatchCase = False
'    .Orientation = xlTopToBottom
'    .SortMethod = xlPinYin
'    .Apply
'End With
'    Create and name structured table (aka ListObjects in VBA) of all petitions.csv data
    On Error Resume Next
    Sht.ListObjects(1).Unlist   'throws error if one doesn't exist
    If Err > 0 Then Err = 0
    ActiveSheet.ListObjects.Add(xlSrcRange, Range("$A$1").CurrentRegion, , xlYes).Name = "WXYZ"

Columns("H:H").Select   'Street Address column
With Selection
    .WrapText = True
    .ColumnWidth = 35
    .Rows.AutoFit
End With
Columns("L:L").Select   'Reasons column
With Selection
    .WrapText = True
    .ColumnWidth = 22
End With

Range(Cells(2, 1), Cells(LastRow, 1)).NumberFormat = "m/d/yyyy"     'set first column format to short date.  Again.
'Range(Nm & "[#Headers]").Font.Color = vbBlack
Format_Header  'see sub below for formatting first row of the data
Application.ScreenUpdating = True

Application.StatusBar = "Print Formatting"
Print_Format 'PrintHeaderStr   'see sub below for formatting header of printed pages
Range("A1").Select
On Error Resume Next
Debug.Print "Results of Check Table:"
Debug.Print "Irish Searches = " & IrishSearchCount
Debug.Print "Abbreviations for First Names = " & AbbrCount
Debug.Print "Error Code          Count"
TotalCount = 0
For k = 1 To LastCodeRow - 1
    m = 24 - Len(Codes(k)) - 1
    m = Application.WorksheetFunction.RoundDown(m / 4, 0)
    StrTab = String(m, vbTab)
    Debug.Print Codes(k) & StrTab & CodeCounts(k)
    TotalCount = TotalCount + CodeCounts(k)
Next
Debug.Print "Total Count =       " & TotalCount
TTime = Round(Timer - STime, 0)
Debug.Print "Processing time, sec = " & TTime

Application.StatusBar = ""
MsgBox "Processing is complete." & vbCrLf & HighCells & " rows were highlighted." & vbCrLf & TTime & " seconds." & vbCrLf & vbCrLf
'        "If you see issues that were not highlighted - or false highlighting - paste the rows below the text in the Demo tab in this file and " & vbCrLf & _
'        "email to Bill Wiltschko, william.wiltschko@cosaction.com."
End Sub
