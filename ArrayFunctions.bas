Attribute VB_Name = "arrayFunctions"
Option Explicit
Option Compare Text
Option Private Module
Option Base 0

'ERROR CODES CONSTANTS
Public Const ARRAY_NOT_PASSED_IN As Integer = 5000
Public Const ARRAY_DIMENSION_INCORRECT As Integer = 5001

'@AUTHOR: ROBERT TODAR

'DEPENDENCIES
' - No dependencies for other modules or library references :)

'PUBLIC FUNCTIONS
' - ArrayAverage
' - ArrayContainsEmpties
' - ArrayDimensionLength
' - ArrayExtractColumn
' - ArrayExtractRow
' - ArrayFilter
' - ArrayFilterTwo
' - ArrayFromRecordset
' - ArrayGetColumnIndex
' - ArrayGetIndexes
' - ArrayIncludes
' - ArrayIndexOf
' - ArrayLength
' - ArrayPluck
' - ArrayPop
' - ArrayPush
' - ArrayPushTwoDim
' - ArrayQuery
' - ArrayRemoveDuplicates
' - ArrayReverse
' - ArrayShift
' - ArraySort
' - ArraySplice
' - ArraySpread
' - ArraySum
' - ArrayToCSVFile
' - ArrayToString
' - ArrayTranspose
' - ArrayUnShift
' - Assign
' - ConvertToArray
' - IsArrayEmpty

'PRIVATE METHODS/FUNCTIONS
' -

'NOTES:
' -

'TODO:
' - CLEAN UP CODE! ADD MORE NOTES AND EXAMPLES.
' - NEED TO REALLY TEST ALL OF THESE FUNCTIONS, CHECK FOR ERRORS.
' - ADD MORE CUSTOM ERROR MESSAGES FOR SPECIFIC ERRORS.
'
' - LOOK THROUGH FUNCTIONS DESIGNED FOR SINGLE DIM ARRAYS, SEE IF CAN CONVERT TO WORK
'   WITH 2 DIM AS WELL
'
' - Create ArrayConcat function


'EXAMPLES OF VARIOUS FUNCTIONS
Private Sub ArrayFunctionExamples()
    
    Dim A As Variant
    
    'SINGLE DIM FUNCTIONS TO MANIPULATE
    ArrayPush A, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
    ArrayPop A                               '--> Banana,Apple --> returns Carrot
    ArrayUnShift A, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
    ArrayShift A                             '--> Orange,Banana,Apple
    ArraySplice A, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
    ArraySplice A, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
    ArrayRemoveDuplicates A                  '--> Mango,Coffee,Banana,Apple
    ArraySort A                              '--> Apple,Banana,Coffee,Mango
    ArrayReverse A                           '--> Mango,Coffee,Banana,Apple
    
    'ARRAY PROPERTIES
    ArrayLength A                            '--> 4
    ArrayIndexOf A, "Coffee"                 '--> 1
    ArrayIncludes A, "Banana"                '--> True
    ArrayContains A, Array("Test", "Banana") '--> True
    ArrayContainsEmpties A                   '--> False
    ArrayDimensionLength A                   '--> 1 (single dim array)
    IsArrayEmpty A                           '--> False
    
    'CAN FLATTEN JAGGED ARRAY WITH SPREAD FORMULA
    A = Array(1, 2, 3, Array(4, 5, 6, Array(7, 8, 9))) 'COULD ALSO SPREAD DICTIONAIRES AND COLLECTIONS AS WELL
    A = ArraySpread(A)                       '--> 1,2,3,4,5,6,7,8,9
    
    'MATH EXAMPLES
    ArraySum A                               '--> 45
    ArrayAverage A                           '--> 5
    
    'FILTER USE'S REGEX PATTERN
    A = Array("Banana", "Coffee", "Apple", "Carrot", "Canolope")
    A = ArrayFilter(A, "^Ca|^Ap")
    
    'ARRAY TO STRING WORKS WITH BOTH SINGLE AND DOUBLE DIM ARRAYS!
    Debug.Print ArrayToString(A)
    
End Sub

'******************************************************************************************
' TESTING SECTION  (REALLY ALL CODE NEEDS TO BE TESTED, BUT THESE ARE MUCH LESS PROVEN)
'******************************************************************************************

'TESTER SUB FOR NEW FUNCTIONS
Private Sub ArrayPlayground()
    
    Dim Arr As Variant
    Arr = Array(0, 1, 2, Array(3, 4, 5), Array(6, 7, Array(8, 9, Array(10, 11, 12, 13, Array(14, 15, 16)))))
    Arr = ArraySpread(Arr)
    
    Debug.Print ArrayToString(Arr)
    
End Sub

'FILTER SINGLE DIM ARRAY ELEMENTS BASED ON REGEX PATTERN
Public Function ArrayFilter(ByVal SourceArray As Variant, ByVal RegExPattern As String) As Variant
    
    '@AUTHOR: ROBERT TODAR
    '@DIM: SINGLE DIM ONLY
    '@REF: https://regexr.com/
    '@EXAMPLE: ArrayFilter(Array("Banana", "Coffee", "Apple"), "^Ba|^Ap") ->  [Banana,Apple]
    
    If ArrayDimensionLength(SourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .Global = False
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = RegExPattern 'SET THE PATTERN THAT WAS PASSED IN
    End With
    
    Dim Index As Long
    For Index = LBound(SourceArray) To UBound(SourceArray)
    
        If RegEx.TEST(SourceArray(Index)) Then
            ArrayPush ArrayFilter, SourceArray(Index)
        End If
        
    Next Index

End Function

'FILTERS MULTIDIMENSIONAL ARRAY. ARGS ARE PAIR BASED: (HEADING TITLE, REGEX) https://regexr.com/ for help
Public Function ArrayFilterTwo(ByVal SourceArray As Variant, ParamArray Args() As Variant) As Variant
    
    '@AUTHOR: ROBERT TODAR
    '@DIM: TWO DIM ONLY
    '@DEPENDINCES: IsValidConditions, ArrayGetConditions, RegExTest
    '@EXAMPLE: ArrayFilterTwo(TwoDimArray, "Name", "^R","ID", "\d{6}", ...) can add as many pair args as you'd like
    
    'THIS FUNCTION IS FOR TWO DIMS ONLY
    If ArrayDimensionLength(SourceArray) <> 2 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a two dimensional array."
    End If
    
    'SHOULD I ALWAYS RETURN HEADING?? THIS ALSO ASSUMES THERE IS A HEADING...
    ArrayPushTwoDim ArrayFilterTwo, ArrayExtractRow(SourceArray, LBound(SourceArray))
    
    'GET CONDITIONS JAGGED ARRAY. (HEADING INDEX, AND REGEX CONDITION)
    Dim Conditions As Variant
    Conditions = ArrayGetConditions(SourceArray, Args)
    
    'CHECK CONDITIONS ON EACH ROW AFTER HEADER
    Dim RowIndex As Integer
    For RowIndex = LBound(SourceArray) + 1 To UBound(SourceArray)
        
        If IsValidConditions(SourceArray, Conditions, RowIndex) Then
            ArrayPushTwoDim ArrayFilterTwo, ArrayExtractRow(SourceArray, RowIndex)
        End If

    Next RowIndex
    
End Function

'SUM A SINGLE DIM ARRAY
Public Function ArraySum(ByVal SourceArray As Variant) As Double
    
    '@AUTHOR: ROBERT TODAR
    '@DIM: SINGLE DIM ONLY
    '@EXAMPLE: ArraySum (Array(5, 6, 4, 3, 2)) -> 20
       
    'SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(SourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a 1 dimensional array."
    End If
    
    Dim Index As Integer
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If Not IsNumeric(SourceArray(Index)) Then
            Err.Raise 55, "ArrayFunctions: ArraySum", SourceArray(Index) & vbNewLine & "^ Element in Array is not numeric"
        End If
        
        ArraySum = ArraySum + SourceArray(Index)
    Next Index
    
End Function

'GET AVERAGE OF SINGLE DIM ARRAY
Public Function ArrayAverage(ByVal SourceArray As Variant) As Double
    
    'SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(SourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    ArrayAverage = ArraySum(SourceArray) / ArrayLength(SourceArray)
    
End Function

'GET LENGTH OF SINGLE DIM ARRAY, REGAURDLESS OF OPTION BASE
Public Function ArrayLength(ByVal SourceArray As Variant) As Integer
    
    On Error Resume Next 'empty means 0 lenght
    ArrayLength = (UBound(SourceArray, 1) - LBound(SourceArray, 1)) + 1
    
End Function

'SPREADS OUT AN ARRAY INTO A SINGLE ARRAY. EXAMPLE: JAGGED ARRAYS, dictionaries, collections.
Public Function ArraySpread(ByVal SourceArray As Variant, Optional SpreadObjects As Boolean = False) As Variant
    
    'THIS FUNCTION IS FOR SINGLE DIMS ONLY
    If ArrayDimensionLength(SourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    'CONVERT ANY DICTIONARY OR COLLECTION INTO AN ARRAY FIRST.
    Dim Temp As Variant
    Temp = ConvertToArray(SourceArray)
    
    Dim Index As Integer
    For Index = LBound(Temp, 1) To UBound(Temp, 1)
        
        'CHECK IF ELEMENT IS AN ARRAY OR OBJECT, RUN RECURSIVE IF SO ON THAT ELEMENT
        If IsArray(Temp(Index)) Or (IsObject(Temp(Index)) And SpreadObjects) Then
            
            'RECURSIVE CALLS UNTIL AT BASE ELEMENTS
            Dim InnerTemp As Variant
            If SpreadObjects Then
                InnerTemp = ArraySpread(ConvertToArray(Temp(Index)), True)
            Else
                InnerTemp = ArraySpread(Temp(Index))
            End If
            
            'ADD EACH ELEMENT TO THE FUNCTION ARRAY
            Dim InnerIndex As Integer
            For InnerIndex = LBound(InnerTemp, 1) To UBound(InnerTemp, 1)
                ArrayPush ArraySpread, InnerTemp(InnerIndex)
            Next InnerIndex
            
        'ELEMENT IS SINGLE ITEM, SIMPLY TO FUNCTION ARRAY
        Else
            
            ArrayPush ArraySpread, Temp(Index)
            
        End If
        
    Next Index
    
End Function

'RETURNS A SINGLE DIM ARRAY OF THE INDEXES OF COLUMN HEADERS
'HEADERS NOT FOUND RETURNS EMPTY IN THAT INDEX
'EXPERIMENTAL CODE PART OF A BIGGER PLAN....
Public Function ArrayGetIndexes(ByVal SourceArray As Variant, ByVal IndexArray As Variant) As Variant
    
    Dim Temp As Variant
    ReDim Temp(LBound(IndexArray) To UBound(IndexArray))
    
    Dim Index As Integer
    For Index = LBound(IndexArray) To UBound(IndexArray)
        Temp(Index) = ArrayGetColumnIndex(SourceArray, IndexArray(Index))
        
        If Temp(Index) = -1 Then
            Temp(Index) = Empty
        End If
        
    Next Index
    
    ArrayGetIndexes = Temp
    
End Function

'CHECK TO SEE IF SINGLE DIM ARRAY CONTAINS ANY EMPTY INDEXES
Public Function ArrayContainsEmpties(ByVal SourceArray As Variant) As Boolean
    
    'THIS FUNCTION IS FOR SINGLE DIMS ONLY
    If ArrayDimensionLength(SourceArray) <> 1 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a single dimensional array."
    End If
    
    Dim Index As Integer
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If IsEmpty(SourceArray(Index)) Then
            ArrayContainsEmpties = True
            Exit Function
        End If
    Next Index
    
End Function

'CHECKS TO SEE IF VALUE IS IN SINGLE DIM ARRAY. VALUE CAN BE SINGLE VALUE OR ARRAY OF VALUES.
'NEED NOTES....
Public Function ArrayContains(ByVal SourceArray As Variant, ByVal Value As Variant) As Boolean
    
    If IsArrayEmpty(SourceArray) Then
        Exit Function
    End If
    
    If IsArray(Value) Then
        Dim ValueIndex As Long
        For ValueIndex = LBound(Value) To UBound(Value)
            
            If ArrayContains(SourceArray, Value(ValueIndex)) Then
                ArrayContains = True
                Exit Function
            End If
            
        Next ValueIndex
        
        Exit Function
    End If
    
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If SourceArray(Index) = Value Then
            ArrayContains = True
            Exit Function
        End If
    Next Index
    
End Function

'CHECK TO SEE IF TWO DIM ARRAY CONTAINS HEADERS STORED IN HEADERS ARRAY
Public Function ArrayContainsHeaders(ByVal SourceArray As Variant, ByVal Headers As Variant) As Variant
    
    If Not IsArray(SourceArray) Or ArrayDimensionLength(SourceArray) <> 2 Then
        Err.Raise 555, "SourceArray must be passed in as an two dimensional array"
    End If
    
    If Not IsArray(Headers) Or ArrayDimensionLength(Headers) <> 1 Then
        Err.Raise 555, "Headers must be passed in as a 1 dimensional array"
    End If
    
    Dim HeaderArray As Variant
    HeaderArray = ArrayExtractRow(SourceArray, LBound(SourceArray, 1))
    
    Dim HedIndex As Integer
    For HedIndex = LBound(Headers, 1) To UBound(Headers, 1)
        
        If ArrayIncludes(HeaderArray, Headers(HedIndex)) = False Then
            Exit Function
        End If
        
    Next HedIndex
    
    ArrayContainsHeaders = True
    
End Function

'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    On Error GoTo Catch
    
    Dim I As Integer
    Dim TEST As Integer

    Do
        I = I + 1
        
        'WAIT FOR ERROR
        TEST = UBound(SourceArray, I)
    Loop
    
Catch:
    ArrayDimensionLength = I - 1

End Function


'GET A COLUMN FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
Public Function ArrayExtractColumn(ByVal SourceArray As Variant, ByVal ColumnIndex As Integer) As Variant
    
    'SINGLE DIM ARRAYS ONLY
    If ArrayDimensionLength(SourceArray) <> 2 Then
        Err.Raise ARRAY_DIMENSION_INCORRECT, , "SourceArray must be a two dimensional array."
    End If
    
    Dim Temp As Variant
    ReDim Temp(LBound(SourceArray, 1) To UBound(SourceArray, 1))
    
    Dim RowIndex As Integer
    For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        Temp(RowIndex) = SourceArray(RowIndex, ColumnIndex)
    Next RowIndex
    
    ArrayExtractColumn = Temp
    
End Function

'GET A ROW FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
Public Function ArrayExtractRow(ByVal SourceArray As Variant, ByVal RowIndex As Long) As Variant
    
    Dim Temp As Variant
    ReDim Temp(LBound(SourceArray, 2) To UBound(SourceArray, 2))
    
    Dim ColIndex As Integer
    For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
        Temp(ColIndex) = SourceArray(RowIndex, ColIndex)
    Next ColIndex
    
    ArrayExtractRow = Temp
    
End Function

'RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
'ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
Public Function ArrayFromRecordset(Rs As Object, Optional IncludeHeaders As Boolean = True) As Variant
    
    '@NOTE: -Int(IncludeHeaders) RETURNS A BOOLEAN TO AN INT (0 OR 1)
    Dim HeadingIncrement As Integer
    HeadingIncrement = -Int(IncludeHeaders)
    
    'CHECK TO MAKE SURE THERE ARE RECORDS TO PULL FROM
    If Rs.BOF Or Rs.EOF Then
        Exit Function
    End If
    
    'STORE RS DATA
    Dim rsData As Variant
    rsData = Rs.GetRows
    
    'REDIM TEMP TO ALLOW FOR HEADINGS AS WELL AS DATA
    Dim Temp As Variant
    ReDim Temp(LBound(rsData, 2) To UBound(rsData, 2) + HeadingIncrement, LBound(rsData, 1) To UBound(rsData, 1))
        
    If IncludeHeaders = True Then
        'GET HEADERS
        Dim headerIndex As Long
        For headerIndex = 0 To Rs.fields.Count - 1
            Temp(LBound(Temp, 1), headerIndex) = Rs.fields(headerIndex).Name
        Next headerIndex
    End If
    
    'GET DATA
    Dim RowIndex As Long
    Dim ColIndex As Long
    For RowIndex = LBound(Temp, 1) + HeadingIncrement To UBound(Temp, 1)
        
        For ColIndex = LBound(Temp, 2) To UBound(Temp, 2)
            Temp(RowIndex, ColIndex) = rsData(ColIndex, RowIndex - HeadingIncrement)
        Next ColIndex
        
    Next RowIndex
    
    'RETURN
    ArrayFromRecordset = Temp
    
End Function

'LOOKS FOR VALUE IN FIRST ROW OF A TWO DIMENSIONAL ARRAY, RETURNS IT'S COLUMN INDEX
Public Function ArrayGetColumnIndex(ByVal SourceArray As Variant, ByVal HeadingValue As String) As Integer
    
    Dim ColumnIndex As Integer
    For ColumnIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
        If SourceArray(LBound(SourceArray, 1), ColumnIndex) = HeadingValue Then
            ArrayGetColumnIndex = ColumnIndex
            Exit Function
        End If
    Next ColumnIndex
    
    'RETURN NEGATIVE IF NOT FOUND
    ArrayGetColumnIndex = -1
    
End Function

'CHECKS TO SEE IF VALUE IS IN SINGLE DIM ARRAY
Public Function ArrayIncludes(ByVal SourceArray As Variant, ByVal Value As Variant) As Boolean
    
    If IsArrayEmpty(SourceArray) Then
        Exit Function
    End If
    
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If SourceArray(Index) = Value Then
            ArrayIncludes = True
            Exit For
        End If
    Next Index
    
End Function

'RETURNS INDEX OF A SINGLE DIM ARRAY ELEMENT
Public Function ArrayIndexOf(ByVal SourceArray As Variant, ByVal SearchElement As Variant) As Integer
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If SourceArray(Index) = SearchElement Then
            ArrayIndexOf = Index
            Exit Function
        End If
    Next Index
    Index = -1
End Function

'EXTRACTS LIST OF GIVEN PROPERTY. MUST BE ARRAY THAT CONTAINS DICTIONRIES AT THIS TIME.
Public Function ArrayPluck(ByVal SourceArray As Variant, ByVal Key As Variant) As Variant
    
    Dim Temp As Variant
    ReDim Temp(LBound(SourceArray, 1) To UBound(SourceArray, 1))
    
    Dim Index As Integer
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        Assign Temp(Index), SourceArray(Index)(Key)
    Next Index

    ArrayPluck = Temp
    
End Function

'REMOVES LAST ELEMENT IN ARRAY, RETURNS POPPED ELEMENT
Public Function ArrayPop(ByRef SourceArray As Variant) As Variant
    
    If Not IsArrayEmpty(SourceArray) Then
        Select Case ArrayDimensionLength(SourceArray)
            
            Case 1:
                ArrayPop = SourceArray(UBound(SourceArray, 1))
                ReDim Preserve SourceArray(LBound(SourceArray, 1) To UBound(SourceArray, 1) - 1)
            
            Case 2:
            
                Dim Temp As Variant
                ReDim Temp(LBound(SourceArray, 2) To UBound(SourceArray, 2))
                
                Dim ColIndex As Integer
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    Temp(ColIndex) = SourceArray(UBound(SourceArray, 1), ColIndex)
                Next ColIndex
                ArrayPop = Temp
                
                ArrayTranspose SourceArray
                ReDim Preserve SourceArray(LBound(SourceArray, 1) To UBound(SourceArray, 1), LBound(SourceArray, 2) To UBound(SourceArray, 2) - 1)
                ArrayTranspose SourceArray
                
        End Select
        
    End If
    
End Function

'ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function ArrayPush(ByRef SourceArray As Variant, ParamArray Element() As Variant) As Long

    Dim Index As Long
    Dim FirstEmptyBound As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    'THIS IS ONLY FOR SINGLE DIMENSIONS.
    If ArrayDimensionLength(SourceArray) = 2 Then  'Or IsArray(Element(LBound(Element)))
    
        'THIS SECTION IS EXPERIMENTAL... ArrayPushTwoDim IS NOT YET PROVEN. REMOVE IF DESIRED.
        ArrayPush = ArrayPushTwoDim(SourceArray, CVar(Element))
        Exit Function
    End If
    
    'REDIM IF EMPTY, OR INCREASE ARRAY IF NOT EMPTY
    If IsArrayEmpty(SourceArray) Then
    
        ReDim SourceArray(OptionBase To UBound(Element, 1) + OptionBase)
        FirstEmptyBound = LBound(SourceArray, 1)
        
    Else
        FirstEmptyBound = UBound(SourceArray, 1) + 1
        ReDim Preserve SourceArray(UBound(SourceArray, 1) + UBound(Element, 1) + 1)
        
    End If
    
    'LOOP EACH NEW ELEMENT
    For Index = LBound(Element, 1) To UBound(Element, 1)
        
        'ADD ELEMENT TO THE END OF THE ARRAY
        Assign SourceArray(FirstEmptyBound), Element(Index)
        
        'INCREMENT TO THE NEXT firstEmptyBound
        FirstEmptyBound = FirstEmptyBound + 1
        
    Next Index
    
    'RETURN NEW ARRAY LENGTH
    ArrayPush = UBound(SourceArray, 1) + 1

End Function

'ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function ArrayPushTwoDim(ByRef SourceArray As Variant, ParamArray Element() As Variant) As Long

    Dim FirstEmptyRow As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    'REDIM IF EMPTY, OR INCREASE ARRAY IF NOT EMPTY
    If IsArrayEmpty(SourceArray) Then
    
        ReDim SourceArray(OptionBase To UBound(Element, 1) + OptionBase, OptionBase To ArrayLength(Element(LBound(Element))) + OptionBase - 1)
        FirstEmptyRow = LBound(SourceArray, 1)
        
    Else
    
        FirstEmptyRow = UBound(SourceArray, 1) + 1
        SourceArray = ArrayTranspose(SourceArray)
        ReDim Preserve SourceArray(LBound(SourceArray, 1) To UBound(SourceArray, 1), LBound(SourceArray, 2) To UBound(SourceArray, 2) + ArrayLength(Element))
        SourceArray = ArrayTranspose(SourceArray)
    End If
    
    'LOOP EACH ARRAY
    Dim Index As Long
    For Index = LBound(Element, 1) To UBound(Element, 1)
        
        
        Dim CurrentIndex As Long
        CurrentIndex = LBound(Element(Index))
        
        'LOOP EACH ELEMENT IN CURRENT ARRAY
        Dim ColIndex As Long
        For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
            
            'ADD ELEMENT TO THE END OF THE ARRAY. NOTE IF ERROR CHANCES ARE ARRAY DIM WAS NOT THE SAME
            Assign SourceArray(FirstEmptyRow, ColIndex), Element(Index)(CurrentIndex)
            
            CurrentIndex = CurrentIndex + 1
            
        Next ColIndex
        
        'INCREMENT TO THE NEXT firstEmptyRow
        FirstEmptyRow = FirstEmptyRow + 1
        
    Next Index
    
    'RETURN NEW ARRAY LENGTH
    ArrayPushTwoDim = UBound(SourceArray, 1) - LBound(SourceArray, 1) + 1

End Function


' CREATES TEMP TEXT FILE AND SAVES ARRAY VALUES IN A CSV FORMAT,
' THEN QUERIES AND RETURNS ARRAY.
Public Function ArrayQuery(SourceArray As Variant, sql As String, Optional IncludeHeaders As Boolean = True) As Variant
    
    '@USES ArrayToCSVFile
    '@USES ArrayFromRecordset
    '@RETURNS 2D ARRAY || EMPTY (IF NO RECORDS)
    '@PARAM {ARR} MUST BE A TWO DIMENSIONAL ARRAY, SETUP AS IF IT WERE A TABLE.
    '@PARAM {SQL} ADO SQL STATEMENT FOR A TEXT FILE. MUST INCLUDE 'FROM []'
    '@PARAM {IncludeHeaders} BOOLEAN TO RETURN HEADERS WITH DATA OR NOT
    '@EXAMPLE SQL = "SELECT * FROM [] WHERE [FIRSTNAME] = 'ROBERT'"
    
    'CREATE TEMP FOLDER AND FILE NAMES
    Const FileName As String = "temp.txt"
    Dim FilePath As String
    FilePath = Environ("temp")
    
    'UPDATE SQL WITH TEMP FILE NAME
    sql = Replace(sql, "FROM []", "FROM [" & FileName & "]")
    
    'SEND ARRAY TO TEMP TEXTFILE IN CSV FORMAT
    ArrayToCSVFile SourceArray, FilePath & "\" & FileName
    
    'CREATE CONNECTION TO TEMP FILE - CONNECTION IS SET TO COMMA SEPERATED FORMAT
    Dim cnn As Object
    Set cnn = CreateObject("ADODB.Connection")
    cnn.Provider = "Microsoft.Jet.OLEDB.4.0"
    cnn.ConnectionString = "Data Source=" & FilePath & ";" & "Extended Properties=""text;HDR=Yes;FMT=Delimited;"""
    cnn.Open
    
    'CREATE RECORDSET AND QUERY ON PASSED IN SQL (QUERIES THE TEMP TEXT FILE)
    Dim Rs As Object
    Set Rs = CreateObject("ADODB.RecordSet")
    With Rs
        .ActiveConnection = cnn
        .Open sql
        
        'GET AN ARRAY FROM THE RECORDSET
         ArrayQuery = ArrayFromRecordset(Rs, IncludeHeaders)
        .Close
    End With
    
    'CLOSE CONNECTION AND KILL TEMP FILE
    cnn.Close
    Kill FilePath & "\" & FileName
    
End Function

'REMOVED DUPLICATES FROM SINGLE DIM ARRAY
Public Function ArrayRemoveDuplicates(SourceArray As Variant) As Variant
    Dim Dic As Object
    Dim Key As Variant
    
    If Not IsArray(SourceArray) Then
        SourceArray = ConvertToArray(SourceArray)
    End If
    
    Set Dic = CreateObject("Scripting.Dictionary")
    For Each Key In SourceArray
        Dic(Key) = 0
    Next
    ArrayRemoveDuplicates = Dic.Keys
    SourceArray = ArrayRemoveDuplicates
End Function

'REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
Public Function ArrayReverse(SourceArray As Variant) As Variant
    
    Dim Temp As Variant
    
    'REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To ((UBound(SourceArray) + LBound(SourceArray)) \ 2)
        
        'STORE LAST VALUE MINUS THE ITERATION
        Assign Temp, SourceArray(UBound(SourceArray) + LBound(SourceArray) - Index)
        
        'SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        Assign SourceArray(UBound(SourceArray) + LBound(SourceArray) - Index), SourceArray(Index)
        
        'SET FIRST VALUE TO THE STORED LAST VALUE
        Assign SourceArray(Index), Temp
        
    Next Index
    
    ArrayReverse = SourceArray
    
End Function

'REMOVES ELEMENT FROM ARRAY - RETURNS REMOVED ELEMENT **[SINGLE DIMENSION]
Public Function ArrayShift(SourceArray As Variant, Optional ElementNumber As Long = 0) As Variant
    
    If Not IsArrayEmpty(SourceArray) Then

        ArrayShift = SourceArray(ElementNumber)
        
        Dim Index As Long
        For Index = ElementNumber To UBound(SourceArray) - 1
            Assign SourceArray(Index), SourceArray(Index + 1)
        Next Index
        
        ReDim Preserve SourceArray(UBound(SourceArray, 1) - 1)
        
    End If
    
End Function

'SORT AN ARRAY [SINGLE DIMENSION]
Public Function ArraySort(SourceArray As Variant) As Variant
    
    'SORT ARRAY A-Z
    Dim OuterIndex As Long
    For OuterIndex = LBound(SourceArray) To UBound(SourceArray) - 1
        
        Dim InnerIndex As Long
        For InnerIndex = OuterIndex + 1 To UBound(SourceArray)
            
            If UCase(SourceArray(OuterIndex)) > UCase(SourceArray(InnerIndex)) Then
                Dim Temp As Variant
                Temp = SourceArray(InnerIndex)
                SourceArray(InnerIndex) = SourceArray(OuterIndex)
                SourceArray(OuterIndex) = Temp
            End If
            
        Next InnerIndex
    Next OuterIndex
    
    ArraySort = SourceArray

End Function

'CHANGES THE CONTENTS OF AN ARRAY BY REMOVING OR REPLACING EXISTING ELEMENTS AND/OR ADDING NEW ELEMENTS.
Public Function ArraySplice(SourceArray As Variant, Where As Long, HowManyRemoved As Integer, ParamArray Element() As Variant) As Variant
    
    'CHECK TO SEE THAT INSERT IS NOT GREATER THAN THE Array (REDUCE IF SO)
    If Where > UBound(SourceArray, 1) + 1 Then
        Where = UBound(SourceArray, 1) + 1
    End If
    
    'CHECK TO MAKE SURE REMOVED IS NOT MORE THAN THE Array (REDUCE IF SO)
    If HowManyRemoved > (UBound(SourceArray, 1) + 1) - Where Then
        HowManyRemoved = (UBound(SourceArray, 1) + 1) - Where
    End If
    
    If UBound(SourceArray, 1) + UBound(Element, 1) + 1 - HowManyRemoved < 0 Then
        ArraySplice = Empty
        SourceArray = Empty
        Exit Function
    End If
    
    'SET BOUNDS TO TEMP Array
    Dim Temp As Variant
    ReDim Temp(LBound(SourceArray, 1) To UBound(SourceArray, 1) + UBound(Element, 1) + 1 - HowManyRemoved)
    
    'LOOP TEMP Array, ADDING\REMOVING WHERE NEEDED
    Dim Index As Long
    For Index = LBound(Temp, 1) To UBound(Temp, 1)
        
        'INSERT ONCE AT WHERE, AND ONLY VISIT ONCE
        Dim Visited As Boolean
        If Index = Where And Visited = False Then
            
            Visited = True
            
            'ADD NEW ELEMENTS
            Dim Index2 As Long
            Dim Index3 As Long
            For Index2 = LBound(Element, 1) To UBound(Element, 1)
                Temp(Index) = Element(Index2)
                
                'INCREMENT COUNTERS
                Index3 = Index3 + 1
                Index = Index + 1
            Next Index2
            
            
            'GET REMOVED ELEMENTS TO RETURN
            Dim RemovedArray As Variant
            If HowManyRemoved > 0 Then
                ReDim RemovedArray(0 To HowManyRemoved - 1)
                For Index2 = LBound(RemovedArray, 1) To UBound(RemovedArray, 1)
                    RemovedArray(Index2) = SourceArray(Where + Index2)
                Next Index2
            Else
                RemovedArray = Empty
            End If
            
            'DECREMENT COUNTERS FOR AFTER LOOP
            Index = Index - 1
            Index3 = Index3 - HowManyRemoved
        
        Else
            'ADD PREVIOUS ELEMENTS (Index3 IS A HELPER)
            Temp(Index) = SourceArray(Index - Index3)
        End If
        
    Next Index
    
    SourceArray = Temp
    ArraySplice = RemovedArray
    
End Function

'BASICALY ARRAY TO STRING HOWEVER QUOTING STIRNGS, THEN SAVING TO A TEXTFILE
Public Function ArrayToCSVFile(SourceArray As Variant, FilePath As String) As String
    
    Dim Temp As String
    Const Delimiter = ","
    
    Select Case ArrayDimensionLength(SourceArray)
        'SINGLE DIMENTIONAL ARRAY
        Case 1
            Dim Index As Integer
            For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
                If IsNumeric(SourceArray(Index)) Then
                    Temp = Temp & SourceArray(Index)
                Else
                    Temp = Temp & """" & SourceArray(Index) & """"
                End If
            Next Index
            
        
        '2 DIMENSIONAL ARRAY
        Case 2
            Dim RowIndex As Long
            Dim ColIndex As Long
            
            'LOOP EACH ROW IN MULTI ARRAY
            For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
                'LOOP EACH COLUMN ADDING VALUE TO STRING
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    If IsNumeric(SourceArray(RowIndex, ColIndex)) Then
                        Temp = Temp & SourceArray(RowIndex, ColIndex)
                    Else
                        Temp = Temp & """" & SourceArray(RowIndex, ColIndex) & """"
                    End If
                    
                    If ColIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
                Next ColIndex
                
                'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If RowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
        
            Next RowIndex
    End Select
    
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Ts As Object
    Set Ts = Fso.OpenTextFile(FilePath, 2, True) '2=WRITEABLE
    Ts.Write Temp
    
    Set Ts = Nothing
    Set Fso = Nothing
    
    ArrayToCSVFile = Temp
    
End Function


'RESIZE PASSED IN EXCEL RANGE, AND SET VALUE EQUAL TO THE ARRAY
Public Sub ArrayToRange(ByVal SourceArray As Variant, Optional ByRef Target As Excel.Range)

    '@TODO: NEED TO TEST! ALSO THIS ASSUMES ROW, GIVE OPTION TO TRANSPOSE TO COLUMN??
    'NOTE: THIS ALWAYS FORMATS THE CELLS TO BE A STRING... REMOVE FORMATING IF NEED BE.
    '      THIS WAS CREATED FOR THE PURPOSE OF MAINTAINING LEADING ZEROS FOR MY ALL DATA...
    
    'ADD WORKBOOK IF NOT
    Dim Wb As Workbook
    If Target Is Nothing Then
        Set Wb = Workbooks.Add
        Set Target = Wb.Worksheets("Sheet1").Range("A1")
    End If
    
    Select Case ArrayDimensionLength(SourceArray)
        Case 1:
            Set Target = Target.Resize(UBound(SourceArray) - LBound(SourceArray) + 1, 1)
            Target.NumberFormat = "@"
            Target.Value = Application.Transpose(SourceArray)
            
        Case 2:
            Set Target = Target.Resize((UBound(SourceArray, 1) + 1) - LBound(SourceArray, 1), (UBound(SourceArray, 2) + 1 - LBound(SourceArray, 2)))
            Target.NumberFormat = "@"
            Target.Value = SourceArray
            'Target.Resize((UBound(SourceArray, 1) + 1) - LBound(SourceArray, 1), (UBound(SourceArray, 2) + 1 - LBound(SourceArray, 2))).Value = SourceArray
    
    End Select
    
    'OPTIONAL, PLEASE REMOVE IF DESIRED...
    Columns.AutoFit
    
End Sub

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    
    Dim Temp As String
    
    Select Case ArrayDimensionLength(SourceArray)
        'SINGLE DIMENTIONAL ARRAY
        Case 1
            Temp = Join(SourceArray, Delimiter)
        
        '2 DIMENSIONAL ARRAY
        Case 2
            Dim RowIndex As Long
            Dim ColIndex As Long
            
            'LOOP EACH ROW IN MULTI ARRAY
            For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
                'LOOP EACH COLUMN ADDING VALUE TO STRING
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    Temp = Temp & SourceArray(RowIndex, ColIndex)
                    If ColIndex <> UBound(SourceArray, 2) Then Temp = Temp & Delimiter
                Next ColIndex
                
                'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If RowIndex <> UBound(SourceArray, 1) Then Temp = Temp & vbNewLine
        
            Next RowIndex
    End Select
    
    ArrayToString = Temp
    
End Function

'SENDS AN ARRAY TO A TEXTFILE
Public Sub ArrayToTextFile(Arr As Variant, FilePath As String, Optional delimeter As String = ",")
    
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Dim Ts As Object
    Set Ts = Fso.OpenTextFile(FilePath, 2, True) '2=WRITEABLE
    Ts.Write ArrayToString(Arr, delimeter)
    
    Set Ts = Nothing
    Set Fso = Nothing

End Sub

'APPLICATION.TRANSPOSE HAS A LIMIT ON THE SIZE OF THE ARRAY, AND IS LIMITED TO THE 1ST DIM
Public Function ArrayTranspose(SourceArray As Variant) As Variant

    Dim Temp As Variant

    Select Case ArrayDimensionLength(SourceArray)
        
        Case 2:
        
            ReDim Temp(LBound(SourceArray, 2) To UBound(SourceArray, 2), LBound(SourceArray, 1) To UBound(SourceArray, 1))
            
            Dim I As Long
            Dim j As Long
            For I = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                For j = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                    Temp(I, j) = SourceArray(j, I)
                Next
            Next
    
    End Select
    
    ArrayTranspose = Temp
    SourceArray = Temp

End Function

'ADDS NEW ELEMENT TO THE BEGINING OF THE ARRAY
Public Function ArrayUnShift(SourceArray As Variant, ParamArray Element() As Variant) As Long
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If ArrayDimensionLength(SourceArray) <> 1 Then
        ArrayUnShift = -1
        Exit Function
    End If
    
    'RESIZE TEMP ARRAY
    Dim Temp As Variant
    If IsArrayEmpty(SourceArray) Then
        ReDim Temp(0 To UBound(Element, 1))
    Else
        ReDim Temp(UBound(SourceArray, 1) + UBound(Element, 1) + 1)
    End If
    
    Dim Count As Long
    Count = LBound(Temp, 1)
    
    Dim Index As Long
    
    'ADD ELEMENTS TO TEMP ARRAY
    For Index = LBound(Element, 1) To UBound(Element, 1)
        Assign Temp(Count), Element(Index)
        Count = Count + 1
    Next Index
    
    If Not Count > UBound(Temp, 1) Then
    
        'ADD ELEMENTS FROM ORIGINAL ARRAY
        For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
            Assign Temp(Count), SourceArray(Index)
            Count = Count + 1
        Next Index
    End If
    
    'SET ARRAY TO TEMP ARRAY
    SourceArray = Temp
    
    'RETURN THE NEW LENGTH OF THE ARRAY
    ArrayUnShift = UBound(SourceArray, 1) + 1
    
End Function

'QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Public Function Assign(ByRef Variable As Variant, ByVal Value As Variant)

    If IsObject(Value) Then
        Set Variable = Value
    Else
        Let Variable = Value
    End If
    
End Function

'CONVERT OTHER LIST OBJECTS TO AN ARRAY
Public Function ConvertToArray(ByRef Val As Variant) As Variant
    
    Select Case TypeName(Val)
    
        Case "Collection":
            Dim Index As Integer
            For Index = 1 To Val.Count
                ArrayPush ConvertToArray, Val(Index)
            Next Index
        
        Case "Dictionary":
            ConvertToArray = Val.items()
        
        Case Else
             
            If IsArray(Val) Then
                ConvertToArray = Val
            Else
                ArrayPush ConvertToArray, Val
            End If
            
    End Select
    
End Function

Public Function IsArrayEmpty(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' CPEARSON
' This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Err.Clear
    On Error Resume Next
    If IsArray(Arr) = False Then
        ' we weren't passed an array, return True
        IsArrayEmpty = True
    End If

    ' Attempt to get the UBound of the array. If the array is
    ' unallocated, an error will occur.
    Dim ub As Long
    ub = UBound(Arr, 1)
    If (Err.Number <> 0) Then
        IsArrayEmpty = True
    Else
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' On rare occasion, under circumstances I cannot reliably replicate, Err.Number
        ' will be 0 for an unallocated, empty array. On these occasions, LBound is 0 and
        ' UBound is -1. To accommodate the weird behavior, test to see if LB > UB.
        ' If so, the array is not allocated.
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Err.Clear
        Dim LB As Long
        LB = LBound(Arr)
        If LB > ub Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function


'******************************************************************************************
' PRIVATE FUNCTIONS
'******************************************************************************************

'CHECKS CURRENT ROW OF A TWO DIM ARRAY TO SEE IF CONDITIONS ARRAY PASSES
'HELPER FUNCTION FOR ARRAYFILTERTWO
Private Function IsValidConditions(ByVal SourceArray As Variant, ByVal Conditions As Variant, ByVal RowIndex As Integer)
    
    'DEPENDINCES: RegExTest
    
    'CHECK CONDITIONS
    Dim Index As Integer
    For Index = LBound(Conditions) To UBound(Conditions)
        
        Dim Value As String
        Value = SourceArray(RowIndex, Conditions(Index)(0))
        
        Dim Pattern As String
        Pattern = CStr(Conditions(Index)(1))
        
        If Not RegExTest(Value, Pattern) Then
            Exit Function
        End If
        
    Next Index
    
    IsValidConditions = True
    
End Function

'GROUPS HEADING INDEX WITH CONDITIONS. RETURNS JAGGED ARRAY.
'HELPER FUNCTION FOR ARRAYFILTERTWO
Private Function ArrayGetConditions(ByVal SourceArray As Variant, ByVal Arguments As Variant) As Variant
    
    'ARGUMENTS ARE PAIRED BY TWOS. (0) = COLUMN HEADING, (1) = REGEX CONDITION
    Dim Index As Integer
    For Index = LBound(Arguments) To UBound(Arguments) Step 2
    
        Dim ColumnIndex As Integer
        ColumnIndex = ArrayGetColumnIndex(SourceArray, Arguments(Index))
        ArrayPush ArrayGetConditions, Array(ColumnIndex, Arguments(Index + 1))
        
    Next Index
    
End Function

'SIMPLE FUNCTION TO TEST REGULAR EXPRESSIONS. FOR HELP SEE:
Private Function RegExTest(ByVal Value As String, ByVal Pattern As String) As Boolean
    
    Dim RegEx As Object
    Set RegEx = CreateObject("vbscript.regexp")
    With RegEx
        .Global = True 'TRUE MEANS IT WILL LOOK FOR ALL MATCHES, FALSE FINDS FIRST ONLY
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = Pattern
    End With
    
    RegExTest = RegEx.TEST(Value)
    
End Function


