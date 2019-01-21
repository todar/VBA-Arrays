Attribute VB_Name = "arrayFunctions"
Option Explicit
Option Compare Text
Option Private Module
Option Base 0


'@AUTHOR: ROBERT TODAR

'DEPENDENCIES
' - N/A

'PUBLIC FUNCTIONS
' - ArrayDimensionLength
' - ArrayExtract
' - ArrayFromRecordset
' - ArrayGetColumnNumber
' - ArrayIncludes
' - ArrayIndexOf
' - ArrayPop
' - ArrayPush
' - ArrayQuery
' - ArrayRemoveDuplicates
' - ArrayReverse
' - ArrayShift
' - ArraySort
' - ArraySplice
' - ArrayToCSVFile
' - ArrayToRange
' - ArrayToString
' - ArrayToTextFile
' - ArrayTranspose
' - ArrayUnShift
' - ConvertToArray
' - IsArrayEmpty

'PRIVATE METHODS/FUNCTIONS
' - Asign

'NOTES:
' - I'VE CREATE AN ARRAY CLASS MODULE THAT DOES MANY OF THESE FUNCTIONS, DECIDED TO ALSO
' - CREATE FUNCTIONS AWAY FROM CLASS MODULE OBJECT TO MAKE THEM WORK WITH ANY ARRAY.

'TODO:
' - LOOK THROUGH FUNCTIONS DESIGNED FOR SINGLE DIM ARRAYS, SEE IF CAN CONVERT TO WORK
'   WITH 2 DIM AS WELL
'
' - GO THROUGH AND FIND PLACES TO ADD PRIVATE HELPER FUNTION Asign
' - Create ArrayConcat function


'EXAMPLES
Private Sub ArrayFunctionExamples()
    
    Dim A As Variant
    
    'SINGLE DIM FUNCTIONS
    ArrayPush A, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
    ArrayPop A                               '--> Banana,Apple --> returns Carrot
    ArrayUnShift A, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
    ArrayShift A                             '--> Orange,Banana,Apple
    ArraySplice A, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
    ArraySplice A, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
    ArrayRemoveDuplicates A                  '--> Mango,Coffee,Banana,Apple
    ArraySort A                              '--> Apple,Banana,Coffee,Mango
    ArrayReverse A                           '--> Mango,Coffee,Banana,Apple
    ArrayIndexOf A, "Coffee"                 '--> 1
    ArrayIncludes A, "Banana"                '--> True
    
    Debug.Print ArrayToString(A)
    
End Sub


'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

' RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    On Error GoTo Catch
    
    Dim I As Integer
    Dim Test As Integer

    Do
        I = I + 1
        
        'WAIT FOR ERROR
        Test = UBound(SourceArray, I)
    Loop
    
Catch:
    ArrayDimensionLength = I - 1

End Function

' GET A COLUMN FROM A TWO DIM ARRAY, AND RETURN A SINLGE DIM ARRAY
Public Function ArrayExtract(SourceArray As Variant, ByVal ColumnIndex As Integer) As Variant
    
    Dim Temp As Variant
    ReDim Temp(LBound(SourceArray, 1) To UBound(SourceArray, 1))
    
    Dim RowIndex As Integer
    For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        Temp(RowIndex) = SourceArray(RowIndex, ColumnIndex)
    Next RowIndex
    
    ArrayExtract = Temp
    
End Function

'RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
'ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
'
'@AUTHOR ROBERT TODAR
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

' LOOKS FOR VALUE IN FIRST ROW OF A TWO DIMENSIONAL ARRAY, RETURNS IT'S COLUMN INDEX
Public Function ArrayGetColumnNumber(SourceArray As Variant, HeadingValue As String) As Integer
    
    Dim ColumnIndex As Integer
    For ColumnIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
        If SourceArray(LBound(SourceArray, 1), ColumnIndex) = HeadingValue Then
            ArrayGetColumnNumber = ColumnIndex
            Exit Function
        End If
    Next ColumnIndex
    
    'RETURN NEGATIVE IF NOT FOUND
    ArrayGetColumnNumber = -1
    
End Function

' CHECKS TO SEE IF VALUE IS IN SINGLE DIM ARRAY
Public Function ArrayIncludes(SourceArray As Variant, ByVal Value As Variant) As Boolean
    
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

' RETURNS INDEX OF A SINGLE DIM ARRAY ELEMENT
Public Function ArrayIndexOf(SourceArray As Variant, SearchElement As Variant) As Integer
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
        If SourceArray(Index) = SearchElement Then
            ArrayIndexOf = Index
            Exit Function
        End If
    Next Index
    Index = -1
End Function

' REMOVES LAST ELEMENT IN ARRAY, RETURNS POPPED ELEMENT
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

' ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function ArrayPush(SourceArray As Variant, ParamArray Element() As Variant) As Long

    Dim Index As Long
    Dim FirstEmptyBound As Long
    Dim OptionBase As Integer
    
    OptionBase = 0

    '@TODO: FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. UPDATE TO PUSH TO MULTI DIM ARRAYS?
    If ArrayDimensionLength(SourceArray) > 1 Then
        ArrayPush = -1
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
        Asign SourceArray(FirstEmptyBound), Element(Index)
        
        'INCREMENT TO THE NEXT firstEmptyBound
        FirstEmptyBound = FirstEmptyBound + 1
        
    Next Index
    
    'RETURN NEW ARRAY LENGTH
    ArrayPush = UBound(SourceArray, 1) + 1

End Function

' CREATES TEMP TEXT FILE AND SAVES ARRAY VALUES IN A CSV FORMAT,
' THEN QUERIES AND RETURNS ARRAY.
'
'@AUTHOR ROBERT TODAR
'@USES ArrayToCSVFile
'@USES ArrayFromRecordset
'@RETURNS 2D ARRAY || EMPTY (IF NO RECORDS)
'@PARAM {ARR} MUST BE A TWO DIMENSIONAL ARRAY, SETUP AS IF IT WERE A TABLE.
'@PARAM {SQL} ADO SQL STATEMENT FOR A TEXT FILE. MUST INCLUDE 'FROM []'
'@PARAM {IncludeHeaders} BOOLEAN TO RETURN HEADERS WITH DATA OR NOT
'@EXAMPLE SQL = "SELECT * FROM [] WHERE [FIRSTNAME] = 'ROBERT'"
Public Function ArrayQuery(SourceArray As Variant, SQL As String, Optional IncludeHeaders As Boolean = True) As Variant
    
    'CREATE TEMP FOLDER AND FILE NAMES
    Const FileName As String = "temp.txt"
    Dim FilePath As String
    FilePath = Environ("temp")
    
    'UPDATE SQL WITH TEMP FILE NAME
    SQL = Replace(SQL, "FROM []", "FROM [" & FileName & "]")
    
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
        .Open SQL
        
        'GET AN ARRAY FROM THE RECORDSET
         ArrayQuery = ArrayFromRecordset(Rs, IncludeHeaders)
        .Close
    End With
    
    'CLOSE CONNECTION AND KILL TEMP FILE
    cnn.Close
    Kill FilePath & "\" & FileName
    
End Function

' REMOVED DUPLICATES FROM SINGLE DIM ARRAY
Public Function ArrayRemoveDuplicates(SourceArray As Variant) As Variant
    Dim dic As Object
    Dim Key As Variant
    
    If Not IsArray(SourceArray) Then
        SourceArray = cArray(SourceArray)
    End If
    
    Set dic = CreateObject("Scripting.Dictionary")
    For Each Key In SourceArray
        dic(Key) = 0
    Next
    ArrayRemoveDuplicates = dic.Keys
    SourceArray = ArrayRemoveDuplicates
End Function

'REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
Public Function ArrayReverse(SourceArray As Variant) As Variant
    
    Dim Temp As Variant
    
    'REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    Dim Index As Long
    For Index = LBound(SourceArray, 1) To ((UBound(SourceArray) + LBound(SourceArray)) \ 2)
        
        'STORE LAST VALUE MINUS THE ITERATION
        Asign Temp, SourceArray(UBound(SourceArray) + LBound(SourceArray) - Index)
        
        'SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        Asign SourceArray(UBound(SourceArray) + LBound(SourceArray) - Index), SourceArray(Index)
        
        'SET FIRST VALUE TO THE STORED LAST VALUE
        Asign SourceArray(Index), Temp
        
    Next Index
    
    ArrayReverse = SourceArray
    
End Function

' REMOVES ELEMENT FROM ARRAY - RETURNS REMOVED ELEMENT **[SINGLE DIMENSION]
Public Function ArrayShift(SourceArray As Variant, Optional ElementNumber As Long = 0) As Variant
    
    If Not IsArrayEmpty(SourceArray) Then

        ArrayShift = SourceArray(ElementNumber)
        
        Dim Index As Long
        For Index = ElementNumber To UBound(SourceArray) - 1
            Asign SourceArray(Index), SourceArray(Index + 1)
        Next Index
        
        ReDim Preserve SourceArray(UBound(SourceArray, 1) - 1)
        
    End If
    
End Function

' SORT AN ARRAY [SINGLE DIMENSION]
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

' CHANGES THE CONTENTS OF AN ARRAY BY REMOVING OR REPLACING EXISTING ELEMENTS AND/OR ADDING NEW ELEMENTS.
Public Function ArraySplice(SourceArray As Variant, Where As Long, HowManyRemoved As Integer, ParamArray Element() As Variant) As Variant
    
    'CHECK TO SEE THAT INSERT IS NOT GREATER THAN THE Array (REDUCE IF SO)
    If Where > UBound(SourceArray, 1) + 1 Then
        Where = UBound(SourceArray, 1) + 1
    End If
    
    'CHECK TO MAKE SURE REMOVED IS NOT MORE THAN THE Array (REDUCE IF SO)
    If HowManyRemoved > (UBound(SourceArray, 1) + 1) - Where Then
        HowManyRemoved = (UBound(SourceArray, 1) + 1) - Where
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

' BASICALY ARRAY TO STRING HOWEVER QUOTING STIRNGS, THEN SAVING TO A TEXTFILE
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
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = FSO.OpenTextFile(FilePath, 2, True) '2=WRITEABLE
    ts.Write Temp
    
    Set ts = Nothing
    Set FSO = Nothing
    
    ArrayToCSV = Temp
    
End Function

' RESIZE PASSED IN RANGE, AND SET VALUE EQUAL TO THE ARRAY
Public Sub ArrayToRange(SourceArray As Variant, Optional ByRef Target As Range)
    
    Dim Wb As Workbook
    
    If Target Is Nothing Then
        Set Wb = Workbooks.Add
        Set Target = Wb.Worksheets("Sheet1").Range("A1")
    End If
    
    Select Case ArrayDimensionLength(SourceArray)
        Case 1:
            Set Target = Target.Resize(UBound(SourceArray) - LBound(SourceArray) + 1, 1)
            Target.Value = Application.Transpose(SourceArray)
            
        Case 2:
            Target.Resize((UBound(SourceArray, 1) + 1) - LBound(SourceArray, 1), (UBound(SourceArray, 2) + 1 - LBound(SourceArray, 2))).Value = SourceArray
    
    End Select
    
    Columns.AutoFit
    
End Sub

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
'
'@AUTHOR ROBERT TODAR
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
    
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = FSO.OpenTextFile(FilePath, 2, True) '2=WRITEABLE
    ts.Write ArrayToString(Arr, delimeter)
    
    Set ts = Nothing
    Set FSO = Nothing

End Sub

' APPLICATION.TRANSPOSE HAS A LIMIT ON THE SIZE OF THE ARRAY, AND IS LIMITED TO THE 1ST DIM
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

' - ADDS NEW ELEMENT TO THE BEGINING OF THE ARRAY
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
        Asign Temp(Count), Element(Index)
        Count = Count + 1
    Next Index
    
    If Not Count > UBound(Temp, 1) Then
    
        'ADD ELEMENTS FROM ORIGINAL ARRAY
        For Index = LBound(SourceArray, 1) To UBound(SourceArray, 1)
            Asign Temp(Count), SourceArray(Index)
            Count = Count + 1
        Next Index
    End If
    
    'SET ARRAY TO TEMP ARRAY
    SourceArray = Temp
    
    'RETURN THE NEW LENGTH OF THE ARRAY
    ArrayUnShift = UBound(SourceArray, 1) + 1
    
End Function

' CONVERT OTHER LIST OBJECTS TO AN ARRAY
Public Function ConvertToArray(val As Variant) As Variant
    
    Select Case TypeName(val)
    
        Case "Collection":
            Dim Index As Integer
            For Index = 1 To val.Count
                ArrayPush cArray, val(Index)
            Next Index
        
        Case "Dictionary":
            cArray = val.items()
        
        Case Else
             
            If IsArray(val) Then
                cArray = val
            Else
                ArrayPush cArray, val
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
    Dim UB As Long
    UB = UBound(Arr, 1)
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
        If LB > UB Then
            IsArrayEmpty = True
        Else
            IsArrayEmpty = False
        End If
    End If

End Function


'******************************************************************************************
' PRIVATE FUNCTIONS - BEING DEVELOPED STILL
'******************************************************************************************

' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function Asign(variable As Variant, Value As Variant)

    If IsObject(Value) Then
        Set variable = Value
    Else
        Let variable = Value
    End If
    
End Function


