VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "arrayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Text

Private pArray As Variant
Private Const optionBase As Integer = 0 'SET TO EITHER 0 OR 1

'DEPENDENCIES
' - must import file for Property Value to be set as the class default.
' - no extra refrences to other libriaries needed at this time.

'Properties:
' - value
' - lenght

'Public Funtions:
' - push
' - pop
' - shift
' - unshift
' - filter
' - map
' - forEach
' - exists
' - reduce
' - concat
' - sort
' - reverse
' - toString
' - toRange
' - columnNumber
' - splice
' - returnColumn
' - returnRow

'Private Helper Funtions:
' - arrayFromCollection
' - asign
' - collectionToMultiDimArray
' - collectionFromarray
' - dimensionLength
' - isSingleDimension
' - array2dUnshift
' - sqlArray

'NOTES:
' - THIS AN ARRAY OBJECT THAT CAN EASILY ADD\REMOVE ITEMS TO A SINGLE DIM ARRAY.
'   ALSO HAS VARIOUS FILTER AND SORT FUNCTIONS FOR MULTI DIM ARRAYS

'TODO:
' - UPDATE (SORT, REVERSE, CONCAT) TO BE USED FOR MULTI DIM ARRAYS
' - UPDATE FILTER, TO BETTER WORK WITH MULTIPLE FILTERS
' - REVERSE NEED OPTION TO NOT REVERSE HEADINGS
' - ADD FUNCTIONS: removeDuplicates, slice, indexOf, lastIndexOf


'******************************************************************************
' EVENTS
'******************************************************************************
Private Sub Class_Initialize()
    
    'ALLOWCATE ARRAY
    ReDim pArray(optionBase)
    
End Sub
Private Sub Class_Terminate()
    
    'BEST PRACTICE TO EMPTY ARRAY (NOT NEEDED)
    pArray = Empty
    
End Sub


'******************************************************************************
' PUBLIC PROPERTIES
'******************************************************************************

' - VALUE IS TO GET AND SET THE ARRAY. IT IS SET BY DEFAULT WHEN CLASS IS IMPORTED.
Public Property Let Value(Optional Index As Variant = -1, Arr As Variant)
Attribute Value.VB_UserMemId = 0
    
    'IF SPECIFIC INDEX IS PASSED IN
    If Index > -1 Then
        
        'IF INDEX IS LARGER THAN THE UBOUND, THEN PUSH THE ITEM TO THE END
        If Index >= Me.Length Then
            push Arr
            
        'OTHERWISE, REPLACE THAT INDEXES ELEMENT
        Else
            asign pArray(Index), Arr
        End If
    
    'IF NOT A SINGLE ITEM, CHECK TO SEE IF ARRAY WAS PASSED IN
    ElseIf IsArray(Arr) Then
    
        pArray = Arr
    
    'OTHERWISE CHECK TO SEE IF ARRAY WAS SET TO EMPTY
    Else
        On Error Resume Next
        If Arr = Empty Then
            ReDim pArray(optionBase)
        End If
    End If
    
End Property
Public Property Get Value(Optional Index As Variant = -1) As Variant
    
    'IF INDEX IS PASSED, RETURN ELEMENT, OTHERWISE RETURN ARRAY
    If Index > -1 Then
        asign Value, pArray(Index)
    Else
        asign Value, pArray
    End If
    
End Property

' - LENGTH IS THE NUMBER OF ELEMENTS
Public Property Get Length() As Long
    Length = (UBound(pArray, 1) + 1) - optionBase
End Property

' - HELPER BOUND PROPERITES
Public Property Get lowerBound(Optional bounds As Integer = 1) As Variant
    lowerBound = LBound(pArray, bounds)
End Property
Public Property Get upperBound(Optional bounds As Integer = 1) As Variant
    upperBound = UBound(pArray, bounds)
End Property


'******************************************************************************
' PUBLIC METHODS\FUNCTIONS
'******************************************************************************

' - ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function push(ParamArray Element() As Variant) As Long

    Dim i As Integer
    Dim ub As Long

    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS?
    If Not isSingleDimension Then
        push = -1
        Exit Function
    End If
    
    'CAPTURE THE CURRENT UBOUND
    ub = UBound(pArray, 1)
    
    'INCREASE IF ARRAY INDEX IS NOT EMPTY
    If pArray(UBound(pArray, 1)) <> Empty Then
        ReDim Preserve pArray(UBound(pArray, 1) + UBound(Element, 1) + 1)
        ub = ub + 1
    Else
        ReDim pArray(optionBase To UBound(Element, 1) + optionBase)
    End If
    
    'LOOP EACH NEW ELEMENT
    For i = LBound(Element, 1) To UBound(Element, 1)
        
        'ADD ELEMENT TO THE END OF THE ARRAY
        asign pArray(ub), Element(i)
        
        'INCREMENT UBOUND
        ub = ub + 1
        
    Next i
    
    'RETURN NEW ARRAY LENGTH
    push = Length

End Function

' - REMOVES LAST ELEMENT FROM ARRAY, RETURNS REMOVED ELEMENT
Public Function pop() As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS?
    If Not isSingleDimension Then
        pop = -1
        Exit Function
    End If
    
    'RETURN ELEMENT THAT IS REMOVED
    asign pop, pArray(UBound(pArray, 1))
    
    'REMOVE LAST ELEMENT FROM ARRAY
    ReDim Preserve pArray(UBound(pArray, 1) - 1)
    
End Function

' - ADDS NEW ELEMENT TO THE BEGINING OF THE ARRAY
Public Function unShift(ParamArray Element() As Variant) As Long
    
    Dim i As Long
    Dim count As Long
    Dim temp As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        unShift = -1
        Exit Function
    End If
    
    'RESIZE TEMP ARRAY
    If pArray(UBound(pArray, 1)) <> Empty Then
        ReDim temp(UBound(pArray, 1) + UBound(Element, 1) + 1)
    Else
        ReDim temp(optionBase To UBound(Element, 1) + optionBase)
    End If
    
    count = LBound(temp, 1)
    
    'ADD ELEMENTS TO TEMP ARRAY
    For i = LBound(Element, 1) To UBound(Element, 1)
        asign temp(count), Element(i)
        count = count + 1
    Next i
    
    If Not count > UBound(temp, 1) Then
    
        'ADD ELEMENTS FROM ORIGINAL ARRAY
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            asign temp(count), pArray(i)
            count = count + 1
        Next i
    End If
    
    'SET ARRAY TO TEMP ARRAY
    pArray = temp
    
    'RETURN THE NEW LENGTH OF THE ARRAY
    unShift = Me.Length
    
End Function


' - REMOVES FIRST ELEMENT FROM THE BEGINING OF THE COLLECTION
Public Function Shift() As Variant
    
    Dim i As Long
    Dim temp As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        Shift = -1
        Exit Function
    End If
    
    'RETURN ELEMENT THAT IS BEING REMOVED
    Shift = pArray(LBound(pArray, 1))
    
    'RESIZE TEMP ARRAY TO MINUS ONE
    ReDim temp(LBound(pArray, 1) To UBound(pArray, 1) - 1)
    
    'FILL TEMP ARRAY WITH REMAINING ARRAY ELEMENTS
    For i = LBound(pArray, 1) + 1 To UBound(pArray, 1)
        
        asign temp(i - 1), pArray(i)

    Next i
    
    'SET ARRAY TO TEMP ARRAY
    pArray = temp
    
End Function

' - RETURNS A NEW ARRAY QUERIED BASED ON PARAMETERS
Public Function query(SQL As String) As Variant

    If Not isSingleDimension Then

        'CALL TO PRIVATE HELPER FUNCTION
        query = sqlArray(pArray, SQL)

    End If

End Function

'WILL BE USED TO ADD NEW ITEMS TO AN ARRAY, AS WELL AS REMOVE ITEMS FROM THAT LOCATION
Public Sub splice(Where As Long, howManyRemoved As Integer, ParamArray Element() As Variant)
    
    Dim i As Long
    Dim i2 As Long
    Dim i3 As Long
    Dim Arr As Variant
    Dim visited As Boolean
    
    'CHECK TO SEE THAT INSERT IS NOT GREATER THAN THE ARRAY
    If Where > UBound(pArray, 1) + 1 Then
        Where = UBound(pArray, 1) + 1
    End If
    
    'CHECK TO MAKE SURE REMOVED IS NOT MORE THAN THE ARRAY
    If howManyRemoved > (UBound(pArray, 1) + 1) - Where Then
        howManyRemoved = (UBound(pArray, 1) + 1) - Where
    End If
    
    ReDim Arr(LBound(pArray, 1) To UBound(pArray, 1) + UBound(Element, 1) + 1 - howManyRemoved)
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        If i = Where And visited = False Then
            
            visited = True
            
            'ADD NEW ELEMENTS
            For i2 = LBound(Element, 1) To UBound(Element, 1)
                Arr(i) = Element(i2)
                i3 = i3 + 1
                'If i2 < UBound(Element, 1) Then i = i + 1
                i = i + 1
            Next i2
            
            i = i - 1
            i3 = i3 - howManyRemoved
        
        Else
            Arr(i) = pArray(i - i3)
        End If
        
    Next i
    
    pArray = Arr
    
End Sub

Public Function returnRow(Index As Variant) As Variant

    'RETURN AS A 2-DIMENSIONAL ARRAY, SO THAT IT CAN BE CONCATINATED IF NEEDED

    Dim i As Integer
    Dim Col As Integer
    Dim temp As Variant

    'MUST BE MULTI DIMENSIONED ARRAY
    If Not isSingleDimension Then

        'REDIM TEMP ARR AND SET EQUAL TO ROW
        ReDim temp(optionBase To optionBase, LBound(pArray, 2) To UBound(pArray, 2))

        For i = LBound(pArray, 2) To UBound(pArray, 2)
            temp(optionBase, i) = pArray(Index, i)
        Next i

    End If

    'RETURN COLUMN
    returnRow = temp

End Function

' - RETURNS SINGLE COLUMN FROM MULTI DIMENSTIONAL ARRAY
Public Function returnColumn(name As String) As Variant

    Dim i As Integer
    Dim Col As Integer
    Dim temp As Variant

    'MUST BE MULTI DIMENSIONED ARRAY
    If Not isSingleDimension Then

        'GET COLUMN NUMBER USING FUNCTION
        Col = columnNumber(name)

        'REDIM TEMP ARR AND SET EQUAL TO COLUMN
        ReDim temp(UBound(pArray, 1))

        For i = LBound(pArray, 1) To UBound(pArray, 1)
            temp(i) = pArray(i, Col)
        Next i

    End If

    'RETURN COLUMN
    returnColumn = temp

End Function


' - RETURNS COLUMN NUMBER FROM A MULTIDIM ARRAY
Public Function columnNumber(name As String) As Integer

    Dim i As Integer

    'MUST BE MULTI DIMENSIONED ARRAY
    If Not isSingleDimension Then

        'LOOP THE COLUMNS CHECKING FIRST ROW (HEADING) FOR MATCHES
        For i = LBound(pArray, 2) To UBound(pArray, 2)

            'IF MATCH FOUND, RETURN THAT COLUMN NUMBER
            If pArray(LBound(pArray, 1), i) Like name Then
                columnNumber = i
                Exit Function
            End If

        Next i
    End If

End Function

' - CONVERTS AN ARRAY TO A STRING OF (COMMA SEPARATED) ARRAY VALUES
Public Function toString(Optional Delimiter As String = ",") As String
    
    Dim i As Long 'long in case there are many rows
    Dim i2 As Integer
    Dim Arr As Variant
    Dim s As String
    
    On Error Resume Next
    If isSingleDimension Then
           
        'SINGLE DIMENSION ARRAYS
        For i = LBound(pArray, 1) To UBound(pArray, 1)
        
            'OBJECTS WILL PRINT OUT THEIR TYPE, EXAMPLE: [DICTIONARY]
            If IsObject(pArray(i)) Then
                toString = toString & "[" & TypeName(pArray(i)) & "]" & Delimiter
            
            'NON-OBJECTS
            Else
                toString = toString & pArray(i) & Delimiter
            End If
            
            'ADD NEW LINE FOR NEXT ELEMENT
            'toString = toString & vbNewLine
        Next i
        
        toString = Left(toString, Len(toString) - Len(Delimiter))
        
    Else
    
        'MULTI DIMENSION - OUTER LOOP IS ROWS
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            
            'INNER LOOP IS COLUMNS
            For i2 = LBound(pArray, 2) To UBound(pArray, 2)
                
                'OBJECTS WILL PRINT OUT THEIR TYPE, EXAMPLE: [DICTIONARY]
                If IsObject(pArray(i, i2)) Then
                    toString = toString & "[" & TypeName(pArray(i, i2)) & "]" & Delimiter
                
                'NON-OBJECTS
                Else
                    toString = toString & pArray(i, i2) & Delimiter
                End If
                
            Next i2
            
            'ADD NEW LINE FOR NEXT ELEMENT
            toString = toString & vbNewLine
            
        Next i

    End If

End Function

Public Function exists(Value As Variant) As Boolean
    
    '@TODO: ADD ERROR HANDLING AND OPTION FOR 2D ARRAY
    Dim i As Long
    
    For i = LBound(pArray, 1) To UBound(pArray, 1)
        If pArray(i) Like Value Then
            exists = True
            Exit Function
        End If
    Next i

End Function

' - MUTATE ORIGINAL ARRAY
Public Sub forEach(Formula As String, Optional Condition As String, Optional Element As String = "x")
    
    pArray = Me.map(Formula, Condition, Element)

End Sub

' - LOOPS THROUGH EACH ELEMENT, RUNNINGS EXCELS EVALUATE FUNCTION FROM THE FORMULA STRING
Public Function map(Formula As String, Optional Condition As String, Optional Element As String = "x") As Variant
    
    Dim i As Long
    Dim Arr As Variant
    Dim temp As String
    Dim s As String
    
    Arr = pArray
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        If IsObject(Arr(i)) Then GoTo Nxt
        
        temp = pArray(i)
        If TypeName(temp) = "String" Then temp = """" & temp & """"
        
        s = Replace(Condition, Element, temp)
        If s = "" Then s = True
        If Evaluate(s) = True Then
        
            If InStr(Formula, Element) > 0 Then
                s = Replace(Formula, Element, temp)
            Else
                s = s & temp
            End If
            
            Arr(i) = Evaluate(s)
  
        End If
       
Nxt:
    Next i
    
    map = Arr
    
End Function

Public Function reduce(Formula As String, Optional Condition As String, Optional Element As String = "x", Optional Accumulator As String = "y") As Double
    
    Dim i As Long
    Dim s As String
    
    reduce = 0
    
    For i = LBound(pArray, 1) To UBound(pArray, 1)
        
        s = Replace(Condition, Element, pArray(i))
        If s = "" Then s = True
        If Evaluate(s) = True Then
            
            s = Replace(Formula, Element, pArray(i))
            s = Replace(s, Accumulator, reduce)
            reduce = Evaluate(s)
            
        End If
        
    Next i
    
End Function

Public Function filter(Formula As String, Optional Element As String = "x") As Variant
    
    Dim i As Long
    Dim s As Variant
    Dim temp As Variant
    Dim count As Long
    
    ReDim temp(optionBase)
    
    For i = LBound(pArray, 1) To UBound(pArray, 1)
        
        s = pArray(i)
        If TypeName(s) = "String" Then s = """" & s & """"
        s = Replace(Formula, Element, s)
        
        If Evaluate(s) = True Then
            If count > 0 Then
                ReDim Preserve temp(UBound(temp) + 1)
            End If
            
            temp(UBound(temp)) = pArray(i)
            
            count = count + 1
        End If
            
    Next i
    
    pArray = temp
    
End Function

'SORT THE ARRAY. NO RETURN, EFFECTS pArray
Public Sub Sort()
    
    Dim i As Long
    Dim j As Long
    Dim temp As Variant
    
    '@CHECK - DOES NOT SUPPORT OBJECTS - SKIP ON ERROR FOR NOW
    On Error Resume Next
    
    'SORT THE ARRAY A-Z
    For i = LBound(pArray) To UBound(pArray) - 1
        For j = i + 1 To UBound(pArray)
            If UCase(pArray(i)) > UCase(pArray(j)) Then
                temp = pArray(j)
                pArray(j) = pArray(i)
                pArray(i) = temp
            End If
        Next j
    Next i

End Sub

'REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
Public Sub reverse()
    
    Dim temp As Variant
    Dim i As Long
    
    'REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    For i = LBound(pArray, 1) To ((UBound(pArray) + LBound(pArray)) \ 2)
        
        'STORE LAST VALUE MINUS THE ITERATION
        asign temp, pArray(UBound(pArray) + LBound(pArray) - i)
        
        'SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        asign pArray(UBound(pArray) + LBound(pArray) - i), pArray(i)
        
        'SET FIRST VALUE TO THE STORED LAST VALUE
        asign pArray(i), temp
        
    Next i
    
    
End Sub


'RETURN A NEW ARRAY, JOINS CURRENT ARRAY WITH ANOTHER ONE.
Public Function concat(Arr) As Variant
    
    Dim i As Long
    Dim temp As Variant
    
    temp = pArray
    
    If dimensionLength(pArray) = dimensionLength(Arr) Then
    
        For i = LBound(Arr, 1) To UBound(Arr, 1)
        
            ReDim Preserve temp(UBound(temp, 1) + 1)
            asign temp(UBound(temp, 1)), Arr(i)
    
        Next i
        
        concat = temp
    
    End If
    
End Function

' - RESIZE PASSED IN RANGE, AND SET VALUE EQUAL TO THE ARRAY
Public Sub toRange(target As Range)
    
    If isSingleDimension Then
        Set target = target.Resize(UBound(pArray), 1)
        target.Value = Application.Transpose(pArray)
    Else
        target.Resize(UBound(pArray, 1), UBound(pArray, 2)).Value = pArray
    End If
    
    Columns.AutoFit
    
End Sub


'******************************************************************************
' PRIVATE METHODS\FUNCTIONS         (HELPER FUNCTIONS)
'******************************************************************************

' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function asign(Arr As Variant, Val As Variant)

    If IsObject(Val) Then
        Set Arr = Val
    Else
        Let Arr = Val
    End If
    
End Function

' - RETURNS TRUE IF SINGLE DIMENSION
Private Function isSingleDimension() As Boolean
    If dimensionLength(pArray) = 1 Then isSingleDimension = True
End Function

' - RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Private Function dimensionLength(Arr As Variant) As Integer
    
    Dim i As Integer
    Dim Test As Integer
    
    On Error Resume Next
    Do
        i = i + 1
        Test = UBound(Arr, i)
    Loop Until Err.Number <> 0
    
    dimensionLength = i - 1

End Function

' - COLLECTION MUST BE STORE A COLLECTION OF ROW DATA
Private Function collectionToMultiDimArray(C As Collection) As Variant
    
    Dim Arr As Variant
    Dim i As Integer
    Dim i2 As Integer
    Dim lb As Integer
    Dim minus As Integer
    
    On Error GoTo catch
    
    'ADJUST VARAIABLES FOR OPTION BASE
    If optionBase = 0 Then
        minus = 1
    ElseIf optionBase = 1 Then
        lb = 1
    Else
        Err.Raise , , "Option base must be 0 or 1 for Array Objects"
        Exit Function
    End If
    
    'REDIM ARR TO THE SIZE OF THE COLLECTION
    ReDim Arr(lb To C.count - minus, lb To C.item(1).count - minus)
    
    'GET ARRAY FROM COLLECTION
    For i = 1 To C.count
        For i2 = 1 To C.item(1).count
            asign Arr(i - minus, i2 - minus), C.item(i).item(i2)
        Next i2
    Next i
    
    collectionToMultiDimArray = Arr
    
catch: 'IF ERROR, THAT MEANS COLLECTION WAS EMPTY. RETURNS EMPTY
    
End Function

' - RETURN SINGLE DIMENSION ARRAY FROM COLLECTION (NOT IN USE CURRENTLY)
Private Function arrayFromCollection(Col As Collection) As Variant
    
    Dim Arr As Variant
    Dim i As Long
    
    ReDim Arr(0 To Col.count - 1)
    
    For i = 1 To Col.count
        asign Arr(i - 1), Col(i)
    Next i
    
    arrayFromCollection = Arr
    
End Function

' - RETURN COLLECTION FROM A SINGLE DIMENSTION ARRAY (NOT IN USE CURRENTLY)
Private Function collectionFromarray(Arr As Variant) As Collection
    
    Dim Col As New Collection
    Dim i As Integer
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        Col.Add Arr(i)
    Next i
    
    Set collectionFromarray = Col
    
End Function




Private Function sqlArray(Arr As Variant, SQL As String) As Variant
    
    Dim cn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim fields As Variant
    Dim data As Variant
    Dim rng As Range
    Dim connectionString As String
    Dim i As Integer
    
    Set cn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    'CREATE WORKBOOK TO ACT AS THE TABLE TO QUERY FROM
    Application.ScreenUpdating = False
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    
    'ARRAY TO THE RANGE
    Set rng = ws.Range("A1").Resize(UBound(Arr, 1), UBound(Arr, 2))
    rng.Value = Arr
    
    'SAVE WORKBOOK TO TEMP FILE
    Application.DisplayAlerts = False
    wb.SaveAs Environ("TEMP") & "\" & "tempdb.xlsx", AccessMode:=xlExclusive, ConflictResolution:=Excel.XlSaveConflictResolution.xlLocalSessionChanges
    Application.DisplayAlerts = True
    
    'CREATE CONNECTION STRING TO THE TEMP FILE
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & wb.FullName _
        & ";Extended Properties=""Excel 12.0;HDR=Yes;IMEX=1"";"
    
    'OPEN CONNECTION TO THE TEMP FILE
    cn.Open connectionString
    
    'CREATE THE SQL TO BE RUN ON DATA. CURRENTLY HARDCODED SELECT..\FROM..\'WHERE' STATEMENT
    SQL = "SELECT * FROM [" & ws.name & "$" & rng.Address(False, False) & "] WHERE " & SQL
    
    On Error GoTo catch
    'PULL DATA BY OPENING RECORDSET
    rs.Open SQL, cn
    
    'GET HEADINGS FROM RECORDSET
    ReDim fields(LBound(Arr, 2) To UBound(Arr, 2))
    
    For i = LBound(Arr, 2) To UBound(Arr, 2)
        fields(i) = Arr(LBound(Arr, 1), i)
    Next i
    
    
    'GET ROW DATA (RETURNS TRANSPOSED, SO RE-TRANSPOSE)
    data = Application.WorksheetFunction.Transpose(rs.GetRows)
    
    'RECORDSET.GETROWS ARE TRANSPOSED. TRANSPOSE TO GET BACK TO DEFAULT TABLE STYLE
    sqlArray = array2dUnshift(fields, data)
    
catch:
    'CLOSE WB AND DELETE
    wb.Close False
    Application.ScreenUpdating = True
    Kill Environ("TEMP") & "\" & "tempdb.xlsx"

End Function


Private Function array2dUnshift(Arr1D As Variant, Arr2D As Variant) As Variant

    Dim result As Variant
    Dim iRow As Long
    Dim iCol As Long
    Dim count As Long
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Redim the Result array to have number of rows equal to
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ReDim result(LBound(Arr2D, 1) To UBound(Arr2D, 1) + 1, LBound(Arr2D, 2) To UBound(Arr2D, 2))
    
    '''''''''''''''''''''''''''''''''''''''''''''
    ' Add heading to new array
    ''''''''''''''''''''''''''''''''''''''''''''
    For iCol = LBound(Arr1D, 1) To UBound(Arr1D, 1)
        result(LBound(Arr2D, 1), iCol) = Arr1D(iCol)
    Next iCol

    '''''''''''''''''''''''''''''''''''''''''''''
    ' ADD DATA TO NEW ARRAY
    '''''''''''''''''''''''''''''''''''''''''''''
    For iRow = LBound(Arr2D, 1) To UBound(Arr2D, 1)
        
        For iCol = LBound(Arr2D, 2) To UBound(Arr2D, 2)
            result(iRow + 1, iCol) = Arr2D(iRow, iCol)
        Next iCol
        
    Next iRow
   
    '''''''''''''''''''''''''
    ' Return the Result
    '''''''''''''''''''''''''
    array2dUnshift = result

End Function





'WILL BE USED TO REMOVE ITEMS
Private Sub slice(startIndex As Long, endIndex As Long)

    Dim i As Long
    Dim Arr As Variant
    
    ReDim Arr(LBound(pArray, 1) To UBound(pArray, 1) + UBound(Element, 1) + 1 - howManyRemoved)
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        If i = Where Then
            
        Else
            Arr(i) = pArray(i)
        End If
        
    Next i
    
    pArray = Arr
    
End Sub



'Public Function forEach(macroName As String, Optional WorkbookName As String)
'
'    Dim i As Long
'    Dim temp As String
'
'    If WorkbookName = "" Then
'        WorkbookName = ThisWorkbook.name
'    End If
'
'    WorkbookName = "'" & WorkbookName & "'"
'
'    temp = WorkbookName & "!" & macroName
'
'    For i = LBound(pArray, 1) To UBound(pArray, 1)
'        Application.Run temp, pArray(i)
'    Next i
'
'
'End Function

' - RETURNS A NEW ARRAY FILTERED BASED ON PARAMETERS
'Public Function filter(match As Variant, Optional Column As Variant = 1, Optional includeHeader As Boolean = True) As Variant
'
'    Dim i As Integer
'    Dim i2 As Integer
'    Dim Arr As Variant
'    Dim Count As Integer
'
'    Dim item As New Collection
'    Dim rowItems As Collection
'
'    'USER CAN SEARCH FOR THE COLUMN BY THE NAME
'    If Not IsNumeric(Column) Then Column = columnNumber(CStr(Column))
'
'    If isSingleDimension Then
'
'        'VBA HAS FUNCTION ALREADY FOR SINGLE DIMENSION ARRAYS
'        filter = VBA.filter(pArray, match, , vbTextCompare)
'    Else
'
'        'LOOP EVERY ROW OF MULTIDIMINSION ARRAY IN PARAMETER COLUMN
'        For i = LBound(pArray, 1) To UBound(pArray, 1)
'
'
'            'Debug.Print Evaluate("""" & pArray(i, column) & """" & "= ""gregarious""")
'
'            'IF MATCH IS FOUND, GET STORE ROW DATA IN COLLECTION, ADD TO MAIN COLLECTION
'            If (CStr(pArray(i, Column)) = CStr(match)) Or (includeHeader = True And i = LBound(pArray, 1)) Then
'
'                'CREATE A NEW ROW IN THE COLLECTION
'                Set rowItems = New Collection
'
'                'LOOP ARRAY ROW ADDING ELEMENTS TO THE NEW ARRAY
'                For i2 = LBound(pArray, 2) To UBound(pArray, 2)
'                    rowItems.Add pArray(i, i2)
'                Next i2
'
'                'ADD ROW TO THE LIST COLLECTION
'                item.Add rowItems
'
'            End If
'        Next i
'
'        'SET FILTER USING FUNCTION THAT RETURNS A MULTIDIMENSION ARRAY FROM A MULTIDIMENSION COLLECTION
'        filter = collectionToMultiDimArray(item)
'
'    End If
'
'End Function


