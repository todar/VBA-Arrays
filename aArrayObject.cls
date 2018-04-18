VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "aArrayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Option Compare Text

Private pArray As Variant

'DEPENDENCIES
' - must import file for Property Value to be set as the class default
' - this was the orginal version

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
' - concat *
' - sort @TODO - UPDATE FOR MULTI DIM ARRAYS
' - reverse @TODO - UPDATE FOR MULTI DIM ARRAYS
' - toString
' - toRange
' - columnNumber
' - extractColumn

'Private Helper Funtions:
' - arrayFromCollection
' - asign
' - collectionToMultiDimArray
' - collectionFromarray
' - dimensionLength
' - isSingleDimension


'NOTES:
' - THIS AN ARRAY OBJECT THAT CAN EASILY ADD\REMOVE ITEMS TO A SINGLE DIM ARRAY.
' ALSO HAS VARIOUS FILTER AND SORT FUNCTIONS FOR MULTI DIM ARRAYS
'

'******************************************************************************
' EVENTS
'******************************************************************************
Private Sub Class_Initialize()
    
    'ALLOWCATE ARRAY
    ReDim pArray(0)
    
End Sub
Private Sub Class_Terminate()
    
    'BEST PRACTICE TO EMPTY ARRAY (NOT NEEDED)
    pArray = Empty
    
End Sub


'******************************************************************************
' PUBLIC PROPERTIES
'******************************************************************************
Public Property Let value(Optional Index As Variant = -1, Arr As Variant)
Attribute value.VB_UserMemId = 0
    
    If Index > -1 Then
        
        If Index >= Me.Length Then
            push Arr
        Else
            asign pArray(Index), Arr
        End If
        
    ElseIf IsArray(Arr) Then
    
        pArray = Arr

    Else
        On Error Resume Next
        If Arr = Empty Then
            ReDim pArray(0)
        End If
    End If
    
End Property
Public Property Get value(Optional Index As Variant = -1) As Variant
    
    If Index > -1 Then
        asign value, pArray(Index)
    Else
        asign value, pArray
    End If
    
End Property

Public Property Get Length() As Long
    Length = UBound(pArray, 1) + 1
End Property

Public Property Get lowerBound(Optional bounds As Integer = 1) As Variant
    lowerBound = LBound(pArray, bounds)
End Property

Public Property Get upperBound(Optional bounds As Integer = 1) As Variant
    upperBound = UBound(pArray, bounds)
End Property


'******************************************************************************
' PUBLIC METHODS\FUNCTIONS
'******************************************************************************

' - ADDS A NEW ELEMENT TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function push(Element As Variant) As Long
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        push = -1
        Exit Function
    End If
    
    'INCREASE UBOUND IF ARRAY IS NOT EMPTY
    If pArray(0) <> Empty Then
        ReDim Preserve pArray(UBound(pArray, 1) + 1)
    End If
    
    'ADD ELEMENT TO THE END OF THE ARRAY
    asign pArray(UBound(pArray, 1)), Element

    'RETURN NEW ARRAY LENGTH
    push = UBound(pArray, 1) + 1
    
End Function

' - REMOVES LAST ELEMENT FROM ARRAY, RETURNS REMOVED ELEMENT
Public Function pop() As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        pop = -1
        Exit Function
    End If
    
    'RETURN ELEMENT THAT IS REMOVED
    asign pop, pArray(UBound(pArray, 1))
    
    'REMOVE LAST ELEMENT FROM ARRAY
    ReDim Preserve pArray(UBound(pArray, 1) - 1)
    
End Function

' - ADDS NEW ELEMENT TO THE BEGINING OF THE COLLECTION
Public Function unShift(Element As Variant) As Long
    
    Dim i As Long
    Dim Temp As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        unShift = -1
        Exit Function
    End If
    
    'RETURN THE NEW LENGTH OF THE ARRAY
    unShift = UBound(pArray, 1) + 1
    
    'RESIZE TEMP ARRAY TO PLUS ONE
    ReDim Temp(LBound(pArray, 1) To UBound(pArray, 1) + 1)
    
    'ADD NEW ELEMENT TO TEMP
    asign Temp(LBound(Temp, 1)), Element
    
    'ADD ELEMENTS FROM ARRAY TO TEMP
    For i = LBound(pArray, 1) To UBound(pArray, 1)
  
        asign Temp(i + 1), pArray(i)

    Next i
    
    'SET ARRAY TO TEMP ARRAY
    pArray = Temp
    
End Function

' - REMOVES FIRST ELEMENT FROM THE BEGINING OF THE COLLECTION
Public Function shift() As Variant
    
    Dim i As Long
    Dim Temp As Variant
    
    'FOR NOW THIS IS ONLY FOR SINGLE DIMENSIONS. @TODO: UPDATE TO PUSH TO MULTI DIM ARRAYS
    If Not isSingleDimension Then
        shift = -1
        Exit Function
    End If
    
    'RETURN ELEMENT THAT IS BEING REMOVED
    shift = pArray(LBound(pArray, 1))
    
    'RESIZE TEMP ARRAY TO MINUS ONE
    ReDim Temp(LBound(pArray, 1) To UBound(pArray, 1) - 1)
    
    'FILL TEMP ARRAY WITH REMAINING ARRAY ELEMENTS
    For i = LBound(pArray, 1) + 1 To UBound(pArray, 1)
        
        asign Temp(i - 1), pArray(i)

    Next i
    
    'SET ARRAY TO TEMP ARRAY
    pArray = Temp
    
End Function


' - RETURNS A NEW ARRAY FILTERED BASED ON PARAMETERS
Public Function filter(Match As Variant, Optional column As Variant = 1, Optional includeHeader As Boolean = True) As Variant
    
    Dim i As Integer
    Dim i2 As Integer
    Dim Arr As Variant
    Dim Count As Integer
    
    Dim item As New Collection
    Dim rowItems As Collection
    
    'USER CAN SEARCH FOR THE COLUMN BY THE NAME
    If Not IsNumeric(column) Then column = columnNumber(CStr(column))
    
    If isSingleDimension Then
        
        'VBA HAS FUNCTION ALREADY FOR SINGLE DIMENSION ARRAYS
        filter = VBA.filter(pArray, Match, , vbTextCompare)
    Else
        
        'LOOP EVERY ROW OF MULTIDIMINSION ARRAY IN PARAMETER COLUMN
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            
            
            'Debug.Print Evaluate("""" & pArray(i, column) & """" & "= ""gregarious""")
            
            'IF MATCH IS FOUND, GET STORE ROW DATA IN COLLECTION, ADD TO MAIN COLLECTION
            If (CStr(pArray(i, column)) = CStr(Match)) Or (includeHeader = True And i = LBound(pArray, 1)) Then
                
                'CREATE A NEW ROW IN THE COLLECTION
                Set rowItems = New Collection
                
                'LOOP ARRAY ROW ADDING ELEMENTS TO THE NEW ARRAY
                For i2 = LBound(pArray, 2) To UBound(pArray, 2)
                    rowItems.Add pArray(i, i2)
                Next i2
                
                'ADD ROW TO THE LIST COLLECTION
                item.Add rowItems
             
            End If
        Next i
        
        'SET FILTER USING FUNCTION THAT RETURNS A MULTIDIMENSION ARRAY FROM A MULTIDIMENSION COLLECTION
        filter = collectionToMultiDimArray(item)
   
    End If
    
End Function

' - RETURNS SINGLE COLUMN FROM MULTI DIMENSTION ARRAY
Public Function extractColumn(Name As String) As Variant
    
    Dim i As Integer
    Dim Col As Integer
    Dim List As New Collection
    
    If Not isSingleDimension Then
        
        Col = columnNumber(Name)
        ReDim Arr(UBound(pArray, 1))
        
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            List.Add pArray(i, Col)
        Next i
        
    End If
    
    extractColumn = arrayFromCollection(List)
    
End Function

' - RETURNS COLUMN NUMBER FROM A MULTIDIM ARRAY
Public Function columnNumber(Name As String) As Integer
    
    Dim i As Integer
    
    If Not isSingleDimension Then
        For i = LBound(pArray, 2) To UBound(pArray, 2)
            If pArray(LBound(pArray, 1), i) Like Name Then
                columnNumber = i
                Exit Function
            End If
        Next i
    End If
    
End Function

' - CONVERTS AN ARRAY TO A STRING OF (COMMA SEPARATED) ARRAY VALUES
Public Function toString(Optional Delimiter As String = ",") As String
    
    Dim i As Integer
    Dim i2 As Integer
    Dim Arr As Variant
    Dim S As String
    
    On Error Resume Next
    If isSingleDimension Then
              
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            If IsObject(pArray(i)) Then
                toString = toString & "[" & TypeName(pArray(i)) & "]" & Delimiter
            Else
                toString = toString & pArray(i) & Delimiter
            End If
            
            toString = toString & vbNewLine
        Next i
        
    Else
    
        'LOOP ROWS
        For i = LBound(pArray, 1) To UBound(pArray, 1)
            
            'LOOP COL
            For i2 = LBound(pArray, 2) To UBound(pArray, 2)
            
                If IsObject(pArray(i, i2)) Then
                    toString = toString & "[" & TypeName(pArray(i, i2)) & "]" & Delimiter
                Else
                    toString = toString & pArray(i, i2) & Delimiter
                End If
                
            Next i2
            
            toString = toString & vbNewLine
            
        Next i

    End If

End Function

Public Function forEach(macroName As String, Optional WorkbookName As String)
    
    Dim i As Integer
    Dim Temp As String
    
    If WorkbookName = "" Then
        WorkbookName = ThisWorkbook.Name
    End If
        
    WorkbookName = "'" & WorkbookName & "'"
        
    Temp = WorkbookName & "!" & macroName
    
    For i = LBound(pArray, 1) To UBound(pArray, 1)
        Application.Run Temp, pArray(i)
    Next i
    

End Function

'LOOPS THROUGH EACH ELEMENT, RUNNINGS EXCELS EVALUATE FUNCTION FROM THE FORMULA STRING
Public Function map(Formula As String, Optional Element As String = "${}") As Variant
    
    Dim i As Integer
    Dim Arr As Variant
    Dim S As String
    
    Arr = pArray
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        If IsObject(Arr(i)) Then GoTo Nxt
        
        If TypeName(Arr(i)) = "String" Then Arr(i) = """" & Arr(i) & """"
        If InStr(Formula, Element) > 0 Then
            S = Replace(Formula, Element, Arr(i))
        Else
            S = S & Arr(i)
        End If
        
        Arr(i) = Evaluate(S)
Nxt:
    Next i
    
    map = Arr
    
End Function

'SORT THE ARRAY. NO RETURN, EFFECTS pArray
Public Sub Sort()
    
    Dim i As Long
    Dim j As Long
    Dim Temp As Variant
    
    '@CHECK - DOES NOT SUPPORT OBJECTS - SKIP ON ERROR FOR NOW
    On Error Resume Next
    
    'SORT THE ARRAY A-Z
    For i = LBound(pArray) To UBound(pArray) - 1
        For j = i + 1 To UBound(pArray)
            If UCase(pArray(i)) > UCase(pArray(j)) Then
                Temp = pArray(j)
                pArray(j) = pArray(i)
                pArray(i) = Temp
            End If
        Next j
    Next i

End Sub

'REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
Public Sub reverse()
    
    Dim Temp As Variant
    Dim i As Integer
    
    'REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    For i = LBound(pArray, 1) To ((UBound(pArray) - LBound(pArray)) \ 2)
        
        'STORE LAST VALUE MINUS THE ITERATION
        asign Temp, pArray(UBound(pArray) - i)
        
        'SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        asign pArray(UBound(pArray) - i), pArray(i)
        
        'SET FIRST VALUE TO THE STORED LAST VALUE
        asign pArray(i), Temp
        
    Next i
    
End Sub



'RETURN A NEW ARRAY, JOINS CURRENT ARRAY WITH ANOTHER ONE.
Public Function concat(Arr) As Variant
    
    Dim i As Integer
    Dim Temp As Variant
    
    Temp = pArray
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        
        ReDim Preserve Temp(UBound(Temp, 1) + 1)
        asign Temp(UBound(Temp, 1)), Arr(i)
   
    Next i
    
    concat = Temp
    
End Function



Public Sub toRange(target As Range)
    
    If isSingleDimension Then
        Set target = target.Resize(UBound(pArray), 1)
        target.value = Application.Transpose(pArray)
    Else
        target.Resize(UBound(pArray, 1) + 1, UBound(pArray, 2) + 1).value = pArray
    End If
End Sub


'******************************************************************************
' PRIVATE METHODS\FUNCTIONS         (HELPER FUNCTIONS)
'******************************************************************************


Private Function collectionToMultiDimArray(C As Collection) As Variant
    
    Dim Arr As Variant
    Dim i As Integer
    Dim i2 As Integer
    
    On Error GoTo catch
    ReDim Arr(0 To C.Count - 1, 0 To C.item(1).Count - 1)
    
    For i = 1 To C.Count
        For i2 = 1 To C.item(1).Count
            asign Arr(i - 1, i2 - 1), C.item(i).item(i2)
        Next i2
    Next i
    
    collectionToMultiDimArray = Arr
    
catch: 'IF ERROR, THAT MEANS COLLECTION WAS EMPTY. RETURNS EMPTY
    
End Function

Private Function arrayFromCollection(Col As Collection) As Variant
    
    Dim Arr As Variant
    Dim i As Long
    
    ReDim Arr(0 To Col.Count - 1)
    
    For i = 1 To Col.Count
        asign Arr(i - 1), Col(i)
    Next i
    
    arrayFromCollection = Arr
    
End Function

Private Function collectionFromarray(Arr As Variant) As Collection
    
    Dim Col As New Collection
    Dim i As Integer
    
    For i = LBound(Arr, 1) To UBound(Arr, 1)
        Col.Add Arr(i)
    Next i
    
    Set collectionFromarray = Col
    
End Function

' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function asign(Arr As Variant, Val As Variant)

    If IsObject(Val) Then
        Set Arr = Val
    Else
        Let Arr = Val
    End If
    
End Function

Private Function isSingleDimension() As Boolean
    If dimensionLength(pArray) = 1 Then isSingleDimension = True
End Function

' - RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Private Function dimensionLength(Arr As Variant) As Integer
    
    Dim i As Integer
    Dim test As Integer
    
    On Error Resume Next
    Do
        i = i + 1
        test = UBound(Arr, i)
    Loop Until Err.Number <> 0
    
    dimensionLength = i - 1

End Function







