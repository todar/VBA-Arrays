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

'Properties:
' - value ^
' - lenght

'Public Funtions:
' - push ^
' - pop ^
' - shift
' - unshift
' - filter ^
' - map
' - concat
' - sort
' - reverse
' - toString ^
' - join
' - columnNumber ^
' - getColumn ^

'Private Helper Funtions:
' - asign ^
' - incrementIfNeeded ^
' - collectionToMultiDimArray ^
' - dimensionLength ^
' - isSingleDimension ^
' - arrayFromCollection ^


'********************************************************************
' EVENTS
'********************************************************************
Private Sub Class_Initialize()
    
    'ALLOWCATE ARRAY
    ReDim pArray(0)
    
End Sub
Private Sub Class_Terminate()
    
    'BEST PRACTICE TO EMPTY ARRAY (NOT NEEDED)
    pArray = Empty
    
End Sub


'********************************************************************
' PUBLIC PROPERTIES
'********************************************************************
Public Property Let value(Arr As Variant)
Attribute value.VB_UserMemId = 0
    
    If IsArray(Arr) Then
        pArray = Arr
    Else
        On Error Resume Next
        If Arr = Empty Then
            pArray = Empty
        End If
    End If
    
End Property
Public Property Get value() As Variant

    value = pArray
    
End Property



'*****************************************************************************
' PUBLIC METHODS\FUNCTIONS
'*****************************************************************************

' - ADDS A NEW ELEMENT TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function push(Element As Variant) As Long

    'INCREASE UBOUND IF ARRAY IS NOT EMPTY
    incrementIfNeeded
    
    'ASGIN USING LET OR SET
    asign pArray(UBound(pArray, 1)), Element
    
    'RETURN NEW ARRAY LENGTH
    push = UBound(pArray, 1) + 1
    
End Function

' @Check - Currently just works for single dimension
' - REMOVES LAST ELEMENT FROM ARRAY, RETURNS REMOVED ELEMENT
Public Function pop() As Variant
    
    pop = pArray(UBound(pArray, 1))
    ReDim Preserve pArray(UBound(pArray, 1) - 1)
    
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
Public Function getColumn(Name As String) As Variant
    
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
    
    getColumn = arrayFromCollection(List)
    
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



'********************************************************************
' PRIVATE METHODS\FUNCTIONS         (HELPER FUNCTIONS)
'********************************************************************

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
    Dim i As Integer
    
    ReDim Arr(0 To Col.Count - 1)
    
    For i = 1 To Col.Count
        asign Arr(i - 1), Col(i)
    Next i
    
    arrayFromCollection = Arr
    
End Function

' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function asign(Arr As Variant, Val As Variant)

    If IsObject(Val) Then
        Set Arr = Val
    Else
        Let Arr = Val
    End If
    
End Function

Private Function incrementIfNeeded()
    
    'INCREASE UBOUND IF ARRAY IS NOT EMPTY
    If pArray(0) <> Empty Then '@Check - This is limited to single dimension
        ReDim Preserve pArray(UBound(pArray, 1) + 1)
    End If

End Function

Private Function isSingleDimension() As Boolean
    If dimensionLength(pArray) = 1 Then isSingleDimension = True
End Function

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



