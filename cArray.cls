VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Option Base 0 'DEFAULT

'THIS IS WHERE THE ARRAY IS STORED. PUBLIC, SO IT CAN BE RETREIVED AND USED NORMALLY.
Public pArray As Variant


'TODO:
' - ADD MORE ERROR HANDLING. ESPECIALLY FOR ENSURING SINGLE DIMENSIONAL ARRAYS ONLY
' - ADD MORE FUNCTIONS THAT ARE COMMON WITH OTHER PROGRAMMING LANGUAGES.

'Properties:
' - value
' - lenght

'Funtions:
' - push
' - pop
' - shift
' - unshift
' - map
' - concat
' - sort
' - reverse
' - toString
' - join

'********************************************************************
' INITIALIZE
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
' PROPERTIES
'********************************************************************
'MAIN PROPERTY (USED TO GET OR LET THE VALUE OF AN INDEX.)
'Attribute Value.VB_UserMemId = 0' SET IT AS THE DEFAULT PROPERTY (MUST IMPORT CLASS TO WORK)
Public Property Let Value(Index As Long, Element As Variant)
    
    'INDEX DOESN'T EXIST, PUSH ITEM TO THE END
    If UBound(pArray, 1) < Index Then
        push Element
        
    'USE SET IF AN OBJECT
    ElseIf IsObject(Element) Then
        Set pArray(Index) = Element
        
    'REGULAR LET
    Else
        pArray(Index) = Element
    End If
    
End Property
Public Property Get Value(Index As Long) As Variant
    If IsObject(pArray(Index)) Then
        Set Value = pArray(Index)
    Else
        Value = pArray(Index)
    End If
End Property

'THE LENGTH PROPERTY IS ALWAYS ONE MORE THAN THE HIGHEST ARRAY INDEX.
Public Property Get Length() As Long
    Length = UBound(pArray) + 1
End Property

'********************************************************************
' MAIN FUNCTIONS
'********************************************************************
'CONVERTS AN ARRAY TO A STRING OF (COMMA SEPARATED) ARRAY VALUES
Public Function toString() As String
    
    On Error Resume Next
    toString = VBA.join(pArray, ",")

End Function

'IT BEHAVES JUST LIKE TOSTRING(), BUT IN ADDITION YOU CAN SPECIFY THE SEPARATOR (VBA CODE)
Public Function join(Delimiter As String) As String
    
    On Error Resume Next
    join = VBA.join(pArray, Delimiter)
    
End Function

'REMOVES LAST ELEMENT FROM ARRAY, RETURNS REMOVED ELEMENT
Public Function pop() As Variant
    
    pop = pArray(UBound(pArray, 1))
    ReDim Preserve pArray(UBound(pArray, 1) - 1)
    
End Function

'ADDS A NEW ELEMENT TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Public Function push(Element As Variant) As Long
    
    'INCREASE UBOUND IF ARRAY IS NOT EMPTY
    If pArray(0) <> Empty Then
        ReDim Preserve pArray(UBound(pArray, 1) + 1)
    End If
    
    'CHECK TO SEE IF IT IS AN OBJECT
    If IsObject(Element) = True Then
       Set pArray(UBound(pArray, 1)) = Element
    Else
        pArray(UBound(pArray, 1)) = Element
    End If
    
    'RETURN NEW ARRAY LENGTH
    push = UBound(pArray, 1) + 1
    
End Function

'REMOVES THE FIRST ARRAY ELEMENT AND 'SHIFTS' ALL OTHER ELEMENTS TO A LOWER INDEX.
'RETURNS THE ELEMENT THAT WAS SHIFTED OUT
Public Function shift() As Variant
    
    Dim I As Long
    Dim Arr As Variant
    
    shift = pArray(LBound(pArray, 1))
    ReDim Arr(LBound(pArray, 1) To UBound(pArray, 1) - 1)
    
    For I = LBound(pArray, 1) + 1 To UBound(pArray, 1)
        
        If IsObject(pArray(I)) Then
            Set Arr(I - 1) = pArray(I)
        Else
            Arr(I - 1) = pArray(I)
        End If
        
    Next I
    
    pArray = Arr
    
End Function

'ADDS A NEW ELEMENT TO AN ARRAY (AT THE BEGINNING), AND 'PUSHES' OLDER ELEMENTS
'RETURNS THE NEW ARRAY LENGTH
Public Function unshift(Element) As Long
    
    Dim I As Long
    Dim Arr As Variant
    
    unshift = UBound(pArray, 1) + 1
    
    ReDim Arr(LBound(pArray, 1) To UBound(pArray, 1) + 1)
    
    Arr(LBound(Arr, 1)) = Element
    
    For I = LBound(pArray, 1) To UBound(pArray, 1)
        If IsObject(pArray(I)) Then
            Set Arr(I + 1) = pArray(I)
        Else
            Arr(I + 1) = pArray(I)
        End If
    Next I

    pArray = Arr
    
End Function

'RETURN A NEW ARRAY, JOINS CURRENT ARRAY WITH ANOTHER ONE.
Public Function concat(Arr) As Variant
    
    Dim I As Integer
    Dim pArr As Variant
    
    pArr = pArray

    For I = LBound(Arr, 1) To UBound(Arr, 1)
        ReDim Preserve pArr(UBound(pArr, 1) + 1)
        If IsObject(Arr(I)) Then
            Set pArr(UBound(pArr, 1)) = Arr(I)
        Else
            pArr(UBound(pArr, 1)) = Arr(I)
        End If
        
        
    Next I
    
    concat = pArr
    
End Function

Public Function forEach(macroName As String, Optional WorkbookName As String)
    
    Dim I As Integer
    Dim Temp As String
    
    If WorkbookName = "" Then
        WorkbookName = ThisWorkbook.Name
    End If
        
        WorkbookName = "'" & WorkbookName & "'"
        
    Temp = WorkbookName & "!" & macroName
    
    For I = LBound(pArray, 1) To UBound(pArray, 1)
        Application.Run Temp, pArray(I)
    Next I
    

End Function


'LOOPS THROUGH EACH ELEMENT, RUNNINGS EXCELS EVALUATE FUNCTION FROM THE FORMULA STRING
Public Function map(Formula As String) As Variant
    
    Dim I As Integer
    Dim Arr As Variant
    Dim S As String
    
    Arr = pArray
    
    For I = LBound(Arr, 1) To UBound(Arr, 1)
        
        If IsObject(Arr(I)) Then GoTo Nxt
        
        If TypeName(Arr(I)) = "String" Then Arr(I) = """" & Arr(I) & """"
        If InStr(Formula, "{}") > 0 Then
            S = Replace(Formula, "{}", Arr(I))
        Else
            S = S & Arr(I)
        End If
        
        Arr(I) = Evaluate(S)
Nxt:
    Next I
    
    map = Arr
    
End Function

'********************************************************************
' MAIN METHODS
'********************************************************************
'SORT THE ARRAY. NO RETURN, EFFECTS pArray
Public Sub sort()
    
    Dim Arr As Object
    Dim Element
    
    'Creating a array list
    Set Arr = CreateObject("System.Collections.ArrayList")
    
    
    ' adding the elements in the array to array_list
    For Each Element In pArray
        Arr.Add Element
    Next
    
    On Error GoTo Catch
    'sorting happens
    Arr.sort
    
    'Converting ArrayList to an array
    'so now a sorted array of elements is stored in the array sorted_array.
    
    pArray = Arr.ToArray
    
Catch:

End Sub

'REVERSE ARRAY (CAN BE USED AFTER SORT TO GET THE DECENDING ORDER)
Public Sub reverse()
    
    Dim Temp As Variant
    Dim I As Integer
    
    'REVERSE LOOP (HALF OF IT, WILL WORK FROM BOTH SIDES ON EACH ITERATION)
    For I = LBound(pArray, 1) To UBound(pArray, 1) / 2
        
        'STORE LAST VALUE MINUS THE ITERATION
        If IsObject(pArray(UBound(pArray) - I)) Then
            Set Temp = pArray(UBound(pArray) - I)
        Else
            Temp = pArray(UBound(pArray) - I)
        End If
        
        'SET LAST VALUE TO FIRST VALUE OF THE ARRAY
        If IsObject(pArray(I)) Then
            Set pArray(UBound(pArray) - I) = pArray(I)
        Else
            pArray(UBound(pArray) - I) = pArray(I)
        End If
        
        'SET FIRST VALUE TO THE STORED LAST VALUE
        If IsObject(Temp) Then
            Set pArray(I) = Temp
        Else
             pArray(I) = Temp
        End If
        
    Next I
    

    
End Sub

'EMPTY THE ARRAY
Public Sub emptyAndRedim()
    pArray = Empty
    ReDim pArray(0)
End Sub

'TODO:
' - NEED TO CREATE FUNCTIONS BELOW

'**********************************************************************
' NEED TO CREATE FUNCTIONS BELOW. PRIVATE UNTIL THEY ARE FINISHED
'**********************************************************************

'WILL BE USED TO ADD NEW ITEMS TO AN ARRAY, AS WELL AS REMOVE ITEMS FROM THAT LOCATION
Private Sub splice(Where As Long, howManyRemoved As Integer, Element As Variant)
    
    
    Dim I As Long
    Dim Arr As Variant
    
    ReDim Arr(LBound(pArray, 1) To UBound(pArray, 1) + UBound(Element, 1) + 1 - howManyRemoved)
    
    For I = LBound(Arr, 1) To UBound(Arr, 1)
        
        If I = Where Then
            
        Else
            Arr(I) = pArray(I)
        End If
        
    Next I
    
    pArray = Arr
    
End Sub

Private Function slice(sliceStart As Integer, Optional sliceEnd As Integer)

End Function

Private Function filter() As Variant
    
End Function

'RETURNS TRUE IF ANY ELEMENT MATCHES
Private Function some() As Boolean
    
End Function

'RETURNS THE LAST INSTANCE OF AN ELEMENT
Private Function lastIndexOf(Element As Variant) As Long
    
End Function

'RETURNS THE FIRST INSTANCE OF AN ELEMENT
Private Function IndexOf(Element As Variant) As Long
    
End Function








