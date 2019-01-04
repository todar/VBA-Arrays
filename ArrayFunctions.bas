Attribute VB_Name = "arrayFunctions"
Option Explicit
Option Compare Text
Option Private Module
Option Base 0


'@AUTHOR: ROBERT TODAR

'DEPENDENCIES
' -

'PUBLIC FUNCTIONS
' - arrayGetColumnNumber
' - ArrayQuery
' - ArrayFromRecordset
' - ArrayToString
' -
' -
' -

'PRIVATE METHODS/FUNCTIONS (IN DEVELOPMENT)
' - ArrayToTextFile -NEED TO ADJUST BEFORE MAKING PUBLIC
' - arrayPush
' - isSingleDimension
' - dimensionLength
' - asign
' -
' -

'NOTES:
' - I'VE CREATE AN ARRAY CLASS MODULE THAT DOES MANY OF THESE FUNCTIONS, DECIDED TO ALSO
' - CREATE FUNCTIONS AWAY FROM CLASS MODULE OBJECT TO MAKE THEM WORK WITH ANY ARRAY.

'TODO:
' - ADD MORE FUNCTIONS FROM ARRAYOBJECT CLASS MODULE
' - FINISH PRIVATE METHODS AND MAKE THEM PUBLIC.
' - REMOVE THE NEED TO HAVE A ADODB REFERENCE

'EXAMPLES:
' -

'******************************************************************************************
' TESTING
'******************************************************************************************

'USED TO GET SAMPLE DATA FROM ACTIVESHEET
Private Property Get TestData() As Variant
    TestData = Range("A1").CurrentRegion
End Property

'USED FOR TESTING NEW AND MODIFIED FUNCTIONS
Public Sub TestingArrayFunctions()

    Dim Arr As Variant
    Dim sql As String
    
    sql = "SELECT * FROM []"
    
    Arr = ArrayQuery(TestData, sql)
    Debug.Print ArrayToString(Arr)
    
End Sub


'******************************************************************************************
' PUBLIC FUNCTIONS
'******************************************************************************************

'LOOKS FOR VALUE IN FIRST ROW OF A TWO DIMENSIONAL ARRAY, RETURNS IT'S COL INDEX
Public Function ArrayGetColumnNumber(Arr As Variant, HeadingValue As String) As Integer
    
    Dim columnIndex As Integer
    For columnIndex = LBound(Arr, 2) To UBound(Arr, 2)
        If Arr(LBound(Arr, 1), columnIndex) = HeadingValue Then
            ArrayGetColumnNumber = columnIndex
            Exit Function
        End If
    Next columnIndex
    
    'RETURN NEGATIVE IF NOT FOUND
    ArrayGetColumnNumber = -1
    
End Function


'CREATES TEMP TEXT FILE AND SAVES ARRAY VALUES IN A CSV FORMAT, THEN QUERIES AND RETURNS ARRAY.
'
'@AUTHOR ROBERT TODAR
'@USES ArrayToTextFile
'@USES ArrayFromRecordset
'@RETURNS 2D ARRAY || EMPTY (IF NO RECORDS)
'@PARAM {ARR} MUST BE A TWO DIMENSIONAL ARRAY, SETUP AS IF IT WERE A TABLE.
'@PARAM {SQL} ADO SQL STATEMENT FOR A TEXT FILE. MUST INCLUDE 'FROM []'
'@PARAM {IncludeHeaders} BOOLEAN TO RETURN HEADERS WITH DATA OR NOT
'@EXAMPLE SQL = "SELECT * FROM [] WHERE [FIRSTNAME] = 'ROBERT'"
Public Function ArrayQuery(Arr As Variant, sql As String, Optional IncludeHeaders As Boolean = True) As Variant
    
    'CREATE TEMP FOLDER AND FILE NAMES
    Const fileName As String = "temp.txt"
    Dim FilePath As String
    FilePath = Environ("temp")
    
    'UPDATE SQL WITH TEMP FILE NAME
    sql = Replace(sql, "FROM []", "FROM [" & fileName & "]")
    
    'SEND ARRAY TO TEMP TEXTFILE IN CSV FORMAT
    ArrayToTextFile Arr, FilePath & "\" & fileName, ","
    
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
    Kill FilePath & "\" & fileName
    
End Function

'RETURNS A 2D ARRAY FROM A RECORDSET, OPTIONALLY INCLUDING HEADERS, AND IT TRANSPOSES TO KEEP
'ORIGINAL OPTION BASE. (TRANSPOSE WILL SET IT TO BASE 1 AUTOMATICALLY.)
'
'@AUTHOR ROBERT TODAR
Public Function ArrayFromRecordset(Rs As Recordset, Optional IncludeHeaders As Boolean = True) As Variant
    
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

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    Dim i As Integer
    Dim test As Long

    On Error GoTo catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
    
catch:
    ArrayDimensionLength = i - 1

End Function

'SENDS AN ARRAY TO A TEXTFILE
Public Sub ArrayToTextFile(Arr As Variant, FilePath As String, Optional delimeter As String = ",")
    
    Dim Fso As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
    
    Dim ts As Object
    Set ts = Fso.OpenTextFile(FilePath, 2, True) '2=WRITEABLE
    ts.Write ArrayToCSV(Arr)
    
    Set ts = Nothing
    Set Fso = Nothing

End Sub


Public Function ArrayToCSV(SourceArray As Variant) As String
    
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
    
    ArrayToCSV = Temp
    
End Function


'******************************************************************************************
' PRIVATE FUNCTIONS - BEING DEVELOPED STILL
'******************************************************************************************

Private Sub TestArrayPush()
    
    Dim Arr As Variant
    
    ArrayPush Arr, 1, 2, 3, 4, 5
    ArrayPush Arr, 6, 7, 8
    
    Debug.Print ArrayToString(Arr)

End Sub

' - ADDS A NEW ELEMENT(S) TO AN ARRAY (AT THE END), RETURNS THE NEW ARRAY LENGTH
Private Function ArrayPush(SourceArray As Variant, ParamArray Element() As Variant) As Long

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

'==================================================
' CONVERT TO ARRAY
'==================================================
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


' - QUICK TOOL TO EITHER SET OR LET DEPENDING ON IF ELEMENT IS AN OBJECT
Private Function Asign(variable As Variant, Value As Variant)

    If IsObject(Value) Then
        Set variable = Value
    Else
        Let variable = Value
    End If
    
End Function






