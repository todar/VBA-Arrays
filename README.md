# VBA-Arrays
A whole bunch of Array functions to make it easier and faster coding. Many functions are to try and mimic JavaScript. Example: Push, Pop, Shift, Unshift, Splice, Sort, Reverse, length, toString.

> See my [Style Guide](https://github.com/todar/VBA-Style-Guide) in how to write clean and maintainable VBA code.
> Also learn more about me on my [portfolio](https://www.roberttodar.com/).

# Public Funtions:
- ArrayAverage
- ArrayContainsEmpties
- ArrayDimensionLength
- ArrayExtractColumn
- ArrayExtractRow
- ArrayFilter
- ArrayFilterTwo
- ArrayFromRecordset
- ArrayGetColumnIndex
- ArrayGetIndexes
- ArrayIncludes
- ArrayIndexOf
- ArrayLength
- ArrayPluck
- ArrayPop
- ArrayPush
- ArrayPushTwoDim
- ArrayQuery
- ArrayRemoveDuplicates
- ArrayReverse
- ArrayShift
- ArraySort
- ArraySplice
- ArraySpread
- ArraySum
- ArrayToCSVFile
- ArrayToString
- ArrayTranspose
- ArrayUnShift
- Assign
- ConvertToArray
- IsArrayEmpty

# Usage

Import ArrayFunctions.bas file. That's it! :)

Below are some of the examples you can do with single dim arrays. Note, there are several functions for two dim arrays as well.

```vb
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
```
