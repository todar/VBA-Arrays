# VBA Arrays

A whole bunch of Array functions to make it easier and faster coding. Many functions are to try and mimic JavaScript. Example: Push, Pop, Shift, Unshift, Splice, Sort, Reverse, length, toString.

---

## Other Helpful Resources

- [www.roberttodar.com](https://www.roberttodar.com/) About me and my background and some of my other projects.
- [Style Guide](https://github.com/todar/VBA-Style-Guide) A guide for writing clean VBA code. Notes on how to take notes =)
- [Boilerplate](https://github.com/todar/VBA-Boilerplate) Boilerplate that contains a bunch of helper libraries such as JSON tools, Code Analytics, LocalStorage, Unit Testing, version control and local network distribution, userform events, and more!
- [Strings](https://github.com/todar/VBA-Strings) String function library. `ToString`, `Inject`, `StringSimilarity`, and more.
- [Analytics](https://github.com/todar/VBA-Analytics) Way of tracking code analytics and metrics. Useful when multiple users are running code within a shared network.
- [Userform EventListener](https://github.com/todar/VBA-Userform-EventListener) Listen to events such as `mouseover`, `mouseout`, `focus`, `blur`, and more.

---

## List of Available Functions

| Function Name           | Description                                                                                                        |
| :----------------------- | :------------------------------------------------------------------------------------------------------------------ |
| `ArrayAverage`          | Returns the average of all the numbers inside an array.                                                            |
| `ArrayContainsEmpties`  | Returns `True` if the array contains any empties.                                                                  |
| `ArrayDimensionLength`  | Returns the dimensionlenght of the array.                                                                          |
| `ArrayExtractColumn`    | Extracts a column from a 2 dim array and returns it as a 1 dim array                                               |
| `ArrayExtractRow`       | Extracts a row from a 2 dim array and returns it as a 1 dim array                                                  |
| `ArrayFilter`           | Uses regex to filter items in a single dim array                                                                   |
| `ArrayFilterTwo`        | Uses regex to filter items in a two dim array.                                                                     |
| `ArrayFromRecordset`    | Converts a recordset into a 2 dim array including it's headers                                                     |
| `ArrayGetColumnIndex`   | Return the column index based on the header name                                                                   |
| `ArrayGetIndexes`       | Returns a single dim array of the indexes of column headers                                                        |
| `ArrayIncludes`         | Checks to see if a value is in single dim array                                                                    |
| `ArrayIndexOf`          | Returns the index of an item in a single dim array                                                                 |
| `ArrayLength`           | Returns the number of items in an array                                                                            |
| `ArrayPluck`            | Extracts a list of a given property. Must be array of dictionries                                                  |
| `ArrayPop`              | Removes the last element in array, returns the popped element                                                      |
| `ArrayPush`             | Adds a new element(s) to an array (at the end), returns the new array length                                       |
| `ArrayPushTwoDim`       | Adds a new element(s) to an array (at the end). Must be full row of data                                           |
| `ArrayQuery`            | Saves array in CSV file and allows the ability to run ADODB queries on it.                                         |
| `ArrayRemoveDuplicates` | Removed duplicates from single dim array                                                                           |
| `ArrayReverse`          | Reverse array (can be used after sort to get the descending order)                                                 |
| `ArrayShift`            | Removes element from array - returns removed element                                                               |
| `ArraySort`             | Sort an array                                                                                                      |
| `ArraySplice`           | Changes the contents of an array by removing or replacing existing elements and/or adding new elements.            |
| `ArraySpread`           | Spreads out an array into a single array. example: jagged arrays, dictionaries, collections.                       |
| `ArraySum`              | Returns the Sum of a single dim array containing numbers                                                           |
| `ArrayToCSVFile`        | Saves a two dim array to a CSV file                                                                                |
| `ArrayToString`         | Returns a string from a 1 or 2 dim array, separated by optional delimiter and vbnewline for each row               |
| `ArrayTranspose`        | Application.Transpose has a limit on the size of the array and is limited to the 1st dim. This fixes those issues. |
| `ArrayUnShift`          | Adds a new element to the begining of the array                                                                    |
| `Assign`                | Quick tool to either set or let depending on if element is an object                                               |
| `ConvertToArray`        | Convert other list type objects to an array                                                                        |
| `IsArrayEmpty`          | This function tests whether the array is empty (unallocated). Returns TRUE or FALSE.                               |

---

## How to use

1.  Import ArrayFunctions.bas file.
2.  Set a reference to `Microsoft Scripting Runtime` as this uses dictionaries for removing duplicates.

---

## Examples

Below are some of the examples you can do with single dim arrays. Note, there are several functions for two dim arrays as well.

```vb
'EXAMPLES OF VARIOUS FUNCTIONS
Private Sub arrayFunctionExamples()
    ' For simplicity using `a` as the variable. Otherwise, don't do that in your real code! =)
    Dim a As Variant

    ' Single dim functions that manipulate the array.
    ArrayPush a, "Banana", "Apple", "Carrot" '--> Banana,Apple,Carrot
    ArrayPop a                               '--> Banana,Apple --> returns Carrot
    ArrayUnShift a, "Mango", "Orange"        '--> Mango,Orange,Banana,Apple
    ArrayShift a                             '--> Orange,Banana,Apple
    ArraySplice a, 2, 0, "Coffee"            '--> Orange,Banana,Coffee,Apple
    ArraySplice a, 0, 1, "Mango", "Coffee"   '--> Mango,Coffee,Banana,Coffee,Apple
    ArrayRemoveDuplicates a                  '--> Mango,Coffee,Banana,Apple
    ArraySort a                              '--> Apple,Banana,Coffee,Mango
    ArrayReverse a                           '--> Mango,Coffee,Banana,Apple

    ' Array properties functions.
    ' These get details of the array: index of items, lenght, ect.
    ArrayLength a                            '--> 4
    ArrayIndexOf a, "Coffee"                 '--> 1
    ArrayIncludes a, "Banana"                '--> True
    arrayContains a, Array("Test", "Banana") '--> True
    ArrayContainsEmpties a                   '--> False
    ArrayDimensionLength a                   '--> 1 (single dim array)
    IsArrayEmpty a                           '--> False

    ' Here is an example of a jagged array.
    a = Array(1, 2, 3, Array(4, 5, 6, Array(7, 8, 9)))

    ' Can flatten jagged array with the spread formula. Note this is a deep spread.
    ' This formula also spreads dictionaires and collections as well!
    a = ArraySpread(a)                       '--> 1,2,3,4,5,6,7,8,9

    ' Math function examples
    ArraySum a                               '--> 45
    ArrayAverage a                           '--> 5

    ' Filter use's regex pattern
    a = Array("Banana", "Coffee", "Apple", "Carrot", "Canolope")
    a = ArrayFilter(a, "^Ca|^Ap")

    ' Array to string works with both single and double dim arrays!
    Debug.Print ArrayToString(a)
End Sub
```
