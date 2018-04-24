# VBA-Arrays
Custom Arrays in class module, that have similar functions as JavaScript. Example: Push, Pop, Shift, Unshift, Sort, map, length, concat,  toString.

*Currently working on a version that has support for multi dim arrays!! Adding funtions to query & sort them. If interested, check it out!!

# Properties:
- value
- lenght

# Public Funtions:
- push
- pop
- shift
- unshift
- filter
- map
- forEach
- reduce
- exists
- concat
- sort
- reverse
- toString
- toRange
- columnNumber
- returnColumn
- returnRow

# Private Helper Funtions:
- arrayFromCollection
- asign
- collectionToMultiDimArray
- collectionFromarray
- dimensionLength
- isSingleDimension
- array2dUnshift
- sqlArray

# Usage

Must import file for Property Value to be set as the class default.
No extra refrences to other libriaries needed at this time.

Below is a test module, that shows some of the functions, and how to work with them. (Orginal, need to add new arrayObject examples)

```vb
'==============================================================================
' NOTES: CURRENTLY THIS CLASS IS IN THE EARLY DESIGN STAGES. TESTING IS STILL
' GOING ON, AS WELL AS ADDING MORE FUNCTIONS. 
'==============================================================================
Private Sub testClassArrays()
  
    Dim A As New arrayObject
    
    'ADD VALUES TO END OF CLASS ARRAY
    A.push "apple", "carrot"
    A.push "bannana"
    
    
    'REMOVE VALUES FROM END OF CLASS ARRAY (RETURNS ITEM REMOVED(BANNANA))
    Debug.Print A.pop
    
    
    'ADD VALUES TO THE START OF THE CLASS ARRAY
    A.unshift "bannana"
    A.unshift "mango"
    
    
    'REMOVE VALUES FROM START OF THE ARRAY (RETURNS ITEM REMOVED(MANGO))
    Debug.Print A.shift
    
    
    'CHANGE VALUES BASED ON INDEX (WILL PUSH IF INDEX > UBOUND)(OPTION BASE 0)
    A(1) = "zeebra"
    
    
    'DISPLAY LENGHT OF CLASS ARRAY (ALWAYS ONE MORE THAN THE UBOUND)
    Debug.Print A.length
    A(A.length) = "apple" 'CAN FORCE A PUSH THIS WAY...
    
    
    'CONCATE RETURNS THE CURRENT ARRAY JOINED WITH ANOTHER ARRAY.
    A = A.concat(Array("Bacon", "Tuna", "Apple"))
    
    'SPLICE CAN OPTIONALLY INSERT ELEMENTS AT ANY INDEX, CAN ALSO OPTIONALLY REMOVE ELEMENTS FROM THAT SPOT
    A.splice 1, 1, "Lemon", "Kiwi"  'ADDS AT THE 1ST INDEX (BY DEFAULT BASE 0) AND REMOVES 1 ELEMENT
    A.splice 3, 2    'REMOVES TWO ITEMS AT THE 3RD INDEX (BY DEFAULT BASE 0)
    
    'YOU CAN REMOVE DUPLICATES (CURRENTLY ONLY WORKS ON 1D ARRAY)
    a.removeDuplicates
    
    'CHECK IF AN ELEMENT EXISTS
    Debug.Print A.exists("apple")
    
    'toString WILL RETURN THE ARRAY JOINED WITH COMMAS (Optionaly you can set the delimeter)
    Debug.Print A.toString
    
    'YOU CAN ALSO SORT
    A.sort
    Debug.Print A.toString
    
    'REVERSING THE ARRAY WILL GIVE YOU DECENDING ORDER
    A.reverse
    Debug.Print A.ToString
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NOTE: THESE FUNCTIONS BELOW ARE NOT AS EFFICIENT AS LOOPING, JUST MORE CONVENIENT
    '
    ' MAPPING WILL USE EXCELS EVALUATE FUNCTION, TO TAKE A STRING AND DO
    ' BASIC FORMULAS TO EACH ITEM IN THE ARRAY, AND RETURN A NEW ARRAY.
    ' EXAMPLE BELOW IS UPPER CASE EACH ITEM. 
    ' NOTE: ${} IS WHERE THE ELEMENT WILL GO. YOU CAN PASS IN A DIFFERENT KEY AS WELL.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    A = A.Map("upper(${})")
    Debug.Print A.toString(", ")
    
    'FOR EACH IS THE SAME AS MAP, HOWEVER IT MANIPULATES THE ORIGINAL ARRAY. DOESN'T RETURN ANYTHING
    A.forEach("${} & "" are good!""")
    Debug.Print A.toString(", ")
    
    'EXAMPLES WITH NUMBERS
    A = empty
    A.push 1, 2, 3, 4
    
    A.forEach("SUM(2 * ${})")
    Debug.Print A.toString(", ")
    
    'REDUCE USES X AS EACH ELEMENT, AND Y AS THE ACCUMILATION. YOU CAN USE IT TO DO THING SUCH AS ADD EACH ELEMENT TOGETHER.
    A.reduce "X + Y"
    
    'YOU CAN FILTER ON YOUR ARRAY
    A.filter "X > 10"
    
    'NOTE: NEED TO ADD EXAMPLES OF THE TOOLS USED FOR 2D ARRAYS!!
    
End Sub





```
