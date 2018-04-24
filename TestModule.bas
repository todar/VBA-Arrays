Attribute VB_Name = "TestModule"
Option Explicit

'==============================================================================
' NOTES: CURRENTLY THIS CLASS IS IN THE EARLY DESIGN STAGES. TESTING IS STILL
' GOING ON, AS WELL AS ADDING MORE FUNCTIONS.
'==============================================================================
Private Sub testingArrayObject()
  
    Dim A As New arrayObject
    
    'ADD VALUES TO END OF CLASS ARRAY
    A.push "apple", "carrot"
    A.push "bannana"
    
    
    'REMOVE VALUES FROM END OF CLASS ARRAY (RETURNS ITEM REMOVED(BANNANA))
    Debug.Print A.pop
    
    
    'ADD VALUES TO THE START OF THE CLASS ARRAY
    A.unShift "bannana"
    A.unShift "mango"
    
    
    'REMOVE VALUES FROM START OF THE ARRAY (RETURNS ITEM REMOVED(MANGO))
    Debug.Print A.Shift
    
    
    'CHANGE VALUES BASED ON INDEX (WILL PUSH IF INDEX > UBOUND)(OPTION BASE 0)
    A(1) = "zeebra"
    
    
    'DISPLAY LENGHT OF CLASS ARRAY (ALWAYS ONE MORE THAN THE UBOUND)
    Debug.Print A.Length
    A(A.Length) = "apple" 'CAN FORCE A PUSH THIS WAY...
    
    
    'CONCATE RETURNS THE CURRENT ARRAY JOINED WITH ANOTHER ARRAY.
    A = A.concat(Array("Bacon", "Tuna", "Apple"))
    
    'SPLICE CAN OPTIONALLY INSERT ELEMENTS AT ANY INDEX, CAN ALSO OPTIONALLY REMOVE ELEMENTS FROM THAT SPOT
    A.splice 1, 1, "Lemon", "Kiwi"  'ADDS AT THE 1ST INDEX (BY DEFAULT BASE 0) AND REMOVES 1 ELEMENT
    A.splice 3, 2    'REMOVES TWO ITEMS AT THE 3RD INDEX (BY DEFAULT BASE 0)
    
    'YOU CAN REMOVE DUPLICATES (CURRENTLY ONLY WORKS ON 1D ARRAY)
    A.removeDuplicates
    
    'CHECK IF AN ELEMENT EXISTS
    Debug.Print A.exists("apple")
    
    'toString WILL RETURN THE ARRAY JOINED WITH COMMAS (Optionaly you can set the delimeter)
    Debug.Print A.toString
    
    'YOU CAN ALSO SORT
    A.sort
    Debug.Print A.toString
    
    'REVERSING THE ARRAY WILL GIVE YOU DECENDING ORDER
    A.reverse
    Debug.Print A.toString
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NOTE: THESE FUNCTIONS BELOW ARE NOT AS EFFICIENT AS LOOPING, JUST MORE CONVENIENT
    '
    ' THESE FUNCTIONS LOOP EACH ELEMENT AND USES EXCELS EVALUATE FUNCTION.
    ' NOTE: x IS WHERE THE ELEMENT WILL GO. YOU CAN PASS IN A DIFFERENT KEY AS WELL.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    A = A.map("upper(x)")
    Debug.Print A.toString(", ")
    
    'FOR EACH IS THE SAME AS MAP, HOWEVER IT MANIPULATES THE ORIGINAL ARRAY. DOESN'T RETURN ANYTHING
    A.forEach ("x & "" tastes good!""")
    Debug.Print A.toString(", ")
    
    'EXAMPLES WITH NUMBERS
    A = Empty
    A.push 1, 2, 3, 4
    
    A.forEach ("SUM(2 * x)")
    Debug.Print A.toString(", ")
    
    'REDUCE USES X AS EACH ELEMENT, AND Y AS THE ACCUMILATION. YOU CAN USE IT TO DO THING SUCH AS ADD EACH ELEMENT TOGETHER.
    Debug.Print A.reduce("x + y")
    
    'YOU CAN FILTER ON YOUR ARRAY
    A.filter "x > 5"
    Debug.Print A.toString(", ")
    
    'NOTE: NEED TO ADD EXAMPLES OF THE TOOLS USED FOR 2D ARRAYS!!
    
End Sub




Private Sub createSampleData()
    
    Dim i As Integer
    Dim data As Variant: data = Array("monkey", "Banana", "apple", "carrot", "cage", "elephant", "registration", "agile", "arena", "adviser", "kneel", "steward", "bake", "profession", "costume", "feedback", "begin", "carry", "exercise", "retailer", "gregarious", "rib", "seminar", "Koran")
    Const lb As Integer = 1 'Dont change lower bound

    'SAMPLE HEADINGS
    Range("A1").value = "Sample Text"
    Range("B1").value = "Number"
    Range("C1").value = "Dates"
    Range("D1").value = "Currency"

    For i = lb To 25
        
        'ADD TEXT FROM DATA ARRAY
        Range("A1").Offset(i).value = data(Int((UBound(data, 1) - 0 + 1) * Rnd + 0))
        
        'ADD RANDOM NUMBERS
        Range("b1").Offset(i).value = Int((50 - 0 + 1) * Rnd + 0)
        
        'RANDOM DATES
        Range("c1").Offset(i).value = DateSerial(2017, 1, 1) + Int(Rnd * 730)
        
        'RANDOM CURRENCY
        Range("D1").Offset(i).value = CCur(Int((50 - 0 + 1) * Rnd + 0))
        
    Next i
    
    Columns.AutoFit
    
End Sub
