Attribute VB_Name = "TestModule"
Option Explicit

Private Sub CreateSampleData()
    
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
