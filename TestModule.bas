Attribute VB_Name = "TestModule"
Option Explicit

'CREATE SAMPLE TEST DATA FOR TESTING FUNCTIONS
Private Sub CreateSampleData()

    '@Author: Robert Todar <robert@roberttodar.com>
    
    'CHANGE FOR MORE OR LESS SAMPLE DATA
    Const NumberOfRows As Integer = 42
    
    'CREATE HEADINGS
    Range("A1").Value = "Sample Text"
    Range("B1").Value = "Number"
    Range("C1").Value = "Dates"
    Range("D1").Value = "Currency"
    
    'RANDOM DATA, FEEL FREE TO CHANGE FOR NEEDS
    Dim Data As Variant
    Data = Array("monkey", "Banana", "apple", "carrot", "cage", "elephant", "registration", "agile", "arena", "adviser", "kneel", "steward", "bake", "profession", "costume", "feedback", "begin", "carry", "exercise", "retailer", "gregarious", "rib", "seminar", "Koran")
    
    'ADD SAMPLE DATA TO ACTIVESHEET
    Dim Index As Integer
    For Index = 1 To NumberOfRows
        
        'ADD RANDOM TEXT FROM DATA ARRAY
        Range("A1").Offset(Index).Value = Data(Int((UBound(Data, 1) - 0 + 1) * Rnd + 0))
        
        'ADD RANDOM NUMBERS
        Range("b1").Offset(Index).Value = Int((50 - 0 + 1) * Rnd + 0)
        
        'RANDOM DATES
        Range("c1").Offset(Index).Value = DateSerial(2017, 1, 1) + Int(Rnd * 730)
        
        'RANDOM CURRENCY
        Range("D1").Offset(Index).Value = CCur(Int((50 - 0 + 1) * Rnd + 0))
        
    Next Index
    
    Columns.AutoFit
    
End Sub
