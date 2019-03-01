Attribute VB_Name = "CalculateMean"
''
' ArithmeticMeanCalculator
'   (c) 2019 mikhail911
'       license: MIT
''

Sub Calculate_Click()
    Dim i As Long
    Dim numeratorWeighted As Single: numeratorWeighted = 0
    Dim denumeratorWeighted As Single: denumeratorWeighted = 0
    Dim avgSum As Single: avgSum = 0
    Dim avgItemsCount As Integer: avgItemsCount = 0
    Dim calcAvg As Single: calcAvg = 0
    Dim calcWeightedAvg As Single: calcWeightedAvg = 0
    Dim valueItems As Integer: valueItems = 0
    Dim weightItems As Integer: weightItems = 0
    
    For i = 2 To 1000
        'Calculate weighted arithmetic mean
        If Not IsEmpty(Cells(i, 2).Value) And Not IsEmpty(Cells(i, 3).Value) Then
            weightItems = weightItems + 1
            numeratorWeighted = numeratorWeighted + (Cells(i, 2).Value * Cells(i, 3).Value)
            denumeratorWeighted = denumeratorWeighted + Cells(i, 3).Value
            r = WorksheetFunction.RandBetween(0, 255)
            g = WorksheetFunction.RandBetween(0, 255)
            b = WorksheetFunction.RandBetween(0, 255)
            Cells(i, 2).Borders.Color = RGB(r, g, b)
            Cells(i, 3).Borders.Color = RGB(r, g, b)
        Else
            Cells(i, 2).Borders.ColorIndex = xlColorIndexNone
            Cells(i, 3).Borders.ColorIndex = xlColorIndexNone
        End If
    
        'Calculate arithmetic mean
        If Not IsEmpty(Cells(i, 2).Value) Then
            valueItems = valueItems + 1
            avgSum = avgSum + Cells(i, 2).Value
            avgItemsCount = avgItemsCount + 1
            Cells(i, 2).Interior.Color = RGB(255, 255, 204)
        Else
            Cells(i, 2).Interior.ColorIndex = xlColorIndexNone
        End If
    Next i
    
    'Display results
    If Not numeratorWeighted = 0 And Not denumeratorWeighted = 0 Then
        calcWeightedAvg = numeratorWeighted / denumeratorWeighted
        Range("H3") = calcWeightedAvg
    End If
    If Not avgSum = 0 And Not avgItemsCount = 0 Then
        calcAvg = avgSum / avgItemsCount
        Range("H2") = calcAvg
    End If
    If Not valueItems = weightItems Then
        Range("I3") = "Mean calculated only for values which have weight!"
        Range("I3").EntireColumn.AutoFit
        Range("I3").Interior.Color = RGB(255, 51, 51)
    Else
        Range("I3") = ""
        Range("I3").Interior.ColorIndex = xlColorIndexNone
    End If
End Sub
