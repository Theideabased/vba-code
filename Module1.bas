Attribute VB_Name = "Module1"
Sub seet()
Dim name
Dim test
Dim statistics
Sheets.Add(, Sheets(Sheets.Count)).name = "Test1"
Sheets.Add(, Sheets(Sheets.Count)).name = "statistics1"

ActiveWorkbook.Sheets("Test1").Range("c2:c6").Interior.ColorIndex = 4
ActiveWorkbook.Sheets("Test1").Range("d2:d6").Interior.ColorIndex = 6
ActiveWorkbook.Sheets("Test1").Range("e2:e6").Interior.ColorIndex = 3

For Each Value In ActiveWorkbook.Sheets("Test1").Range("a2:a6")
    ActiveWorkbook.Sheets("Test1").Range("a2") = 1
    ActiveWorkbook.Sheets("Test1").Range("a3") = 2
    ActiveWorkbook.Sheets("Test1").Range("a4") = 3
    ActiveWorkbook.Sheets("Test1").Range("a5") = 4
    ActiveWorkbook.Sheets("Test1").Range("a6") = 5
Next Value
For Each Value In ActiveWorkbook.Sheets("Test1").Range("b2:b6")
    ActiveWorkbook.Sheets("Test1").Range("b2") = ActiveWorkbook.Sheets("Test1").Range("a2") + 100
    ActiveWorkbook.Sheets("Test1").Range("b3") = ActiveWorkbook.Sheets("Test1").Range("a3") + 100
    ActiveWorkbook.Sheets("Test1").Range("b4") = ActiveWorkbook.Sheets("Test1").Range("a4") + 100
    ActiveWorkbook.Sheets("Test1").Range("b5") = ActiveWorkbook.Sheets("Test1").Range("a5") + 100
    ActiveWorkbook.Sheets("Test1").Range("b6") = ActiveWorkbook.Sheets("Test1").Range("a6") + 100
Next Value

If ActiveWorkbook.Sheets("name").Range("b2").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("c2").Value = 1
If ActiveWorkbook.Sheets("name").Range("b3").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("c3").Value = 1
If ActiveWorkbook.Sheets("name").Range("b4").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("c4").Value = 1
If ActiveWorkbook.Sheets("name").Range("b5").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("c5").Value = 1
If ActiveWorkbook.Sheets("name").Range("b6").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("c6").Value = 1
If ActiveWorkbook.Sheets("name").Range("c2").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("d2").Value = 1
If ActiveWorkbook.Sheets("name").Range("c3").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("d3").Value = 1
If ActiveWorkbook.Sheets("name").Range("c4").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("d4").Value = 1
If ActiveWorkbook.Sheets("name").Range("c5").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("d5").Value = 1
If ActiveWorkbook.Sheets("name").Range("c6").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("d6").Value = 1
If ActiveWorkbook.Sheets("name").Range("d2").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e2").Value = 1
If ActiveWorkbook.Sheets("name").Range("d3").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e3").Value = 1
If ActiveWorkbook.Sheets("name").Range("d4").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e4").Value = 1
If ActiveWorkbook.Sheets("name").Range("d5").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e5").Value = 1
If ActiveWorkbook.Sheets("name").Range("d6").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e6").Value = 1
If ActiveWorkbook.Sheets("name").Range("e2").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e2").Value = ActiveWorkbook.Sheets("Test1").Range("e2").Value + 1
If ActiveWorkbook.Sheets("name").Range("e3").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e3").Value = ActiveWorkbook.Sheets("Test1").Range("e3").Value + 1
If ActiveWorkbook.Sheets("name").Range("e4").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e4").Value = ActiveWorkbook.Sheets("Test1").Range("e4").Value + 1
If ActiveWorkbook.Sheets("name").Range("e5").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e5").Value = ActiveWorkbook.Sheets("Test1").Range("e5").Value + 1
If ActiveWorkbook.Sheets("name").Range("e6").Value = True Then _
    ActiveWorkbook.Sheets("Test1").Range("e6").Value = ActiveWorkbook.Sheets("Test1").Range("e6").Value + 1
    
ActiveWorkbook.Sheets("statistics1").Range("a1").Value = "total"
ActiveWorkbook.Sheets("statistics1").Range("a2").Value = "percent"

ActiveWorkbook.Sheets("statistics1").Range("b1").Formula = "=sum(Test1!c2:c6)"
ActiveWorkbook.Sheets("statistics1").Range("c1").Formula = "=sum(Test1!d2:d6)"
ActiveWorkbook.Sheets("statistics1").Range("d1").Formula = "=sum(Test1!e2:e6)"

ActiveWorkbook.Sheets("statistics1").Range("b2").Formula = "=b1/(b1+c1+d1) * 100"
ActiveWorkbook.Sheets("statistics1").Range("c2").Formula = "=(c1/(b1+c1+d1))* 100"
ActiveWorkbook.Sheets("statistics1").Range("d2").Formula = "=(d1/(b1+c1+d1))*100"

    
End Sub
