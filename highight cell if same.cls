VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
#'Here we want to highlight the first cell
'if certains cell of the same row is the same
Sub highlight_same()
' Firstly you declare your variables
Dim r
Dim f
r = Range("a1").End(xlDown).Row
'you willuse the if statement to check if condituon is meant
'And the for statement for repetition among columns
For f = r To 1 Step -1
  If Cells(f, 2) = Cells(f, 5) Then
     Cells(f, 1).Interior.Color = vbYellow
  End If
 Next

End Sub#
