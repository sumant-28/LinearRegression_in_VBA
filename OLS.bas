Attribute VB_Name = "OLS"
Sub OLS()

Cells(1, 1) = 1
Cells(2, 1) = 1
Cells(3, 1) = 1
Cells(4, 1) = 1
Cells(5, 1) = 1

Dim X1 As Variant
X1 = Range("A1:A5")

Cells(1, 2) = 1
Cells(2, 2) = 2
Cells(3, 2) = 3
Cells(4, 2) = 4
Cells(5, 2) = 5

Dim X2 As Variant
X2 = Range("B1:B5")

Cells(1, 3) = 1
Cells(2, 3) = 4
Cells(3, 3) = 9
Cells(4, 3) = 16
Cells(5, 3) = 25

Dim X3 As Variant
X3 = Range("C1:C5")

Cells(1, 4) = 1
Cells(2, 4) = 5
Cells(3, 4) = 9
Cells(4, 4) = 23
Cells(5, 4) = 36

Dim Y As Variant
Y = Range("D1:D5")

Dim X As Variant
X = Range("A1:C5")

ProjectionMatrix = WorksheetFunction.MInverse(WorksheetFunction.MMult(WorksheetFunction.transpose(X), X))
XTY = WorksheetFunction.MMult(WorksheetFunction.transpose(X), Y)

Beta = WorksheetFunction.MMult(ProjectionMatrix, XTY)

' The coefficient on the constant is 2.4
' The coefficient on X2 is - 3.2
' The coefficient on X3 is -2

End Sub
