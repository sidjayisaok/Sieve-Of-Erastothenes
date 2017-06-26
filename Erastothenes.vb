'handy prime number tester, param is long and function is boolean
Function testPrimes(param As Long) As Boolean

'variable used
Dim x As Integer

'algorithm to sort primes from non-primes
If param < 2 Or (param <> 2 And param Mod 2 = 0) Or param <> Int(param) Then
    Exit Function
End If
'loop for the rest of the numbers    
For x = 3 To Math.Sqr(param) Step 2
    If param Mod x = 0 Then
        Exit Function
    End If
Next x

testPrimes = True

End Function

'procedural code to solve problem
Sub Erastothenes()

'variables
Dim x, a As Integer
Dim y, z, c As Long

'improve excel efficiency
With Application
    .Calculation = xlCalculationManual
    .EnableEvents = False
    .ScreenUpdating = False

'variables used to manipulate sheet
x = 4
a = Application.InputBox("Enter a number, but keep it under 1000 for brevity")

'clean the sheet
Cells.Delete

'renders display to squares
With Cells
    .RowHeight = 5 * x + 5
    .ColumnWidth = x + 5
    .WrapText = True
    .Font.Color = RGB(0, 0, 0)
End With

'Sieve of Eratosthenes used to generate graph
For y = 1 To a
For z = 1 To a
    'matrix manipulation
    Cells(y, z) = (y * a) + (z - a)
    'pass function along to sort primes from non-primes
    If testPrimes(Cells(y, z)) = True Then
            With Cells(y, z)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
            End With
    End If
Next z
Next y

    'restore excel parameters
    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub


