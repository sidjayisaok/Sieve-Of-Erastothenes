Sub Erastothenes()

'variables
Dim x, a as Integer
Dim y, z, c As Long

'improve efficiency
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
For c = 2 To a
    'matrix manipulation
    Cells(y, z) = (y * a) + (z - a)
        'determine if number is not prime
        If Cells(y, z) Mod c = 0 And Cells(y, z) <> c Then
            With Cells(y, z)
                .Interior.Color = RGB(0, 0, 0)
                .Font.Color = RGB(255, 255, 255)
            End With
        End If
Next c
Next z
Next y

    .Calculation = xlCalculationAutomatic
    .ScreenUpdating = True
    .EnableEvents = True
End With

End Sub
