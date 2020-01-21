Option Explicit "To enforce the definition of the the types of variables

Public Function calculatePi(n As Long) As Double
    Dim x As Double, y As Double
    Dim i As Long, inside As Long
    c = 0
    For i = 1 To n
        x = Rnd
        y = Rnd
        If (x ^ 2 + y ^ 2 <= 1#) Then
            inside = inside + 1
        End If
    Next
    calculatePi = 4 * inside / n
End Function
