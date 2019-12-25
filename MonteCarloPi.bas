Attribute VB_Name = "MonteCarloPi"
Option Explicit

Public Function RandomNumber() As Double

    Randomize
    
    RandomNumber = Rnd

End Function


Public Function calculatePi(n As Long) As Double

    Dim x As Double, y As Double
    Dim i As Long, c As Long
    c = 0
                
    Randomize
    
    For i = 1 To n
        x = Rnd
        y = Rnd
        If (x ^ 2 + y ^ 2 <= 1#) Then
            c = c + 1
        End If
    Next
    
    calculatePi = 4# * c / n

End Function
