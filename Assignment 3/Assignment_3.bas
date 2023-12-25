Attribute VB_Name = "Module1"
Option Explicit

' NOTE: You need to only complete ONE of the following functions to get
' credit for Assignment 3

Function medication()
'Place your code here
End Function

Function payment(P As Double, i As Double, n As Integer) As Double
    Dim monthlyInterestRate As Double
    Dim numberOfPayments As Integer
    
    monthlyInterestRate = i / 12
    numberOfPayments = n * 12
    
    payment = P * (monthlyInterestRate / (1 - (1 + monthlyInterestRate) ^ -numberOfPayments))
End Function
