VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCashFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public dt As Date
Public interest As Double
Public amortization As Double
Public discountFactor As Double


Property Get getAll() As Double()
    Dim result(nrOfCfElements - 1) As Double
    result(0) = interest
    result(1) = interest * discountFactor
    result(2) = amortization
    result(3) = amortization * discountFactor
    result(4) = interest + amortization
    result(5) = (interest + amortization) * discountFactor
    getAll = result
End Property
