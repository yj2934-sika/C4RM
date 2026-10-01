Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public Function getBondPrice( _
    ByVal y As Double, _
    ByVal face As Double, _
    ByVal couponRate As Double, _
    ByVal m As Double, _
    Optional ByVal ppy As Long = 1) As Double

    Dim period As Long
    Dim numberOfPayments As Long
    Dim couponPayment As Double
    Dim discountRate As Double
    Dim cashFlow As Double
    Dim price As Double

    numberOfPayments = m * ppy
    couponPayment = face * couponRate / ppy
    discountRate = y / ppy

    price = 0

    For period = 1 To numberOfPayments

        cashFlow = couponPayment

        If period = numberOfPayments Then
            cashFlow = cashFlow + face
        End If

        price = price + cashFlow / _
            (1 + discountRate) ^ period

    Next period

    getBondPrice = price

End Function


Public Function getBondDuration( _
    ByVal y As Double, _
    ByVal face As Double, _
    ByVal couponRate As Double, _
    ByVal m As Double) As Double

    Dim year As Long
    Dim numberOfYears As Long
    Dim couponPayment As Double
    Dim cashFlow As Double
    Dim presentValue As Double
    Dim totalPresentValue As Double
    Dim weightedTime As Double

    numberOfYears = m
    couponPayment = face * couponRate

    totalPresentValue = 0
    weightedTime = 0

    For year = 1 To numberOfYears

        cashFlow = couponPayment

        If year = numberOfYears Then
            cashFlow = cashFlow + face
        End If

        presentValue = cashFlow / (1 + y) ^ year

        totalPresentValue = totalPresentValue + presentValue
        weightedTime = weightedTime + year * presentValue

    Next year

    getBondDuration = weightedTime / totalPresentValue

End Function

