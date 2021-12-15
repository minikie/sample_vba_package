Sub test()

Dim S As Double, x As Double, r As Double, d As Double, v As Double, T As Double
Dim out_barrier As Double

S = 1112.3
x = 1158
r = 0.00721719
d = 0.00171989
v = 0.069
T = 0.263013698630137

out_barrier = 1112#

CallPrice = VanillaCallOptionPrice(S, x, d, r, v, T)
PutPrice = VanillaPutOptionPrice(S, x, d, r, v, T)
AONCallPrice = AssetOrNothingCallPrice(S, out_barrier, d, r, v, T)
AONPutPrice = AssetOrNothingPutPrice(S, out_barrier, d, r, v, T)
DCPrice = x * DigitalCallPrice(S, out_barrier, d, r, v, T)
DPPrice = x * DigitalPutPrice(S, out_barrier, d, r, v, T)

Debug.Print (CallPrice)
Debug.Print (PutPrice)
Debug.Print (AONCallPrice)
Debug.Print (AONPutPrice)
Debug.Print (DCPrice)
Debug.Print (DPPrice)

End Sub
Private Function VanillaCallPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

Dim dt As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + 0.5 * v ^ 2) * T) / dt
d2 = d1 - dt

ND1 = WorksheetFunction.NormSDist(d1)
ND2 = WorksheetFunction.NormSDist(d2)

VanillaCallPrice = (Exp(-d * T) * S * ND1) - (x * Exp(-r * T) * ND2)


End Function

Private Function VanillaPutPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

Dim dt As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + 0.5 * v ^ 2) * T) / dt
d2 = d1 - dt

NND1 = WorksheetFunction.NormSDist(-d1)
NND2 = WorksheetFunction.NormSDist(-d2)

VanillaPutPrice = (-S * Exp(-d * T) * NND1) + (x * Exp(-r * T) * NND2)


End Function

Private Function AssetOrNothingCallPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

b = r - d
dt = v * Sqr(T)
d1 = (Log(S / x) + (b + (v ^ 2) / 2) * T) / (v * Sqr(T))
d2 = d1 - v * Sqr(T)

ND1 = WorksheetFunction.NormSDist(d1)
ND2 = WorksheetFunction.NormSDist(d2)

AssetOrNothingCallPrice = S * Exp((b - r) * T) * ND1


End Function

Private Function AssetOrNothingPutPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

b = r - d
dt = v * Sqr(T)
d1 = (Log(S / x) + (b + (v ^ 2) / 2) * T) / (v * Sqr(T))
d2 = d1 - v * Sqr(T)
NND1 = WorksheetFunction.NormSDist(-d1)
NND2 = WorksheetFunction.NormSDist(-d2)

AssetOrNothingPutPrice = S * Exp((b - r) * T) * NND1

End Function

Private Function DigitalCallPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + v ^ 2 / 2) * T) / (v * Sqr(T))
d2 = d1 - v * Sqr(T)
ND1 = WorksheetFunction.NormSDist(d1)
ND2 = WorksheetFunction.NormSDist(d2)
DigitalCallPrice = Exp(-r * T) * ND2

End Function


Private Function DigitalPutPrice(S As Double, x As Double, d As Double, r As Double, v As Double, T As Double) As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + v ^ 2 / 2) * T) / (v * Sqr(T))
d2 = d1 - v * Sqr(T)
NND1 = WorksheetFunction.NormSDist(-d1)
NND2 = WorksheetFunction.NormSDist(-d2)
DigitalPutPrice = Exp(-r * T) * NND2

End Function

Private Function GapCallPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + (0.5 * v ^ 2)) * T) / dt
d2 = d1 - dt

ND1 = WorksheetFunction.NormSDist(d1)
ND2 = WorksheetFunction.NormSDist(d2)

GapCallPrice = (Exp(-d * T) * S * ND1) - ((x + barrier) * Exp(-r * T) * ND2)

End Function

Private Function GapPutPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double

dt = v * Sqr(T)
d1 = (Log(S / x) + (r - d + (0.5 * v ^ 2)) * T) / dt
d2 = d1 - dt
NND1 = WorksheetFunction.NormSDist(-d1)
NND2 = WorksheetFunction.NormSDist(-d2)

GapPutPrice = ((x - barrier) * Exp(-r * T) * NND2) - (Exp(-d * T) * S * NND1)

End Function


Private Function UpOutCallPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double

' this is for fx knock out forward
' condition : x < barrier , barrier is checked at maturity date only
' call - aon + con
UpOutCallPrice = VanillaCallPrice(S, x, d, r, v, T) - AssetOrNothingCallPrice(S, barrier, d, r, v, T) + x * DigitalCallPrice(S, barrier, d, r, v, T)

End Function
Private Function UpOutPutPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double
' this is for fx knock out forward
' condition x < barrier - useless barrier , barrier is checked at maturity date only
' put
UpOutPutPrice = VanillaPutPrice(S, x, d, r, v, T)

End Function

Private Function DownOutCallPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double
' this is for fx knock out forward
' condition : barrier < x  , useless barrier , barrier is checked at maturity date only
' call
DownOutCallPrice = VanillaCallPrice(S, x, d, r, v, T)

End Function
Private Function DownOutPutPrice(S As Double, x As Double, barrier As Double, d As Double, r As Double, v As Double, T As Double) As Double
' this is for fx knock out forward
' condition : barrier < x  , barrier is checked at maturity date only
' put + aon - con
DownOutPutPrice = VanillaPutPrice(S, x, d, r, v, T) + AssetOrNothingPutPrice(S, barrier, d, r, v, T) - x * DigitalPutPrice(S, barrier, d, r, v, T)

End Function

Function OptionCalculator(option_type As String, S As Double, strike As Double, rf As Double, div As Double, vol As Double, T As Double, result_type As String) As Variant

' option_type : Call, Put, AssetOrNothingCall, AssetOrNothingPut, CashOrNothingCall, CashOrNothingPut

Dim res As Variant

option_type_upp = UCase(option_type)
result_type_upp = UCase(result_type)

' Vanilla option
If option_type_upp = "CALL" Or option_type_upp = "C" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = VanillaCallPrice(S, strike, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "ASSETORNOTHINGCALL" Or option_type_upp = "AONCALL" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = AssetOrNothingCallPrice(S, strike, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If

ElseIf option_type_upp = "PUT" Or option_type_upp = "P" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = VanillaPutPrice(S, strike, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "ASSETORNOTHINGPUT" Or option_type_upp = "AONPUT" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = AssetOrNothingPutPrice(S, strike, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If

Else
    res = "Unknown option_type - " & option_type
End If

OptionCalculator = res

End Function
Function OptionCalculator2(option_type As String, S As Double, strike As Double, barrier As Double, rf As Double, div As Double, vol As Double, T As Double, result_type As String) As Variant

Dim res As Variant

option_type_upp = Replace(UCase(option_type), "_", "")
result_type_upp = Replace(UCase(result_type), "_", "")

' Vanilla option
If option_type_upp = "GAPCALL" Or option_type_upp = "GC" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = GapCallPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "GAPPUT" Or option_type_upp = "GP" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = GapPutPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "UPOUTCALL" Or option_type_upp = "UPANDOUTCALL" Or option_type_upp = "UOC" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = UpOutCallPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "DOWNOUTCALL" Or option_type_upp = "DOWNANDOUTCALL" Or option_type_upp = "DOC" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = DownOutCallPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "UPOUTPUT" Or option_type_upp = "UPANDOUTPUT" Or option_type_upp = "UOP" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = UpOutPutPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
ElseIf option_type_upp = "DOWNOUTPUT" Or option_type_upp = "DOWNANDOUTPUT" Or option_type_upp = "DOP" Then
    If result_type_upp = "NPV" Or result_type_upp = "PRICE" Then
        res = DownOutPutPrice(S, strike, barrier, rf, div, vol, T)
    Else
        res = "Unknown Result Type - " & result_type
    End If
Else
    res = "Unknown option_type - " & option_type
End If

OptionCalculator2 = res
End Function
