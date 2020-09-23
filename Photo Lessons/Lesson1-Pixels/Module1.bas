Attribute VB_Name = "Module1"
'Color Controller Module
'Copyright (c) Marco Samy 2002
'Extract The Value Of Red and Green and Blue from a color
Function ColorRGB(ByVal sColor As Single, ByRef ColorR, ByRef ColorG, ByRef ColorB)
    ColorR = (sColor And 255)
    ColorG = (sColor And 65280) / 256
    ColorB = (sColor And 16711680) / 65536
End Function
'mix two colors
Function ColorPercentage(sColor1, sColor2, Optional sPercentTo1 As Single = 50)
Dim vR, vB, vG, vR1, vB1, vG1, vR2, vB2, vG2
'get RGB value from color 1
ColorRGB sColor1, vR, vG, vB
'get RGB value from color 2
ColorRGB sColor2, vR1, vG1, vB1
'setting Percent Value
If Val(sPercentTo1) > 100 Then sPercentTo1 = 1 Else sPercentTo1 = sPercentTo1 / 100
'Average values with percent
vR2 = (vR * sPercentTo1) + (vR1 * (1 - sPercentTo1))
vG2 = (vG * sPercentTo1) + (vG1 * (1 - sPercentTo1))
vB2 = (vB * sPercentTo1) + (vB1 * (1 - sPercentTo1))
'setting the new color value
ColorPercentage = RGB(vR2, vG2, vB2)
End Function
'getting a negative for a fixed color
Function NegativeColor(sColor)
Dim vR, vB, vG
'first getting the RGB values
ColorRGB sColor, vR, vG, vB
'how to negative a color?
'to negative color we get the Invert of the Values of the Reg and Green and Blue Values
'because the maimum value of any one (Red or Green or Blue) is 255 so we get th negative as the following
'Red = 255 - Red
'Green = 255 - Green
'Blue = 255 - Blue
vR = 255 - vR
vG = 255 - vG
vB = 255 - vB
'setting the new color value
NegativeColor = RGB(vR, vG, vB)
End Function
Function GrayColor(sColor)
Dim vR, vB, vG, vMid
'first getting the RGB values
ColorRGB sColor, vR, vG, vB
'how to gray sacle?
'gray scale is to get the value of the gray from a fixed color
'of course that will make (Red = Green = Blue)
vMid = (0.5 + (0.299 * vR) + (0.587 * vG) + (0.114 * vB))
vR = IIf(Value >= 256, 255, vMid)
vG = IIf(Value >= 256, 255, vMid)
vB = IIf(Value >= 256, 255, vMid)
'setting the new color value
GrayColor = RGB(vR, vG, vB)
End Function
