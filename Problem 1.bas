Attribute VB_Name = "Module1"
Sub test()
For i = 2 To 5000
    Dim Total As Integer
    Dim shopid As Integer
    shopid = Cells(i, "B").Value
    Cells(shopid, "K").Value = Cells(shopid, "K").Value + Cells(i, "D").Value
    Cells(shopid, "L").Value = Cells(shopid, "L").Value + 1
Next
For i = 1 To 100
    Cells(i, "M").Value = Cells(i, "K").Value / Cells(i, "L").Value
Next
Dim max As Double
max = Application.max(Columns(13))
Dim min As Double
max = Application.max(Columns(13))
Dim median As Double
max = Application.median(Columns(13))
MsgBox ("max is " + CStr(max) + ", min is " + CStr(min) + ", median is " + CStr(median))
End Sub

