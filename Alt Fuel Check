Sub Alt_Fuel_Check()
'Checking if Alternate Fuel for old VIN in new file changed from its Alternate Fuel in old file "LastVmast"

Dim LastVmast As Range
Set LastVmast = Worksheets("Last Vmast").Range("B:CT")

Range("A1").Select
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Check if Alternate Fuel Changed
AltFuel:
For i = 2 To Lastrow
On Error GoTo ErrorHandler4
AltFuelCheck:
If Cells(i, 86).Value <> Application.WorksheetFunction.VLookup(Cells(i, 2), LastVmast, 85, False) Then
    Cells(i, 107).Value = "Alternate Fuel Changed"
End If

Next

'If a record in new file is not found in old file, then it is a new VIN. We note that and move on to next record.
ErrorHandler4:
If i > Lastrow Then
GoTo EndSub
Else
Cells(i, 107).Value = "New VIN"
i = i + 1
Resume AltFuelCheck
End If


EndSub:
End Sub
