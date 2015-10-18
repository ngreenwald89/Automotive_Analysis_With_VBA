Sub VSR_Group_Num_Doors_and_Body_Type()
'Checking Doors and Body Type based on new criteria (summer 2015)

'Find last row of data
Range("A1").Select
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Do same process for each row of data
For i = 2 To Lastrow

'Find VSR Group, body type, and door count
vsrGroup = Cells(i, 16).Value
bodyType = Cells(i, 7).Value
doorCount = Cells(i, 9).Value

'For VSR groups ending in a or b
Dim vsrGroupString As String
vsrGroupString = Mid(vsrGroup, 2, 1)

'Redirect to correct VSR Group
If vsrGroupString = "a" Then GoTo GroupA
If vsrGroupString = "b" Then GoTo GroupB
If vsrGroup = xx Or vsrGroup = xy Then GoTo GroupXX
If vsrGroup = yy Then GoTo GroupYY
If vsrGroup = zz Or vsrGroup = zy Then GoTo GroupZZ
If vsrGroup = vv Or vsrGroup = vy Then GoTo GroupVV

'If door count or body type is incorrect for its VSR Group, output error:

'If VSR group ends with a a, then redirects here
GroupA:
'Correct door count for VSR groups ending in a, should be 2 or 3:
If doorCount = 2 Or doorCount = 3 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
'Correct body type for VSR groups ending in a
If bodyType = "CABRI" Or bodyType = "CONV" Or bodyType = "CPE" Or bodyType = "HCHBK" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
'Go to next row of data
GoTo nextRow

GroupB:
If doorCount = 4 Or doorCount = 5 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
If bodyType = "HCHBK" Or bodyType = "SED" Or bodyType = "WAG" Or bodyType = "CPE" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
GoTo nextRow

GroupXX:
If doorCount = 2 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
If bodyType = "CONV" Or bodyType = "CPE" Or bodyType = "HCHBK" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
GoTo nextRow

GroupYY:
If doorCount = 2 Or doorCount = 3 Or doorCount = 4 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
If bodyType = "VAN4X2" Or bodyType = "VAN4X4" Or bodyType = "WAG4X2" Or bodyType = "WAG4X4" Or bodyType = "VAN 4X2" Or bodyType = "VAN 4X4" Or bodyType = "WAG 4X2" Or bodyType = "WAG 4X4" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
GoTo nextRow

GroupZZ:
If doorCount = 2 Or doorCount = 3 Or doorCount = 4 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
If bodyType = "PKP4X2" Or bodyType = "PKP4X4" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
GoTo nextRow

GroupVV:
If doorCount = 2 Or doorCount = 3 Or doorCount = 4 Then
    Cells(i, 100).Value = ""
    Else: Cells(i, 100).Value = "Door Count Error"
End If
If bodyType = "CONV4X2" Or bodyType = "CONV4X4" Or bodyType = "UTL4X2" Or bodyType = "UTL4X4" Or bodyType = "SUT4X2" Or bodyType = "SUT4X4" Or bodyType = "CONV 4X2" Or bodyType = "CONV 4X4" Or bodyType = "UTL 4X2" Or bodyType = "UTL 4X4" Or bodyType = "SUT 4X2" Or bodyType = "SUT 4X4" Then
    Cells(i, 108).Value = ""
    Else: Cells(i, 108).Value = "Body Type Error"
End If
GoTo nextRow

'Go to next row of data
nextRow:
Next

End Sub

