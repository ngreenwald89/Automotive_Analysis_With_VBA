Sub VSR_Group_Num_Veh_Type()
'Checking Vehicle Type based on new criteria (summer 2015)

'Find last row of data
Range("A1").Select
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'Do same process for each row of data
For i = 2 To Lastrow

'Find VSR Group and vehicle type
vsrGroup = Cells(i, 16).Value
vehType = Cells(i, 6).Value

'Redirect to correct VSR Group
If vsrGroup >= 100 And vsrGroup <= 199 Then GoTo Group100
If vsrGroup = 200 Or vsrGroup = 201 Then GoTo Group200
If vsrGroup = 300 Then GoTo Group300
If vsrGroup = 400 Or vsrGroup = 401 Then GoTo Group400
If vsrGroup = 500 Or vsrGroup = 501 Then GoTo Group500
                
'If vehicle type is incorrect for its VSR Group, output error:
 
Group100:
If vehType <> "PP" Then
    Cells(i, 99).Value = "Vehicle Type Error"
End If
GoTo nextRow

Group200:
If vehType <> "SPT" Then
    Cells(i, 99).Value = "Vehicle Type Error"
End If
GoTo nextRow

Group300:
If vehType = "TRK" Or vehType = "MPV" Or vehType = "BUS" Then
    Cells(i, 99).Value = ""
    Else: Cells(i, 99).Value = "Vehicle Type Error"
End If
GoTo nextRow

Group400:
If vehType <> "TRK" Then
    Cells(i, 99).Value = "Vehicle Type Error"
End If
GoTo nextRow

Group500:
If vehType <> "MPV" Then
    Cells(i, 99).Value = "Vehicle Type Error"
End If
GoTo nextRow

'Go to next row of data
nextRow:
Next


End Sub
