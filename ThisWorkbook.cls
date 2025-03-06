VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CustTotalPerEmployeePopUp()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim totalVisits As Long
    Dim i As Long
    Dim employee As Variant ' Change from String to Variant
    Dim dict As Object
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("my_sheet") ' Change "my_sheet" to your sheet name with the data
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).row
    
    ' Initialize dictionary to store visit counts per employee
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the rows and sum the visit counts per employee
    For i = 2 To lastRow ' Assuming the data starts from row 2
        employee = ws.Cells(i, 7).Value ' Column G contains the employee names
        If IsNumeric(ws.Cells(i, 8).Value) Then ' Column H contains the visit counts
            If Not dict.exists(employee) Then
                dict.Add employee, ws.Cells(i, 8).Value
            Else
                dict(employee) = dict(employee) + ws.Cells(i, 8).Value
            End If
        End If
    Next i
    
    ' Display the total visits per employee in a message box
    Dim msg As String
    msg = "Total Customer Visits for the Day per Employee:" & vbCrLf
    For Each employee In dict.Keys
        msg = msg & employee & ": " & dict(employee) & vbCrLf
    Next employee
    
    MsgBox msg
End Sub
Sub FindTopEmployee()

    Dim ws As Worksheet
    Dim dataWs As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim employeeDict As Object
    Dim rno_district As Object
    Dim branch As Variant
    Dim employee As Variant
    Dim totalByDate As Variant
    Dim totalByHour As Variant
    Dim maxCustomers As Long
    Dim topEmployee As String
    Dim customers As Long
    Dim maxDate As String
    Dim maxHourDate As String
    
    ' Set the data worksheet
    On Error Resume Next
    Set dataWs = ThisWorkbook.Sheets("my_sheet") ' Change to your sheet name
    On Error GoTo 0
    
    If dataWs Is Nothing Then
        MsgBox "Sheet 'my_sheet' does not exist. Please check the sheet name."
        Exit Sub
    End If
    
    ' Find the last row
    lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).row
    
    ' Set the data range (assuming data starts at row 2 and column 1)
    Set dataRange = dataWs.Range("A2:H" & lastRow) ' Adjust columns as needed
    
    ' Initialize dictionaries
    Set rno_district = CreateObject("Scripting.Dictionary")
    
    ' Loop through the data
    For Each cell In dataRange.Columns(3).Cells ' Column 3 (C): Branch Number
        branch = CStr(cell.Value) ' Treat as text
        employee = CStr(cell.Offset(0, 4).Value) ' Treat as text (Column 7, G: Employee Name)
        totalByDate = CStr(cell.Offset(0, 2).Value) ' Treat as text (Column 5, E: Date)
        totalByHour = CStr(cell.Offset(0, 3).Value) ' Treat as text (Column 6, F: Hour)
        customers = Val(cell.Offset(0, 5).Value) ' Treat as number (Column 8, H: Customers)
        
        If Not rno_district.exists(branch) Then
            Set rno_district(branch) = CreateObject("Scripting.Dictionary")
        End If
        
        If Not rno_district(branch).exists(employee) Then
            Set rno_district(branch)(employee) = CreateObject("Scripting.Dictionary")
            Set rno_district(branch)(employee)("DailyTotals") = CreateObject("Scripting.Dictionary")
            Set rno_district(branch)(employee)("HourlyTotals") = CreateObject("Scripting.Dictionary")
            rno_district(branch)(employee)("Total") = 0 ' SUGGEST RENAME TO: Monthly or monthlyTotal
        End If
        
    rno_district(branch)(employee)("Total") = rno_district(branch)(employee)("Total") + customers
        
        If Not rno_district(branch)(employee)("DailyTotals").exists(totalByDate) Then
            rno_district(branch)(employee)("DailyTotals")(totalByDate) = 0
        End If
        
        If Not rno_district(branch)(employee)("HourlyTotals").exists(totalByHour & totalByDate) Then
            rno_district(branch)(employee)("HourlyTotals")(totalByHour & totalByDate) = 0
        End If
        
    rno_district(branch)(employee)("DailyTotals")(totalByDate) = rno_district(branch)(employee)("DailyTotals")(totalByDate) + customers
    rno_district(branch)(employee)("HourlyTotals")(totalByHour & totalByDate) = rno_district(branch)(employee)("HourlyTotals")(totalByHour & totalByDate) + customers
    Next cell
    
    ' Check if the result sheet exists, delete if it does
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Top Employees")
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create a new sheet for the results
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Top Employees"
    Set ws = Sheets("Top Employees")
    
    ' Output the results
    ws.Cells(1, 1).Value = "Branch"
    ws.Cells(1, 2).Value = "Employee"
    ws.Cells(1, 3).Value = "Total Customers"
    ws.Cells(1, 4).Value = "Most Customers in One Day"
    ws.Cells(1, 5).Value = "Date"
    ws.Cells(1, 6).Value = "Most Customers in One Hour"
    ws.Cells(1, 7).Value = "Hour"
    ws.Cells(1, 8).Value = "Hour Date"
    
    Dim row As Long
    row = 2
    
    For Each branch In rno_district
        For Each employee In rno_district(branch)
            ws.Cells(row, 1).Value = branch
            ws.Cells(row, 2).Value = employee
            ws.Cells(row, 3).Value = rno_district(branch)(employee)("Total")
            
            ' Find the most customers in one day
            maxCustomers = 0
            maxDate = ""
            For Each totalByDate In rno_district(branch)(employee)("DailyTotals")
                If rno_district(branch)(employee)("DailyTotals")(totalByDate) > maxCustomers Then
                    maxCustomers = rno_district(branch)(employee)("DailyTotals")(totalByDate)
                    maxDate = totalByDate
                End If
            Next totalByDate
            ws.Cells(row, 4).Value = maxCustomers
            ws.Cells(row, 5).Value = maxDate
            
            ' Find the most customers in one hour and date
            maxCustomers = 0
            maxHourDate = ""
            For Each totalByHour In rno_district(branch)(employee)("HourlyTotals")
                If rno_district(branch)(employee)("HourlyTotals")(totalByHour) > maxCustomers Then
                    maxCustomers = rno_district(branch)(employee)("HourlyTotals")(totalByHour)
                    maxHourDate = totalByHour
                End If
            Next totalByHour
            ws.Cells(row, 6).Value = maxCustomers
            ws.Cells(row, 7).Value = Left(maxHourDate, Len(maxHourDate) - Len(maxDate))
            ws.Cells(row, 8).Value = Right(maxHourDate, Len(maxDate))
            
            row = row + 1
        Next employee
    Next branch
    
    ' Add sorting controls to the headers
    ws.Range("A1:H1").AutoFilter
    
    MsgBox "Analysis complete. Check the 'Top Employees' sheet for details."
    
End Sub
