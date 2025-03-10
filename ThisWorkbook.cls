' This Subroutine sums the total of customers helped by hour per day by employee for the range of data provided.
' Version 0.1.2
' Copywright 2025 Alex Flores

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub FindTopEmployee()
    Dim newWorksheet As Worksheet
    Dim dataWorksheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim district As Object
    Dim branch As Variant
    Dim employee As Variant
    Dim theDay As Variant
    Dim theHour As Variant
    Dim maxCustomers As Long
    Dim customers As Long
    Dim maxDate As String
    Dim maxHourDate As String
    
    ' Set the worksheet with the data to process to: dataWorksheet. Change "mySheet" to your sheet name.
    On Error Resume Next ' Supress errors if "mySheet" does not exist and continue with the script.
    Set dataWorksheet = ThisWorkbook.Sheets("mySheet") ' Attempts to set referrence to "mySheet."
    On Error GoTo 0 ' Make Errors Great Again. MEGA. Reset error handling.
    If dataWorksheet Is Nothing Then ' Check to see is dataWorksheet was properly set.
        MsgBox "Sheet 'mySheet' does not exist. Please check the sheet name." ' Make the pop-up appear if fail.
        Exit Sub ' Gracefuly exit the subroutine and try again.
    End If
    
    ' Find the last row
    lastRow = dataWorksheet.Cells(dataWorksheet.Rows.Count, "A").End(xlUp).row
    
    ' Set the data range (assuming data starts at row 2 and column 1)
    Set dataRange = dataWorksheet.Range("A2:H" & lastRow) ' Adjust columns as needed
    
    ' Initialize dictionary to store employee customer totals by branch.
    Set district = CreateObject("Scripting.Dictionary")
    
    ' Loop through the data
    For Each cell In dataRange.Columns(3).Cells ' <-- Column C
        branch = CStr(cell.Value) ' Treat as text; Column 3, C: Branch Number
        employee = CStr(cell.Offset(0, 4).Value) ' Treat as text; Column 7, G: Employee Name
        theDay = CStr(cell.Offset(0, 2).Value) ' Treat as text;Column 5, E: Date
        theHour = CStr(cell.Offset(0, 3).Value) ' Treat as text; Column 6, F: Hour
        customers = Val(cell.Offset(0, 5).Value) ' Treat as number; Column 8, H: Customers
        ''' ADD A VAR FOR CAPTURING THE COLUMN WITH THE MONTH '''
        
        ' Initialize dictionary in district, if it has not already, to store employee customer totals by employee
        If Not district.exists(branch) Then
            Set district(branch) = CreateObject("Scripting.Dictionary")
        End If
        
        ' Initialize dictionary in branch.employee dictionary to hold customer totals in dailyTotal and hourlyTotal dictionaries.
        If Not district(branch).exists(employee) Then
            Set district(branch)(employee) = CreateObject("Scripting.Dictionary")
            Set district(branch)(employee)("dailyTotal") = CreateObject("Scripting.Dictionary")
            Set district(branch)(employee)("hourlyTotal") = CreateObject("Scripting.Dictionary")
            district(branch)(employee)("monthlyTotal") = 0 ' To hold employees' monthly customer total.
        End If
        
    district(branch)(employee)("monthlyTotal") = district(branch)(employee)("monthlyTotal") + customers
        
        If Not district(branch)(employee)("dailyTotal").exists(theDay) Then
            district(branch)(employee)("dailyTotal")(theDay) = 0
        End If
        
        If Not district(branch)(employee)("hourlyTotal").exists(theHour & theDay) Then
            district(branch)(employee)("hourlyTotal")(theHour & theDay) = 0
        End If
        
    ' Create unique keys for the dictionaries: dailyTotal and hourlyTotal by using theDay and theHour strings from columns E: Date & F: Hour from mySheet.
    district(branch)(employee)("dailyTotal")(theDay) = district(branch)(employee)("dailyTotal")(theDay) + customers ' Summing the total of var customers by day key.
    district(branch)(employee)("hourlyTotal")(theHour & theDay) = district(branch)(employee)("hourlyTotal")(theHour & theDay) + customers ' Summing the total of var customers by theHour AND theDay key.
    Next cell
    
    ' Check if the result sheet exists, delete if it does
    On Error Resume Next
    Set newWorksheet = ThisWorkbook.Sheets("Top Employees")
    If Not newWorksheet Is Nothing Then
        Application.DisplayAlerts = False
        newWorksheet.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    ' Create a new sheet for the results
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Top Employees"
    Set newWorksheet = Sheets("Top Employees")
    
    ' Output the results NOTE: ADD FORMATING AND MAYBE STYLING. So its not all bunched-up.
    newWorksheet.Cells(1, 1).Value = "Branch"
    newWorksheet.Cells(1, 2).Value = "Employee"
    newWorksheet.Cells(1, 3).Value = "Total Customers"
    newWorksheet.Cells(1, 4).Value = "Most Customers in One Day"
    newWorksheet.Cells(1, 5).Value = "Date"
    newWorksheet.Cells(1, 6).Value = "Most Customers in One Hour"
    newWorksheet.Cells(1, 7).Value = "Hour"
    newWorksheet.Cells(1, 8).Value = "Hour Date"
    
    Dim row As Long
    row = 2
    
    For Each branch In district
        For Each employee In district(branch)
            newWorksheet.Cells(row, 1).Value = branch
            newWorksheet.Cells(row, 2).Value = employee
            newWorksheet.Cells(row, 3).Value = district(branch)(employee)("monthlyTotal")
            
            ' Find the most customers in one day
            maxCustomers = 0
            maxDate = ""
            For Each theDay In district(branch)(employee)("dailyTotal")
                If district(branch)(employee)("dailyTotal")(theDay) > maxCustomers Then
                    maxCustomers = district(branch)(employee)("dailyTotal")(theDay)
                    maxDate = theDay
                End If
            Next theDay
            newWorksheet.Cells(row, 4).Value = maxCustomers
            newWorksheet.Cells(row, 5).Value = maxDate
            
            ' Find the most customers in one hour and date
            maxCustomers = 0
            maxHourDate = ""
            For Each theHour In district(branch)(employee)("hourlyTotal")
                If district(branch)(employee)("hourlyTotal")(theHour) > maxCustomers Then
                    maxCustomers = district(branch)(employee)("hourlyTotal")(theHour)
                    maxHourDate = theHour
                End If
            Next theHour
            newWorksheet.Cells(row, 6).Value = maxCustomers
            newWorksheet.Cells(row, 7).Value = Left(maxHourDate, Len(maxHourDate) - Len(maxDate))
            newWorksheet.Cells(row, 8).Value = Right(maxHourDate, Len(maxDate))
            
            row = row + 1
        Next employee
    Next branch
    
    ' Add sorting controls to the headers
    newWorksheet.Range("A1:H1").AutoFilter
    
    MsgBox "Analysis complete. Check the 'Top Employees' sheet for details."
    
End Sub
Sub CustTotalPerEmployeePopUp()
    Dim newWorksheet As Worksheet
    Dim lastRow As Long
    Dim totalVisits As Long
    Dim i As Long
    Dim employee As Variant ' Change from String to Variant
    Dim dict As Object
    
    ' Set the worksheet
    Set newWorksheet = ThisWorkbook.Sheets("my_sheet") ' Change "my_sheet" to your sheet name with the data
    
    ' Find the last row with data
    lastRow = newWorksheet.Cells(newWorksheet.Rows.Count, "H").End(xlUp).row
    
    ' Initialize dictionary to store visit counts per employee
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Loop through the rows and sum the visit counts per employee
    For i = 2 To lastRow ' Assuming the data starts from row 2
        employee = newWorksheet.Cells(i, 7).Value ' Column G contains the employee names
        If IsNumeric(newWorksheet.Cells(i, 8).Value) Then ' Column H contains the visit counts
            If Not dict.exists(employee) Then
                dict.Add employee, newWorksheet.Cells(i, 8).Value
            Else
                dict(employee) = dict(employee) + newWorksheet.Cells(i, 8).Value
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