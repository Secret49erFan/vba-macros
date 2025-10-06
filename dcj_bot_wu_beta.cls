VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "datasheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub CrunchWUReport()
'
' Parse data from the <INSERT REPORT NAME HERE>.
' Display seperate totals for send and rec.
' Count the number of each.
'
'
' Save the personal workbook
' Future features:
' * Thresholds
' * Frequencies
''''''''''''''''''''''''''''''''''''''''''''''''
    ' Initiate the data worksheet and set as the currently active tab.
  Dim ws As Worksheet
  Set ws = ActiveSheet
    
  ' Initiate the vars the first and last rows of data.
  Dim r As Long, last As Long
    
  ' Initiate the vars for dictionary, type, and amount.
  Dim d As Object, label As String, amt As Double
    
  ' Creat dictionary to hold totals.
  Set d = CreateObject("Scripting.Dictionary")
    
  ' Find the last row of the data.
  last = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
  ' Loops through all available rows.
  For r = 2 To last
    ' Type of item from column B.
    label = ws.Cells(r, "B").Value
        
    ' Dollar amount from column F.
    amt = ws.Cells(r, "F").Value
        
    ' Makes sure label is not blank; checks if it already exists; if not, add it.
    If label <> "" Then
      If d.Exists(label) Then
        d(label) = d(label) + amt
      Else
        d.Add label, amt
      End If
    End If
    Debug.Print "WU Type: " & label
    Debug.Print "Dollar Amount: " & amt
    Debug.Print "---------------"
    Next r
    
    Debug.Print "Unique items: "; d.Count
    Dim k As Variant
    For Each k In d.Keys
        Debug.Print k, FormatCurrency(d(k))
    Next k
    
End Sub