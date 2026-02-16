VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "datasheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub WUReport()
'
' Parse data from the "Western Union Money Transfer Report."
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
    
  ' Initiate the vars for first and last rows of data.
  Dim r As Long, last As Long
    
  ' Initiate the vars for dictionary, type, and amount.
  Dim d As Object, label As String, amt As Double
    
  ' Creat dictionary to hold totals.
  Set d = CreateObject("Scripting.Dictionary")
    
  ' Find the last row of the data.
  last = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  ' Loops through all available rows.
  For r = 5 To last
    ' Type of item from column with label.
    label = ws.Cells(r, "B").Value
        
    ' Dollar amount from column with amount.
    amt = ws.Cells(r, "F").Value
        
    ' Makes sure label is not blank; checks if it already exists; if not, add it.
    If label <> "" Then
      If d.Exists(label) Then
        d(label) = d(label) + amt
      Else
        d.Add label, amt
      End If
    End If
  Next r
  
  'Build the MsgBox
  Dim summary As String
  Dim k As Variant
  
  summary = "Unique labels: " & d.Count & vbNewLine & vbNewLine
  
  For Each k In d.Keys
    summary = summary & k & ":  " & FormatCurrency(d(k)) & vbNewLine
  Next k
  
  MsgBox summary, vbInformation, "WU Totals"
End Sub
Sub MiscReport()
'
' Parse data from the "Western Union Money Transfer Report."
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
    
  ' Initiate the vars for first and last rows of data.
  Dim r As Long, last As Long
    
  ' Initiate the vars for dictionary, type, and amount.
  Dim d As Object, label As String, amt As Double
    
  ' Creat dictionary to hold totals.
  Set d = CreateObject("Scripting.Dictionary")
    
  ' Find the last row of the data.
  last = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  ' Loops through all available rows.
  For r = 5 To last
    ' Type of item from column with label.
    label = ws.Cells(r, "C").Value
        
    ' Dollar amount from column with amount.
    amt = ws.Cells(r, "D").Value
        
    ' Makes sure label is not blank; checks if it already exists; if not, add it.
    If label <> "" Then
      If d.Exists(label) Then
        d(label) = d(label) + amt
      Else
        d.Add label, amt
      End If
    End If
  Next r
  
  'Build the MsgBox
  Dim summary As String
  Dim k As Variant
  
  summary = "Unique labels: " & d.Count & vbNewLine & vbNewLine
  
  For Each k In d.Keys
    summary = summary & k & ":  " & FormatCurrency(d(k)) & vbNewLine
  Next k
  
  MsgBox summary, vbInformation, "WU Totals"
End Sub
Sub GenericReport()
'
' Parse data from the "Western Union Money Transfer Report."
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
    
  ' Initiate the vars for first and last rows of data.
  Dim r As Long, last As Long
    
  ' Initiate the vars for dictionary, type, and amount.
  Dim d As Object, label As String, amt As Double
    
  ' Creat dictionary to hold totals.
  Set d = CreateObject("Scripting.Dictionary")
    
  ' Find the last row of the data.
  last = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
  ' Loops through all available rows.
  For r = 5 To last
    ' Type of item from column with label.
    label = ws.Cells(r, "J").Value
        
    ' Dollar amount from column with amount.
    amt = ws.Cells(r, "L").Value
        
    ' Makes sure label is not blank; checks if it already exists; if not, add it.
    If label <> "" Then
      If d.Exists(label) Then
        d(label) = d(label) + amt
      Else
        d.Add label, amt
      End If
    End If
  Next r
  
  'Build the MsgBox
  Dim summary As String
  Dim k As Variant
  
  summary = "Unique labels: " & d.Count & vbNewLine & vbNewLine
  
  For Each k In d.Keys
    summary = summary & k & ":  " & FormatCurrency(d(k)) & vbNewLine
  Next k
  
  MsgBox summary, vbInformation, "WU Totals"
End Sub
