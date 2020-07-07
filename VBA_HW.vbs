Attribute VB_Name = "Module1"
Sub Stock_Markert()
    Dim Last_Value As Double
    Dim Total_Volume As Double
    Dim First_Value As Double
    Dim Total_Value As Double
    Dim Change As Double
    Dim Percent_Change As Double
    Dim LastRow As Long
    Dim ws As Worksheet
    Dim j As Integer
    ' Loop through all sheets
    For Each ws In Worksheets
        ws.Activate
       
    'MsgBox (ws.Name)
    'Find the last row of each worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       
        'Make headers for table
       ws.Cells(1, 8).Value = "Ticker"
       ws.Cells(1, 9).Value = "Yearly Change"
       ws.Cells(1, 10).Value = "Percent Change"
       ws.Cells(1, 11).Value = "Total Stock Volume"
      'Assigns first opening value
      First_Value = ws.Cells(2, 3).Value
      'Assigns first ticker to created table
     ws.Cells(2, 8).Value = ws.Cells(2, 1).Value
     'Keeps track of row of created table
      j = 2
      'loops through data in sheet
   For i = 2 To LastRow
    'checks for different tickers
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
      'Assighs new Ticker to created table
    ws.Cells(j + 1, 8).Value = ws.Cells(i + 1, 1).Value
      'records closing value
    Last_Value = ws.Cells(i, 6).Value
      'calculates change
      Change = Last_Value - First_Value
      'Assigns change to Created table
      Cells(j, 9).Value = Change
      'Calculates Percent Change
    If (First_Value = 0) Then
      ws.Cells(j, 10).Value = "error"
    Else
      Percent_Change = Change / First_Value
      Percent_Change = Percent_Change * 100
      ws.Cells(j, 10).Value = Percent_Change
    End If
      'Assigns Percent Change to Created Table
      ws.Cells(j, 10).Value = Percent_Change
      'Assigns new opening Value
      First_Value = ws.Cells(i + 1, 3).Value
      'Keeps track of Total Volume
      Total_Volume = Volume + Cells(i, 7).Value
      'Assigns Total Volume to new table
      ws.Cells(j, 11).Value = Total_Volume
      j = j + 1
      'resets Total_Volume
      Total_Volume = 0
   Else
      Total_Volume = Volume + Cells(i, 7).Value

    
      End If
    Next
    For k = 2 To j - 1
    'assign red color
    If ws.Cells(k, 9).Value < 0 Then
    Cells(k, 9).Interior.ColorIndex = 3
    'assign green color
    Else
    Cells(k, 9).Interior.ColorIndex = 4
    End If
    Next
    Next ws
   
End Sub

