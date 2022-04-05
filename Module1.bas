Attribute VB_Name = "Module1"
Sub Dosomething()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call The_VBA_of_Wall_Street
    Next
    Application.ScreenUpdating = True
End Sub

Sub The_VBA_of_Wall_Street()

'Headers & Labels
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Precent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    
'Declaring Variables
Dim Ticker As String
Dim Volume_Total As Double
    Volume_Total = 0
Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
Dim Open_Amt As Double
    Open_Amt = Cells(2, 3).Value
Dim Close_Amt As Double
    

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        Ticker = Cells(i, 1).Value
        Volume_Total = Volume_Total + Cells(i, 7).Value
        Close_Amt = Cells(i, 6).Value
            
        'Print the name of the Ticker
        Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print the Total Stock Vloume
        Range("L" & Summary_Table_Row).Value = Volume_Total
        
        'Print the Yearly Change.
        Range("J" & Summary_Table_Row).Value = Close_Amt - Open_Amt
        
        'Calculate and Print the precent range. alternative method to .Value to .NumberFormat = "0.00%"
        'This if statement is protect from dividing by zero
        If Open_Amt = 0 Then
            Range("K" & Summary_Table_Row).Value = 0 * 100 & "%"
            Else
            Range("K" & Summary_Table_Row).Value = (Close_Amt - Open_Amt) / (Open_Amt) * 100 & "%"
        End If
                    
        Summary_Table_Row = Summary_Table_Row + 1
        Volume_Total = 0
         
        Open_Amt = Cells(i + 1, 3)
        
    Else
        Volume_Total = Volume_Total + Cells(i, 7).Value
        
    End If

Next i

'Color Formating
For i = 2 To LastRow
    If (Cells(i, 10).Value > 0) Then
        Cells(i, 10).Interior.ColorIndex = 4
    
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If
Next i

'Bonus Section
Dim Goat As Double
    Goat = Cells(2, 11).Value
Dim GoatRow As Integer

For i = 2 To LastRow

    If Cells(i, 11).Value > Goat Then
        Goat = Cells(i, 11).Value
        GoatRow = i
    End If

Next i

    Range("Q" & 2).Value = Goat
    Range("P" & 2).Value = Range("I" & GoatRow)
'----------------------------------------------------------
Dim NotGoat As Double
    NotGoat = Cells(2, 11).Value
Dim NotGoatRow As Integer

For i = 2 To LastRow
    
    If Cells(i, 11).Value < NotGoat Then
        NotGoat = Cells(i, 11).Value
        NotGoatRow = i
    End If

Next i

    Range("Q" & 3).Value = NotGoat
    Range("P" & 3).Value = Range("I" & NotGoatRow)
'------------------------------------------------------------
Dim GTV As Double
    GTV = Cells(2, 12).Value
Dim GTVRow As Integer
    
For i = 2 To LastRow

    If Cells(i, 12).Value > GTV Then
        GTV = Cells(i, 12).Value
        GTVRow = i
    End If

Next i

    Range("Q" & 4).Value = GTV
    Range("P" & 4).Value = Range("I" & GTVRow)
        
End Sub

