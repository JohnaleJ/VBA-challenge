Sub stock_market_analysis():

    ' Declare variable to hold the Ticker Name
    Dim TickerName As String
    ' Declare variable to hold the Yearly Change Value
    Dim Yr_Change As Double
    ' Declare variable to hold the First Open Price per Ticker
    Dim First_Open_Price As Double
    ' Declare variable to hold the Last Close Pricer per Ticker
    Dim Last_Close_Price As Double
    'Declare variable to hold row count per Ticker
    Dim Row_Count As Long
    ' Declare variable to hold the Percent Change
    Dim Percent As Double
    ' Declare variable to hold the Total Stock Volume Double
    Dim Total As Double
    ' Declare variable to hold the location of the Summary Table
    Dim Summary_Table_Location As Double
    ' Declare variable to determine the last row
    Dim LastRow As Long
    
 
    ' Determine Each Header Value and Insert in Each Header location
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    ' Determine the LastRow using a method found at https://excel.officetuts.net/en/vba/count-rows-in-excel-vba
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
   
    ' Determine the initial value for Summary_Table_Location
    Summary_Table_Location = 2
    Row_Count = 0
    
    ' Loop through all rows of data 
    For i = 2 To LastRow  
        ' Add to Total for each row
        Total = Total + Cells(i, 7).Value

        Row_Count = Row_Count + 1                           

        ' Check if we are still in the same Ticker Name, if it is not...
        If Cells(i+1, 1).Value <> Cells(i, 1).Value Then
            
            ' Set the Ticker Name
            TickerName = Cells(i, 1).Value

            'Grab the Last Close Pricer per Ticker
            Last_Close_Price = Cells(i, 6). Value

            'Grab the First Open Price per Ticker
            First_Open_Price = Cells(i+ 1 - Int(Row_Count), 3).Value

            'Calculate Yr_Change
            Yr_Change = Last_Close_Price - First_Open_Price

            'Calculate Percent Change
                If First_Open_Price > 0 Then
                    Percent = Yr_Change / First_Open_Price
                End If
            ' Print the Ticker Name in the Summary Table
            Range("I" & Summary_Table_Location).Value = TickerName
            
            'Print the Yr_Change
            Range("J" & Summary_Table_Location).Value = Yr_Change

            'Print the Percent Change
            Range("K" & Summary_Table_Location).Value = Percent

            'Print the Total Amount to The Summary Table
            Range("L" & Summary_Table_Location).Value = Total

            'Create Color Formatting for Yr_Change Summary Column
                If Yr_Change > 0 Then
                    ' Set the Yr_Change to Green If >0
                    Range("J" & Summary_Table_Location).Interior.ColorIndex = 4
                Else 
                    ' Set the Yr_Change to Red If <0
                    Range("J" & Summary_Table_Location).Interior.ColorIndex = 3
                End If

            'Create Percent Format for Percent Summary Column gathered from https://docs.microsoft.com/en-us/office/vba/api/excel.range.numberformat
            
            Range("K" & Summary_Table_Location).NumberFormat = "###.##%"
          
            'Add one to the Summary Table Row to prepare to gather information for the next iteration
            Summary_Table_Location = Summary_Table_Location + 1

            'Reset the total for the next iteration
            Total = 0

            'Reset the Row Count for the next iteration
            Row_Count = 0

              
        End If

    Next i

End Sub