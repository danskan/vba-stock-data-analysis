Attribute VB_Name = "Module1"
Sub StockAnalysis():

' --------------------------------------
' Create Column Titles
' --------------------------------------
For Each k In Worksheets

    Dim Titles As String
    Dim ColumnTitles() As String

    Titles = "Ticker, Annual Change, Percent Change, Total Volume"
    ColumnTitles = Split(Titles, ",")

    k.Cells(1, 9) = ColumnTitles(0)
    k.Cells(1, 10) = ColumnTitles(1)
    k.Cells(1, 11) = ColumnTitles(2)
    k.Cells(1, 12) = ColumnTitles(3)


' ---------------------------------------
' Collecting Data and Placing Into Variables to be Written Later
' ---------------------------------------


    Dim counter As Double
    Dim ticker_symbol As String
    Dim open_price As Double
    Dim close_price As Double
    Dim total_volume As Double
    Dim annual_change As Double
    Dim percent_change As Double


    last_row = k.Cells(Rows.Count, 1).End(xlUp).Row


    k.Columns("H").EntireColumn.Interior.ColorIndex = 1
    counter = 1
    total_volume = 0
    output_row = 2




    For i = 2 To last_row
    
        If k.Cells(i + 1, 1).Value <> k.Cells(i, 1).Value Then
            ticker_symbol = k.Cells(i, 1).Value
            counter = counter + 1
            open_price = k.Cells(counter, 3).Value
            close_price = k.Cells(i, 6).Value

        
            For j = counter To i
                total_volume = total_volume + k.Cells(j, 7).Value
            Next j
        
            If open_price = 0 Then
                percent_change = 0
            Else
                annual_change = close_price - open_price
                percent_change = annual_change / open_price

            End If
    
' ----------------------------------------------------------------
' Writing the Data to the Output Chart
' ----------------------------------------------------------------

            k.Range("I" & output_row).Value = ticker_symbol
            k.Range("J" & output_row).Value = annual_change
            k.Range("K" & output_row).Value = percent_change
            k.Range("L" & output_row).Value = total_volume
        
            output_row = output_row + 1
        
            total_volume = 0
            annual_change = 0
            percent_change = 0
            counter = i
    
        End If

        

    Next i


' --------------------------------------------------------------------
' Apply Conditional Formatting In Order to Highlight Gains and Losses
' -------------------------------------------------------------------
    Dim output_table_last_row As Double
    output_table_last_row = 10000

    For i = 2 To output_table_last_row
        If k.Cells(i, 10).Value > 0 And k.Cells(i, 11).Value > 0 Then
            k.Cells(i, 10).Interior.ColorIndex = 4
            k.Cells(i, 11).Interior.ColorIndex = 4
        Else
            k.Cells(i, 10).Interior.ColorIndex = 3
            k.Cells(i, 11).Interior.ColorIndex = 4
        End If

    Next i
    
' ---------------------------------------------------------
' Format Output Table Data to Prepare for Processing
' ---------------------------------------------------------

    k.Range("K2:K" & output_table_last_row).NumberFormat = "0.00%"
    k.Range("P2:P3").NumberFormat = "0.00%"
    k.Range("L3:L" & output_table_last_row).NumberFormat = "#,##0"
    
    
    
' ----------------------------------------------------------
' Building a Summary Table - Labels First
' -----------------------------------------------------------
    Dim summary_labels As String
    Dim summary_row_titles() As String
    Dim summary_headers As String
    Dim summary_column_headers() As String
    
    summary_labels = "Best % Performer, Worst % Performer, Greatest Volume"
    summary_row_titles = Split(summary_labels, ",")
    
    summary_headers = "Ticker, Value"
    summary_column_headers = Split(summary_headers, ",")
    
    k.Range("N2").Value = summary_row_titles(0)
    k.Range("N3").Value = summary_row_titles(1)
    k.Range("N4").Value = summary_row_titles(2)
    k.Range("O1").Value = summary_column_headers(0)
    k.Range("P1").Value = summary_column_headers(1)
    
' -------------------------------------------------------------
' Traverse the Output Chart, Retrieve Summary Table Data, & Write to Table
' --------------------------------------------------------------


    Dim best_performer As Double
    Dim worst_performer As Double
    Dim greatest_volume As Double
    
    best_perfomer = 0
    worst_performer = 0
    greatest_volume = 0
    
    For i = 2 To output_table_last_row
        If k.Cells(i, 11).Value > best_performer Then
            best_performer = k.Cells(i, 11).Value
            k.Range("P2").Value = best_performer
            k.Range("O2").Value = k.Cells(i, 11 - 2).Value
        End If
        If k.Cells(i, 11).Value < worst_performer Then
            worst_performer = k.Cells(i, 11).Value
            k.Range("P3").Value = worst_performer
            k.Range("O3").Value = k.Cells(i, 11 - 2).Value
        End If
        If k.Cells(i, 12).Value > greatest_volume Then
            greatest_volume = k.Cells(i, 12).Value
            k.Range("P4").Value = greatest_volume
            k.Range("O4").Value = k.Cells(i, 12 - 3).Value
        End If
        
    Next i
    
' ---------------------------------------
' Format Greatest Volume Output to Human Readable Format
' ---------------------------------------
    k.Range("P4").NumberFormat = "#,##0"

' ---------------------------------------
' Autofit Column Width to Fit Data
' ---------------------------------------

    k.Columns("A:T").EntireColumn.AutoFit


Next k

End Sub
