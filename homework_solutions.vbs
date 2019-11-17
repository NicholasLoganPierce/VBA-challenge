Sub Final_Solution()

Dim WS As Worksheet
    For Each WS In ActiveWorkbook.Worksheets
    WS.Activate
   
        'Determine Last Row
        Dim LastRow As Long
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row

        'Fill in Ticker Column
        Columns(1).Copy
        Columns(9).PasteSpecial
        ActiveSheet.Range("H1:J" & LastRow).RemoveDuplicates Columns:=2, Header:=xlNo

        'Fill in Total Volume Column
        LastRow = Cells(Rows.Count, 9).End(xlUp).Row
        Range("L2").Value = "=SUMIF(R2C1:R71226C1,RC[-3],R2C7:R71226C7)"
        Range("L2:L" & LastRow).FillDown
        Columns(12).NumberFormat = "#,###"
       
        ' Re-Determine Last Row
        LastRow = WS.Cells(Rows.Count, 1).End(xlUp).Row

        ' Add Column Headings
        Cells(1, "I").Value = "Ticker"
        Cells(1, "J").Value = "Yearly Change"
        Cells(1, "K").Value = "Percent Change"
        Cells(1, "L").Value = "Total Stock Volume"
        'Define Variables
        Dim Volume As Double
        Volume = 0
        Dim Row As Double
        Row = 2
        Dim Column As Double
        Column = 1
        Dim counter As Long
        Dim Open_Price As Double
        Dim Close_Price As Double
        Dim Yearly_Change As Double
        Dim Ticker_Name As String
        Dim Percent_Change As Double
       
       
        'Set Initial Open Price
        Open_Price = Cells(2, Column + 2).Value
         
       
        For counter = 2 To LastRow
            If Cells(counter + 1, Column).Value <> Cells(counter, Column).Value Then
                ' Set Ticker name
                Ticker_Name = Cells(counter, Column).Value
                Cells(Row, Column + 8).Value = Ticker_Name
                ' Set Close Price
                Close_Price = Cells(counter, Column + 5).Value
                ' Add Yearly Change
                Yearly_Change = Close_Price - Open_Price
                Cells(Row, Column + 9).Value = Yearly_Change
                ' Add Percent Change
                If (Open_Price = 0 And Close_Price = 0) Then
                    Percent_Change = 0
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then
                    Percent_Change = 1
                Else
                    Percent_Change = Yearly_Change / Open_Price
                    Cells(Row, Column + 10).Value = Percent_Change
                    Cells(Row, Column + 10).NumberFormat = "0.00%"
                End If
                ' Add Total Volume
                Volume = Volume + Cells(counter, Column + 6).Value
                Cells(Row, Column + 11).Value = Volume
                ' Add one to the summary table row
                Row = Row + 1
                ' reset the Open Price
                Open_Price = Cells(counter + 1, Column + 2)
                ' reset Volume Total
                Volume = 0
            'if cells are the same ticker
            Else
                Volume = Volume + Cells(counter, Column + 6).Value
            End If
        Next counter
       
        ' Determine NewLastRow
        NewLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row
       
        For j = 2 To NewLastRow
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then
                Cells(j, Column + 9).Interior.ColorIndex = 10
            ElseIf Cells(j, Column + 9).Value < 0 Then
                Cells(j, Column + 9).Interior.ColorIndex = 3
            End If
        Next j
       
        ' Set Greatest % Increase, % Decrease, and Total Volume
        Cells(2, Column + 14).Value = "Greatest % Increase"
        Cells(3, Column + 14).Value = "Greatest % Decrease"
        Cells(4, Column + 14).Value = "Greatest Total Volume"
        Cells(1, Column + 15).Value = "Ticker"
        Cells(1, Column + 16).Value = "Value"
       
        For Z = 2 To NewLastRow
            If Cells(Z, Column + 10).Value = Application.WorksheetFunction.Max(WS.Range("K2:K" & NewLastRow)) Then
                Cells(2, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(2, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(2, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 10).Value = Application.WorksheetFunction.Min(WS.Range("K2:K" & NewLastRow)) Then
                Cells(3, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(3, Column + 16).Value = Cells(Z, Column + 10).Value
                Cells(3, Column + 16).NumberFormat = "0.00%"
            ElseIf Cells(Z, Column + 11).Value = Application.WorksheetFunction.Max(WS.Range("L2:L" & NewLastRow)) Then
                Cells(4, Column + 15).Value = Cells(Z, Column + 8).Value
                Cells(4, Column + 16).Value = Cells(Z, Column + 11).Value
            End If
        Next Z
       
    Next WS
       
End Sub

