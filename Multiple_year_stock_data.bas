Attribute VB_Name = "Module1"
Sub Combine()


'Create new sheet
Sheets.Add.Name = "Combined_Data"
Sheets("Combined_Data").Move Before:=Sheets(1)
Set combined_sheet = Worksheets("Combined_Data")
    
    ' Loop through all sheets
    For Each ws In Worksheets
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1
        lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowState - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value
    Next ws
    
combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A1:G1").Value
combined_sheet.Columns("A:G").AutoFit
    
End Sub


Sub Ticker()

' Define Ticker Id
Dim Ticker_Id As String

' Define Ticker Volume
Dim Ticker_Volume As Double
Ticker_Volume = 0

' Track summary table row
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Define beginning open price
Dim Open_Price As Double
Open_Price = Range("C2").Value
  
' Define beginnging close price
Dim Close_Price As Double
Close_Price = 0


lastRow = Cells(Rows.Count, "A").End(xlUp).Row + 1

    ' Loop through data rows
    For i = 2 To lastRow
  

        ' Check if previous cell matches currnet cell
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ' Set ticker Id and close price
        Ticker_Id = Cells(i, 1).Value
        Close_Price = Cells(i, 6).Value

        ' Add to the Ticker Volume
        Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

        ' Print in the Summary Table
        Range("I" & Summary_Table_Row).Value = Ticker_Id
        Range("L" & Summary_Table_Row).Value = Ticker_Volume
        Range("J" & Summary_Table_Row).Value = Close_Price - Open_Price
        Range("K" & Summary_Table_Row).Value = Close_Price / Open_Price
      
            ' Format "J" column
            If Range("J" & Summary_Table_Row).Value < 0 Then
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            Else
                Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            End If

        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the values
        Ticker_Volume = 0
        Open_Price = 0
        Close_Price = 0
        ' Set open price to new open price for next Id
        Open_Price = Cells(i + 1, 3)
      
      
      

        ' If cell following is the same
        Else
            ' Add to ticker volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value
      
      
        End If
    Next i
End Sub

