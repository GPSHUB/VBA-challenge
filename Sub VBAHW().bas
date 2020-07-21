Sub VBAHomework()

' Column Headers
'----------------------------
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Change Percent"
Cells(1, 12).Value = "Total Stock Volume"

Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest total volume"

Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"


' Defining variables
'----------------------------
Dim Ticker As String

Dim Group_row_Index As Double
    Group_row_Index = 2

Dim Close_Open_Diff As Double
    Close_Open_Diff = 0
    
Dim Yearly_Change As Double
    Yearly_Change = 0

Dim Perc_change As Double
    Perc_change = 0


Dim Open_Price As Double
    Open_Price = 0

Dim Close_Price As Double
    Close_Price = 0

Dim Volume As Double
    Volume = 0

Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2


' Defining Last Row
'----------------------------
lastrow = Cells(Rows.Count, 1).End(xlUp).Row


' Defining iterable
'----------------------------
For i = 2 To lastrow


' Defining conditional loop
'----------------------------
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then ' If ticker row doesnt match next ticker row execute the following
                
        Ticker = Cells(i, 1).Value  ' Iterate Ticker regardless of whether If is true or false
        Close_Price = Cells(i, 6).Value ' Define Close_price as the last i row value when the if-mismatch occurs
        Volume = Volume + Cells(i, 7).Value ' Add the next row to Volume regardless of whether If is true or false
        Range("I" & Summary_Table_Row).Value = Ticker ' Deliver the i row value to the table from ticker when the if-mismatch happens
               
        Close_Open_Diff = (Close_Price - Cells(Group_row_Index, 3).Value) ' Define Close_Open_Diff
            
            If (Cells(Group_row_Index, 3) <> 0) Then ' Nested If: determine if Group_row_index reset to 0 and if not redefine Perc_Change
                    Perc_change = (Close_Open_Diff / Cells(Group_row_Index, 3)) 'Define Perc_Change
            Else
                Perc_change = 0 ' Reset Perc_Change to 0
            End If
     
        Range("j" & Summary_Table_Row).Value = Close_Open_Diff ' Deliver Close_Open_Diff
        Range("K" & Summary_Table_Row).Value = Perc_change ' Deliver Perc_Change

    ' Applying Formatting
    '----------------------------
        If Perc_change < 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3 'Apply Red Color
            Range("J" & Summary_Table_Row).NumberFormat = "0.00" 'Format = % with double
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 3 'Apply Red Color
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%" 'Format = % with double
            
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4 'Apply Green Color
            Range("J" & Summary_Table_Row).NumberFormat = "0.00" 'Format = % with double
            Range("K" & Summary_Table_Row).Interior.ColorIndex = 4 'Apply Green Color
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%" 'Format = % with double
        
        End If
               
        Range("L" & Summary_Table_Row).Value = Volume ' Deliver Volume
        Summary_Table_Row = Summary_Table_Row + 1 ' Add 1 row to Summary_Table_Row for next iteration
        Ticker = 0 ' After delivering the i value to the I column (above), reset Ticker to 0
        Close_Price = 0 ' Reset Close_Price
        Volume = 0 ' Reset Volume
        Group_row_Index = i + 1 ' Group_row_index starts as 2 but adds 1 with each iteration
                    
    Else
    
        Ticker = Ticker + Cells(i, 1).Value ' Iterate Ticker regardless of whether If is true or false
        Volume = Volume + Cells(i, 7).Value ' Add the next row to Volume regardless of whether If is true or false
                        
    End If
           
  Next i

End Sub