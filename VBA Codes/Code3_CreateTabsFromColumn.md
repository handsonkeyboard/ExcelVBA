
# VBA Code #3: Create Worksheet from Column (Dynamic Approach)

## 🎯 Code Overview
This VBC macro generates specific worksheets based on a list of names in a rangee on an existing worksheet. 

## 📗 Usage Guide
Copy and paste the following code into your VBA editor:
1. Press Alt + F11 to open the VBA editor.
2. Inside the VBA editor, select "Insert > Module" to create a new module.
3. Change the name of the model in the "Properties" pane.
4. Copy the VBA code from below and paste it into the module. 
5. Run the macro. 


## 📑 VBA Code
```

' ---------------------------------------------------------------------------------- '
' VBA Code #3: Create Worksheet from Column (Dynamic Approach)
' i.e., generate tabs based on a list of names in a range on an existing worksheet
' ---------------------------------------------------------------------------------- '

Sub CreateTabsFromColumn()

    
    Dim wsSource As Worksheet ' store a reference to the source worksheet containing the list of names
    
    Dim ws As Worksheet ' represent each new worksheet created by the macro
    
    Dim cell As Range ' iterate through each cell in the column containing worksheet names

    Dim startCell As Range ' represent the starting point in the column

    Dim lastRow As Long ' store the last row with data in the column, dynamically determined




    Set wsSource = ThisWorkbook.Sheets("Sheet1")

    ' find the first and last used rows in column A

    Set startCell = wsSource.Range("A1")
    
    
        ' start at the very last cell in the specific column, move upward until it finds the last non-empty cell
        ' return the row number of that cell
        
        ' startCell.Column = 1 in this case (column A)
        ' .End(xlUp): moviing upward to the first non-empty cell or reach the first row - if the column is empty, it stops at the first row e.g., A1
        ' .Row: extract the row number of the cell reached by .End(xlUp)

        lastRow = wsSource.Cells(wsSource.Rows.Count, startCell.Column).End(xlUp).Row
        
        

    ' loop through column A from A1 to the last non-empty cell
    ' specify the range dynamically as A1:A[lastRow]

    For Each cell In wsSource.Range(startCell, wsSource.Cells(lastRow, startCell.Column))
        
        ' skipp empty cells, ensuring only cells with data are processed
        If Not IsEmpty(cell.Value) Then

            On Error Resume Next

            ' check if the sheet name already exists

            If SheetExists(cell.Value) Then

                MsgBox "Worksheet '" & cell.Value & "' already exists."

            Else

                ' add a new sheet and name it
                ' ensure the new sheet is added at the end of the workbook

                Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))


                ' rename the new worksheet using the value from the current cell
                ' if the name is invalid (e.g., contains forbidden characters or is longer than 31 characters), this will trigger an error
                ws.Name = cell.Value

            End If
            
            ' check if an error occurred (e.g., invalid name or duplicate name)
            If Err.Number <> 0 Then

                MsgBox "Error creating sheet: '" & cell.Value & "'."
                ws.Delete
                Err.Clear

            End If

            On Error GoTo 0

        End If
    
    ' move to the next cell in the range and repeats the process
    Next cell

    MsgBox "Worksheets created successfully!"

End Sub


' function to check if a worksheet with a given name already exists

' declare a function named SheetExists taht accepts a string argument and return a boolean value (TRUE or FALSE)
Function SheetExists(sheetName As String) As Boolean
   Dim ws As Worksheet
   
   ' temporarily suppress runtime errors - if the sheet does not exist, VBA will not throw an error, and execution will continue
   On Error Resume Next
   Set ws = ThisWorkbook.Sheets(sheetName)
   
   ' return TRUE if ws is not Nothing (i.e., the sheet exists)
   ' return FALSE if ws is Nothing (i.e., the sheet does not exist)
   SheetExists = Not ws Is Nothing
   
   ' restor normal error handling
   On Error GoTo 0
End Function

```

  ## 📜 License
  
  This code is shared under the [MIT](https://choosealicense.com/licenses/mit/) License. You are free to use, modify, and distribute it, but attribution is appreciated. 

  ## 😊 Authors

- [@handsonkeyboard](https://www.github.com/handsonkeyboard)
    
