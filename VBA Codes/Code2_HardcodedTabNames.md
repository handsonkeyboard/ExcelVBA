
# VBA Code #2: Hardcoded Tab Names (Array Approach)

## 🎯 Code Overview
This VBC macro generate specific worksheets using their hardcoded names. 

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
' VBA Code #2: Hardcoded Tab Names (Array Approach)
' i.e., generate tabs based on a list of names
' ---------------------------------------------------------------------------------- '

' Sub: the keyword for a subroutine, which is a block of code that performs a specific task
Sub CreateTabs()

    ' Dim: short for Dimension, used to declare variables
    ' declare a variable ws to represent a worksheet
    Dim ws As Worksheet
    
    ' declare tab_names as Variant
    ' Variant can store multiple types of data including arrays
    Dim tab_names As Variant
    
    ' declare i as an integer, serving as the loop count
    Dim i As Integer
    
    ' define the names of the tabs
    tab_names = Array("Carbon Emissions", "Renewable Energy", "Energy Efficiency", "Climate Policies")
    
    'loop through the array and create tabs
    'LBound & UBound: functions that return the lower and upper bounds of the array
    For i = LBound(tab_names) To UBound(tab_names)
    
        ' prevent the macro from stopping if an error occurs
        On Error Resume Next ' Ignore errors for duplicate names
        
        ' add a new worksheet to the active workbook
        Set ws = ThisWorkbook.Worksheets.Add
        
        ' set the name of the newly created worksheet to the current array element
        ws.Name = tab_names(i)
        
        ' reset error handling to default - VBA will stop the macro if any error occurs
        On Error GoTo 0 'Reset error handling
        
    Next i

    ' MsgBox: a built-in VBA function that displays a message box on the screen
    MsgBox "Worksheets created successfully!"
        
' End Sub: the end of the subroutine
End Sub

```

  ## 📜 License
  
  This code is shared under the [MIT](https://choosealicense.com/licenses/mit/) License. You are free to use, modify, and distribute it, but attribution is appreciated. 

  ## 😊 Authors

- [@handsonkeyboard](https://www.github.com/handsonkeyboard)
    
