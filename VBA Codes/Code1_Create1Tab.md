# VBA Code #1: Add a New Worksheet

## 🎯 Code Overview
This VBC macro generate a new worksheet. 

## 📗 Usage Guide
Copy and paste the following code into your VBA editor:
1. Press Alt + F11 to open the VBA editor.
2. Inside the VBA editor, select "Insert > Module" to create a new module.
3. Change the name of the model in the "Properties" pane.
4. Copy the VBA code from below and paste it into the module. 
5. Run the macro. 


## 📑 VBA Code
```

Sub Create1Tab()

    Dim ws As Worksheet
    
    ' add a new worksheet to the workbook
    Set ws = ThisWorkbook.Sheets.Add
        ws.Name = "NewSheet"
    

End Sub


```

  ## 📜 License
  
  This code is shared under the [MIT](https://choosealicense.com/licenses/mit/) License. You are free to use, modify, and distribute it, but attribution is appreciated. 

  ## 😊 Authors

- [@handsonkeyboard](https://www.github.com/handsonkeyboard)
