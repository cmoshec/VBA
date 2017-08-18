Sub AddUDFToCustomCategory()
' Corresponds to options in the Macro Options dialog box. You can also use this method to display a user
' defined function (UDF) in a built-in or new category within the Insert Function dialog box

   Application.MacroOptions Macro:="ABC", Description:="Multiply two numbers", Category:="My Custom Category"
  
    
End Sub

Public Function ABC(num1, num2) ' example to a simple function - multiply two numbers
ABC = num1 * num2
End Function



