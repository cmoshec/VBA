Attribute VB_Name = "Module1"

Public Function ABC(num1, num2)
ABC = num1 * num2
End Function

Public Function ABC(num1, num2)
ABC = num1 * num2
End Function


Sub AddUDFToCustomCategory()
' Corresponds to options in the Macro Options dialog box. You can also use this method to display a user
' defined function (UDF) in a built-in or new category within the Insert Function dialog box

   Application.MacroOptions Macro:="ABC", Description:="Multiply two numbers", Category:="My Custom Category"
  
    
End Sub


