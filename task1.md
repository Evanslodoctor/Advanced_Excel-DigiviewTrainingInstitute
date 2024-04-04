## Task 1: Grade Calculator
### Description:
Create a VBA program that calculates and displays the grades for a list of students based on their marks. The program should read the marks of each student from a specified range of cells in an Excel worksheet, calculate their grades according to the following grading scale, and then display the grades in an adjacent column.

***Grading Scale:***
90 or above: A
80 - 89: B
70 - 79: C
60 - 69: D
Below 60: F
***Requirements:***
- The program should prompt the user to select the range of cells containing the students' marks.
- For each student, the program should calculate the grade based on the grading scale provided above.
- The calculated grades should be displayed in an adjacent column next to the marks.
- The program should handle errors gracefully, such as non-numeric input or invalid ranges.
- Provide comments to explain the purpose of each section of your code.
### Example:
- Suppose the student marks are stored in cells A2:A10, and you run the program. After execution, the calculated grades should be displayed in cells B2:B10.

### Additional Challenge:
- Implement error handling to notify the user if any of the input marks are invalid (e.g., negative marks or marks greater than 100).

### Resources:
- VBA programming knowledge.
- Understanding of Excel worksheets and ranges.
- Students can utilize their understanding of variables, loops, conditional statements, and Excel manipulation in VBA to accomplish this task. 
- This task will help reinforce their VBA skills while also practicing problem-solving and logical thinking. Additionally, it provides an opportunity to work with real-world data processing tasks, which can be valuable in various fields.

# Solutions

Here's a basic outline of the VBA code for the Grade Calculator task:

```vb
Sub CalculateGrades()
    Dim marksRange As Range
    Dim cell As Range
    Dim grade As String
    
    ' Prompt the user to select the range of cells containing the students' marks
    On Error Resume Next
    Set marksRange = Application.InputBox("Select the range of cells containing the students' marks:", Type:=8)
    On Error GoTo 0
    
    ' Check if the user canceled the selection
    If marksRange Is Nothing Then
        MsgBox "Operation canceled."
        Exit Sub
    End If
    
    ' Loop through each cell in the selected range
    For Each cell In marksRange
        ' Check if the cell value is numeric
        If IsNumeric(cell.Value) Then
            ' Convert the cell value to a number and calculate the grade
            Select Case cell.Value
                Case Is >= 90
                    grade = "A"
                Case 80 To 89
                    grade = "B"
                Case 70 To 79
                    grade = "C"
                Case 60 To 69
                    grade = "D"
                Case Else
                    grade = "F"
            End Select
            ' Display the calculated grade in the adjacent column
            cell.Offset(0, 1).Value = grade
        Else
            ' Handle non-numeric input
            MsgBox "Invalid input in cell " & cell.Address & ". Please enter a numeric value."
        End If
    Next cell
End Sub
```

This code performs the following tasks:

- Prompts the user to select the range of cells containing the students' marks.
- Checks if the user canceled the selection and exits the subroutine if canceled.
- Loops through each cell in the selected range.
- Checks if the cell value is numeric.
- Calculates the grade based on the grading scale provided.
- Displays the calculated grade in the adjacent column.
- Handles non-numeric input by displaying a message box.
You can further enhance this code by adding error handling for invalid marks, such as negative marks or marks greater than 100, as mentioned in the additional challenge.

## Using if elsif statement

here's the modified code using If statements instead of Select Case:

```vb
Sub CalculateGrades()
    Dim marksRange As Range
    Dim cell As Range
    Dim grade As String
    Dim mark As Integer
    
    ' Prompt the user to select the range of cells containing the students' marks
    On Error Resume Next
    Set marksRange = Application.InputBox("Select the range of cells containing the students' marks:", Type:=8)
    On Error GoTo 0
    
    ' Check if the user canceled the selection
    If marksRange Is Nothing Then
        MsgBox "Operation canceled."
        Exit Sub
    End If
    
    ' Loop through each cell in the selected range
    For Each cell In marksRange
        ' Check if the cell value is numeric
        If IsNumeric(cell.Value) Then
            ' Convert the cell value to an integer
            mark = CInt(cell.Value)
            
            ' Calculate the grade based on the mark
            If mark >= 90 Then
                grade = "A"
            ElseIf mark >= 80 Then
                grade = "B"
            ElseIf mark >= 70 Then
                grade = "C"
            ElseIf mark >= 60 Then
                grade = "D"
            Else
                grade = "F"
            End If
            
            ' Display the calculated grade in the adjacent column
            cell.Offset(0, 1).Value = grade
        Else
            ' Handle non-numeric input
            MsgBox "Invalid input in cell " & cell.Address & ". Please enter a numeric value."
        End If
    Next cell
End Sub
```

This code achieves the same functionality using If statements to determine the grade based on the mark. Each If statement checks a specific range of marks and assigns the corresponding grade. If none of the conditions are met, the grade is set to "F".