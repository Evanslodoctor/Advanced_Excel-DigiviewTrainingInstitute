# Introduction to VBA

## What is VBA?

- **VBA (Visual Basic for Applications)** is a programming language developed by Microsoft.
- It allows users to automate tasks and create custom functions within various Microsoft applications, including Excel, Word, Access, and PowerPoint.
- VBA is specifically powerful in Excel, where it can be used to manipulate data, automate repetitive tasks, and create interactive dashboards and forms.

## Advantages of Using VBA in Excel

- **Automation**: VBA enables users to automate repetitive tasks, saving time and reducing errors.
- **Customization**: Users can create custom solutions tailored to their specific needs and workflows.
- **Integration**: VBA seamlessly integrates with Excel, allowing users to extend Excel's functionality beyond its built-in features.
- **Efficiency**: By leveraging VBA, users can perform complex operations on large datasets more efficiently than manual methods.
- **Flexibility**: VBA provides a wide range of functionalities and control over Excel objects, enabling users to accomplish diverse tasks.

## Getting Familiar with the VBA Environment

- **Accessing the VBA Editor**: Press `Alt + F11` within Excel to open the Visual Basic for Applications (VBA) editor.
- **Project Explorer**: Displays all open workbooks and their corresponding VBA projects. It provides a hierarchical view of the objects within each workbook.
- **Code Window**: Where VBA code is written and edited. Each workbook can have multiple code modules.
- **Immediate Window**: Allows for direct execution of VBA statements and functions, helpful for testing code snippets.
- **Object Browser**: Provides a searchable list of all available objects, properties, methods, and constants in VBA.
- **Properties Window**: Displays properties of selected objects for easy reference and modification.
- **Toolbar and Menu**: Offers various tools and options for writing, debugging, and managing VBA code.

## Conclusion

- Understanding the fundamentals of VBA and its advantages in Excel is essential for leveraging its power to automate tasks and enhance productivity.
- In the subsequent modules, we will delve deeper into VBA syntax, procedures, Excel objects, and advanced techniques to empower you to become proficient in VBA programming.

# VBA Basics

In this chapter, you will acquaint yourself with the commonly used Excel VBA terminologies. These terminologies will be used in further modules, hence understanding each one of these is important.

## Modules

Modules is the area where the code is written. This is a new Workbook, hence there aren't any Modules.

### Module in VBScript

To insert a Module, navigate to `Insert → Module`. Once a module is inserted 'module1' is created.

Within the modules, we can write VBA code and the code is written within a Procedure. A Procedure/Sub Procedure is a series of VBA statements instructing what to do.

### Procedure

Procedures are a group of statements executed as a whole, which instructs Excel how to perform a specific task. The task performed can be a very simple or a very complicated task. However, it is a good practice to break down complicated procedures into smaller ones.

The two main types of Procedures are Sub and Function.

### Function

A function is a group of reusable code, which can be called anywhere in your program. This eliminates the need of writing the same code over and over again. This helps the programmers to divide a big program into a number of small and manageable functions.

Apart from inbuilt Functions, VBA allows writing user-defined functions as well, and statements are written between `Function` and `End Function`.

### Sub-Procedures

Sub-procedures work similar to functions. While sub-procedures DO NOT return a value, functions may or may not return a value. Sub-procedures CAN be called without the `Call` keyword. Sub-procedures are always enclosed within `Sub` and `End Sub` statements.

# Comments in VBA

Comments are used to document the program logic and the user information with which other programmers can seamlessly work on the same code in the future.

It includes information such as developed by, modified by, and can also include incorporated logic. Comments are ignored by the interpreter while execution.

Comments in VBA are denoted by two methods.

1. Any statement that starts with a Single Quote (`'`) is treated as a comment. Following is an example.

   ```vba
   ' This Script is invoked after successful login
   ' Written by: TutorialsPoint
   ' Return Value: True / False
   ```

```
Any statement that starts with the keyword REM. Following is an example.
```
REM This Script is written to Validate the Entered Input 
REM Modified by: Tutorials point/user2
```

