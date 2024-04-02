# VBA - Message Box

The `MsgBox` function displays a message box and waits for the user to click a button, and then an action is performed based on the button clicked by the user.

## Syntax

```vba
MsgBox(prompt[,buttons][,title][,helpfile,context])
```

## Parameter Description

- **Prompt**: A Required Parameter. A String that is displayed as a message in the dialog box. The maximum length of the prompt is approximately 1024 characters. If the message extends to more than a line, then the lines can be separated using a carriage return character (`Chr(13)`) or a linefeed character (`Chr(10)`) between each line.

- **Buttons**: An Optional Parameter. A Numeric expression that specifies the type of buttons to display, the icon style to use, the identity of the default button, and the modality of the message box. If left blank, the default value for buttons is 0.

- **Title**: An Optional Parameter. A String expression displayed in the title bar of the dialog box. If the title is left blank, the application name is placed in the title bar.

- **Helpfile**: An Optional Parameter. A String expression that identifies the Help file to use for providing context-sensitive help for the dialog box.

- **Context**: An Optional Parameter. A Numeric expression that identifies the Help context number assigned by the Help author to the appropriate Help topic. If context is provided, helpfile must also be provided.

## Return Values

The `MsgBox` function can return one of the following values which can be used to identify the button the user has clicked in the message box:

- **1**: `vbOK` - OK was clicked
- **2**: `vbCancel` - Cancel was clicked
- **3**: `vbAbort` - Abort was clicked
- **4**: `vbRetry` - Retry was clicked
- **5**: `vbIgnore` - Ignore was clicked
- **6**: `vbYes` - Yes was clicked
- **7**: `vbNo` - No was clicked

## Example

```vba
Function MessageBox_Demo()
   ' Message Box with just prompt message
   MsgBox "Welcome"

   ' Message Box with title, yes no and cancel buttons
   Dim a As Integer
   a = MsgBox("Do you like blue color?", 3, "Choose options")
   ' Assume that you press No Button
   MsgBox "The Value of a is " & a
End Function
```
