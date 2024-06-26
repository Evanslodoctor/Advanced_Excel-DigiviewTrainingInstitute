# Excel Text Functions
## Excel Text Functions List

### Functions to Remove Extra Characters
- **CLEAN**: Removes all non-printable characters from a supplied text string
  - Syntax: `CLEAN(text)`
  - Example: `=CLEAN("abc#1$%^&*")` returns `"abc1"`

- **TRIM**: Removes duplicate spaces, and spaces at the start and end of a text string
  - Syntax: `TRIM(text)`
  - Example: `=TRIM("  hello  ")` returns `"hello"`

### Functions to Convert Between Upper & Lower Case
- **LOWER**: Converts all characters in a supplied text string to lower case
  - Syntax: `LOWER(text)`
  - Example: `=LOWER("Hello World")` returns `"hello world"`

- **PROPER**: Converts all characters in a supplied text string to proper case
  - Syntax: `PROPER(text)`
  - Example: `=PROPER("john DOE")` returns `"John Doe"`

- **UPPER**: Converts all characters in a supplied text string to upper case
  - Syntax: `UPPER(text)`
  - Example: `=UPPER("hello world")` returns `"HELLO WORLD"`

### Functions to Convert Excel Data Types
- **BAHTTEXT**: Converts a number into Thai text
  - Syntax: `BAHTTEXT(number)`
  - Example: `=BAHTTEXT(1234)` returns `"หนึ่งพันสองร้อยสามสิบสี่"`

- **DOLLAR**: Converts a supplied number into text, using a currency format
  - Syntax: `DOLLAR(number, decimals)`
  - Example: `=DOLLAR(1234.5678, 2)` returns `"$1,234.57"`

- **FIXED**: Rounds a supplied number to a specified number of decimal places and converts this into text
  - Syntax: `FIXED(number, decimals)`
  - Example: `=FIXED(1234.5678, 2)` returns `"1234.57"`

- **TEXT**: Converts a supplied value into text, using a user-specified format
  - Syntax: `TEXT(value, format_text)`
  - Example: `=TEXT(NOW(), "dd-mm-yyyy")` returns the current date in the format "dd-mm-yyyy"

- **VALUE**: Converts a text string into a numeric value
  - Syntax: `VALUE(text)`
  - Example: `=VALUE("123")` returns the numeric value `123`

### Converting Between Characters & Numeric Codes
- **CHAR**: Returns the character that corresponds to a supplied numeric value
  - Syntax: `CHAR(number)`
  - Example: `=CHAR(65)` returns `"A"`

- **CODE**: Returns the numeric code for the first character of a supplied string
  - Syntax: `CODE(text)`
  - Example: `=CODE("A")` returns `65`

### Cutting Up & Piecing Together Text Strings
- **CONCAT**: Joins together two or more text strings
  - Syntax: `CONCAT(text1, [text2], ...)`
  - Example: `=CONCAT("Hello", " ", "World")` returns `"Hello World"`

- **LEFT**: Returns a specified number of characters from the start of a supplied text string
  - Syntax: `LEFT(text, num_chars)`
  - Example: `=LEFT("Hello", 2)` returns `"He"`

- **MID**: Returns a specified number of characters from the middle of a supplied text string
  - Syntax: `MID(text, start_num, num_chars)`
  - Example: `=MID("Hello", 2, 3)` returns `"ell"`

- **RIGHT**: Returns a specified number of characters from the end of a supplied text string
  - Syntax: `RIGHT(text, num_chars)`
  - Example: `=RIGHT("Hello", 2)` returns `"lo"`

- **REPT**: Returns a string consisting of a supplied text string, repeated a specified number of times
  - Syntax: `REPT(text, number_times)`
  - Example: `=REPT("abc", 3)` returns `"abcabcabc"`

- **TEXTJOIN**: Joins together two or more text strings, separated by a delimiter
  - Syntax: `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`
  - Example: `=TEXTJOIN(", ", TRUE, "apple", "banana", "orange")` returns `"apple, banana, orange"`

### Information Functions
- **LEN**: Returns the length of a supplied text string
  - Syntax: `LEN(text)`
  - Example: `=LEN("Hello")` returns `5`

- **FIND**: Returns the position of a supplied character or text string from within a supplied text string (case-sensitive)
  - Syntax: `FIND(find_text, within_text, [start_num])`
  - Example: `=FIND("e", "Hello")` returns `2`

- **SEARCH**: Returns the position of a supplied character or text string from within a supplied text string (non-case-sensitive)
  - Syntax: `SEARCH(find_text, within_text, [start_num])`
  - Example: `=SEARCH("e", "Hello")` returns `2`

- **EXACT**: Tests if two supplied text strings are exactly the same
  - Syntax: `EXACT(text1, text2)`
  - Example: `=EXACT("hello", "Hello")` returns `FALSE`

- **T**: Tests whether a supplied value is text
  - Syntax: `T(value)`
  - Example: `=T("hello")` returns `"hello"`

### Replacing / Substituting Parts of a Text String
- **REPLACE**: Replaces all or part of a text string with another string
  - Syntax: `REPLACE(old_text, start_num, num_chars, new_text)`
  - Example: `=REPLACE("Hello", 2, 3, "123")` returns `"H123o"`

- **SUBSTITUTE**: Substitutes all occurrences of a search text string within an original text string with the supplied replacement text
  - Syntax: `SUBSTITUTE(text, old_text, new_text, [instance_num])`
  - Example
# More

## 1. Introduction to Text Functions

Text functions in Excel are used to manipulate and analyze text strings. They are essential for various tasks such as data cleaning, formatting, and extraction.

## 2. CONCATENATE Function

- **Syntax:** `CONCATENATE(text1, [text2], ...)`
- **Explanation:** Combines multiple text strings into one.
- **Example:** `=CONCATENATE("Hello", " ", "World")` outputs "Hello World".


## Overview
The CONCATENATE function in Excel is used to combine multiple text strings or cell values into a single string. It allows users to concatenate (join together) text from different cells or input values, creating a unified string.

## Syntax
```scss

CONCATENATE(text1, [text2], ...)
```
- text1, text2, ...: The text strings or cell references that you want to concatenate. You can specify up to 255 arguments.
## Examples
### Basic Usage:
```excel

=CONCATENATE("Hello", " ", "World")
```
***Output:*** Hello World

### Combining Cell Values:
Assuming cell A1 contains Hello and cell B1 contains World:

```excel

=CONCATENATE(A1, " ", B1)
```
***Output:*** Hello World

## Concatenating Cell Ranges:
Assuming cells A1:A3 contain One, Two, and Three respectively:

```excel

=CONCATENATE(A1:A3)
```
***Output:*** OneTwoThree

## Notes
- The CONCATENATE function can concatenate both text strings and cell references.
- You can include additional text strings or cell references as arguments to concatenate multiple values.
- If any argument is empty or contains a numeric value, it will be treated as text when concatenated.
## Considerations
- CONCATENATE is a legacy function in Excel, and newer versions provide the & operator or the CONCAT function, which offer the same functionality with simpler syntax.
- Using CONCATENATE with large datasets or a high number of arguments may impact performance. In such cases, consider using alternative methods for concatenation.
## Conclusion
The CONCATENATE function is a powerful tool for combining text strings and cell values in Excel. Whether you need to create complex strings from multiple sources or simply join text together, CONCATENATE provides a flexible and efficient solution. Mastering this function can greatly enhance your ability to manipulate and analyze text data in Excel.
## 3. LEFT and RIGHT Functions

- **Syntax:** `LEFT(text, [num_chars])`, `RIGHT(text, [num_chars])`
- **Explanation:** Extracts characters from the left or right side of a text string.
- **Example:** `=LEFT("Excel", 2)` returns "Ex".

## LEFT Function
The LEFT function in Excel extracts a specified number of characters from the left side of a text string.

Syntax:
```excel
LEFT(text, num_chars)
```
text: The text string from which you want to extract characters.
num_chars: The number of characters you want to extract from the left side of the text string.
## Example:
```excel

=LEFT("Hello World", 5)
```
This formula will return "Hello" because it extracts the first 5 characters from the left side of the text string "Hello World".

## RIGHT Function
The RIGHT function in Excel extracts a specified number of characters from the right side of a text string.

## Syntax:
```excel

RIGHT(text, num_chars)
```
text: The text string from which you want to extract characters.
num_chars: The number of characters you want to extract from the right side of the text string.
## Example:
```excel

=RIGHT("Hello World", 5)
```
This formula will return "World" because it extracts the last 5 characters from the right side of the text string "Hello World".

## 4. MID Function

- **Syntax:** `MID(text, start_num, num_chars)`
- **Explanation:** Extracts characters from the middle of a text string.
- **Example:** `=MID("Excel Functions", 7, 9)` returns "Functions".


The MID function in Excel extracts a specific number of characters from a text string, starting at a specified position.

## Syntax:
```excel
MID(text, start_num, num_chars)
```
- text: The text string from which you want to extract characters.
- start_num: The position in the text string from which to start extracting characters. The first character in the text string is at position 1.
- num_chars: The number of characters you want to extract from the text string.
## Example:
```excel

=MID("Hello World", 7, 5)
```
This formula will return "World" because it starts extracting characters from the 7th position in the text string "Hello World" and extracts 5 characters.

## 5. LEN Function

- **Syntax:** `LEN(text)`
- **Explanation:** Returns the number of characters in a text string.
- **Example:** `=LEN("Excel")` returns 5.

## 6. FIND and SEARCH Functions

- **Syntax:** `FIND(find_text, within_text, [start_num])`, `SEARCH(find_text, within_text, [start_num])`
- **Explanation:** Finds the position of a specific text within another text string.
- **Example:** `=FIND("el", "Excel")` returns 2.

## 7. REPLACE Function

- **Syntax:** `REPLACE(old_text, start_num, num_chars, new_text)`
- **Explanation:** Replaces characters within a text string.
- **Example:** `=REPLACE("Excel", 2, 3, "celent")` returns "Excellent".

## 8. SUBSTITUTE Function

- **Syntax:** `SUBSTITUTE(text, old_text, new_text, [instance_num])`
- **Explanation:** Replaces occurrences of a specific text within a text string.
- **Example:** `=SUBSTITUTE("banana", "a", "o")` returns "bonono".

## 9. UPPER, LOWER, and PROPER Functions

- **Syntax:** `UPPER(text)`, `LOWER(text)`, `PROPER(text)`
- **Explanation:** Changes the case of text strings (uppercase, lowercase, proper case).
- **Example:** `=UPPER("excel")` returns "EXCEL".

## 10. TRIM Function

- **Syntax:** `TRIM(text)`
- **Explanation:** Removes leading and trailing spaces from a text string.
- **Example:** `=TRIM("  Excel  ")` returns "Excel".

## 11. TEXTJOIN Function

- **Syntax:** `TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)`
- **Explanation:** Joins text strings with a specified delimiter.
- **Example:** `=TEXTJOIN(", ", TRUE, "apple", "banana", "orange")` returns "apple, banana, orange".

These text functions are powerful tools for manipulating and analyzing text data in Excel.


