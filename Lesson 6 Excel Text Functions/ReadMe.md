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


# Intermediate:

# NPV (Net Present Value)

## Definition:

- The NPV function in Excel calculates the net present value of an investment by discounting the cash flows at a specified rate. It's commonly used in financial analysis to determine the profitability of an investment by comparing the present value of expected cash inflows with the present value of cash outflows.

## Syntax:

```excel
=NPV(rate, value1, [value2], ...)
```

1. ***rate:*** The discount rate per period.

2. value1, value2, ...: The series of cash flows representing income and expenses. These values must be entered as a list of cash flows, separated by commas.

## Explanation:
- The NPV function calculates the net present value of an investment by discounting the future cash flows back to their present value using a specified discount rate. The present value of cash inflows is subtracted from the present value of cash outflows to determine the net present value. A positive NPV indicates that the investment is profitable, while a negative NPV indicates that it is not.
## Example:
- Suppose you are considering an investment that will generate cash flows of $1,000 in year 1, $1,500 in year 2, and $2,000 in year 3. The discount rate for the investment is 10%. You want to calculate the net present value of the investment.
## Using the NPV function:
```excel

=NPV(10%, 1000, 1500, 2000)
```
- This formula calculates the net present value of the investment, given a discount rate of 10% and cash flows of $1,000, $1,500, and $2,000 in years 1, 2, and 3 respectively.

- The result will give you the net present value of the investment.

## Output:
The output of the NPV function will be the net present value of the investment, representing the difference between the present value of cash inflows and the present value of cash outflows.

# IRR (Internal Rate of Return)

## Definition:

- The IRR function in Excel calculates the internal rate of return for a series of cash flows. It represents the discount rate that makes the net present value of the cash flows equal to zero. IRR is commonly used in financial analysis to evaluate the profitability of an investment or project.

## Syntax:

```excel
=IRR(values, [guess])
```
1. ***values:*** The series of cash flows representing income and expenses. These values must be entered as a list of cash flows, separated by commas.

2. ***guess(optional):*** An initial guess for the internal rate of return. If omitted, Excel uses 0.1 (10%) as the default guess.

# Explanation:
The IRR function calculates the internal rate of return by finding the discount rate that results in a net present value of zero for the series of cash flows. It uses an iterative approach to approximate the rate. The internal rate of return represents the effective annual return on investment and is used to assess the profitability of projects or investments.
# Example:
Suppose you are evaluating an investment project that requires an initial outlay of $10,000 and generates cash inflows of $3,000, $4,000, $5,000, and $6,000 over the next four years. You want to calculate the internal rate of return for the project.
Using the IRR function:
```excel

=IRR(-10000, 3000, 4000, 5000, 6000)
```
- This formula calculates the internal rate of return for the investment project, given the initial outlay of -$10,000 (negative because it's an outgoing payment) and the subsequent cash inflows of $3,000, $4,000, $5,000, and $6,000.

- The result will give you the internal rate of return for the investment project.

# Output:
- The output of the IRR function will be the internal rate of return, representing the effective annual return on investment for the project.

# Discounted Net Present Value for a Non-Periodic Series of Cash Flows

## Definition:

- The Discounted Net Present Value (NPV) for a non-periodic series of cash flows is a financial metric used to assess the profitability of an investment or project. It represents the present value of all future cash inflows and outflows, discounted at a specified rate of return.

## Formula:

The formula for calculating the Discounted Net Present Value (NPV) for a non-periodic series of cash flows is as follows:

```excel
NPV = CF1 / (1 + r)^1 + CF2 / (1 + r)^2 + ... + CFn / (1 + r)^n
```
Where:

1. ***NPV*** = Net Present Value.

2. CF1, CF2, ..., CFn = Cash flows for each period.

3. ***r*** = Discount rate or required rate of return.

4. ***n*** = Number of cash flows.

#Example:
- Suppose you are considering an investment project that requires an initial investment of $10,000. Over the next four years, the project generates cash inflows of $3,000, $4,000, $5,000, and $6,000, respectively. You want to assess the net present value of the project using a discount rate of 10%.

# Using the Formula:
``` excel

NPV = -10000 + 3000 / (1 + 0.10)^1 + 4000 / (1 + 0.10)^2 + 5000 / (1 + 0.10)^3 + 6000 / (1 + 0.10)^4
```
Substitute the cash flows and discount rate into the formula and calculate the NPV.
## Output:
- The output of the calculation will be the net present value of the investment project. A positive NPV indicates that the project is expected to generate returns higher than the required rate of return, while a negative NPV suggests the project may not be viable.





# Internal Rate of Return for a Non-Periodic Series of Cash Flows

## Definition:

- The Internal Rate of Return (IRR) for a non-periodic series of cash flows is a financial metric used to assess the profitability of an investment or project. It represents the discount rate that makes the net present value (NPV) of the cash flows equal to zero.

## Formula:

The IRR calculation for a non-periodic series of cash flows involves finding the discount rate (r) that satisfies the equation:

```excel
NPV = CF1 / (1 + r)^1 + CF2 / (1 + r)^2 + ... + CFn / (1 + r)^n = 0
```
Where:

NPV = Net Present Value.

CF1, CF2, ..., CFn = Cash flows for each period.

r = Internal Rate of Return.

n = Number of cash flows.

## Example:
- Suppose you are evaluating an investment opportunity that requires an initial investment of $10,000. Over the next three years, the project generates cash inflows of $4,000, $5,000, and $6,000, respectively. You want to calculate the internal rate of return for this investment.

## Using the Formula:
To calculate the internal rate of return, you need to find the discount rate (r) that makes the NPV of the cash flows equal to zero. This can be done using iterative methods or built-in functions in spreadsheet software like Excel.

## Output:
The output of the calculation will be the internal rate of return (IRR) for the investment project. A higher IRR indicates a more profitable investment, as it represents the discount rate at which the project breaks even.

# Cumulative Interest Paid on a Loan Between Two Periods

## Definition:

- Cumulative interest paid on a loan between two periods represents the total amount of interest accrued on a loan from the beginning of the loan term to a specific period. It's essential for borrowers to understand the cumulative interest paid to assess the total cost of borrowing and plan their finances effectively.

## Formula:

To calculate the cumulative interest paid on a loan between two periods, you can use the following formula:

```excel
Cumulative Interest = Total Payments - Loan Principal
```
Where:

Total Payments = Total amount paid towards the loan including both principal and interest.

Loan Principal = Original amount borrowed.

# Example:
Suppose you take out a loan of $10,000 at an annual interest rate of 6%. The loan term is 5 years, and you make monthly payments. You want to find out the cumulative interest paid on the loan after the first 3 years.

# Using the Formula:
Calculate the total payments made towards the loan for the first 3 years.
Subtract the original loan principal from the total payments to find the cumulative interest paid.
# Output:
The output of the calculation will be the cumulative interest paid on the loan between the beginning of the loan term and the specified period.