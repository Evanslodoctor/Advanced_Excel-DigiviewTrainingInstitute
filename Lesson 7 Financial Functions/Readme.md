# Financial Functions and formulars

- Financial formulas and functions in Excel are vital tools for performing various financial calculations, such as loan payments, interest rates, depreciation, and more. These functions help users analyze and manage financial data efficiently. Below are some commonly used financial formulas and functions in Excel:

# Sample Financial Data Table

| Period | Investment | Payment | Interest Rate |
| ------ | ---------- | ------- | ------------- |
| 0      | $10,000    | $0      | 5%            |
| 1      |            | $200    |               |
| 2      |            | $200    |               |
| 3      |            | $200    |               |
| 4      |            | $200    |               |
| 5      |            | $200    |               |

# 1. FV (Future Value)

## Definition:

- The FV function in Excel calculates the future value of an investment based on a series of periodic payments and a constant interest rate. It's commonly used in financial analysis to determine the value of an investment or savings account at a future date.

## Syntax:

```excel
=FV(rate, nper, pmt, [pv], [type])
```

1. **_rate:_** The interest rate per period.

2. **_nper:_** The total number of payment periods.

3. **_pmt:_** The payment made each period. Must remain constant throughout.

4. **_pv(optional):_** The present value or principal amount. If omitted, assumed to be 0.

5. **_type(optional):_** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

## Explanation:

- The FV function calculates the future value of an investment or loan based on the provided interest rate, the number of payment periods, and the amount of each payment. It assumes that payments are made at regular intervals and remain constant over the entire period. The future value represents the total value of the investment or loan at the end of the specified period, including both principal and interest.

## Example:

- Suppose you invest $10,000 today in a savings account with an annual interest rate of 5%. You plan to make monthly payments of $200 for the next 5 years. You want to know the future value of your investment at the end of 5 years.

## Using the FV function:

```excel

=FV(5%/12, 5*12, -200, -10000)
```

- This formula calculates the future value of the investment, given a monthly interest rate of 5%/12 (since it's compounded monthly), a total of 5*12 = 60 payment periods (5 years * 12 months per year), a monthly payment of -$200 (negative because it's an outgoing payment), and an initial investment of -$10,000 (also negative because it's an outgoing payment).

- The result will give you the future value of the investment at the end of the 5-year period.

## Output:

- The output of the FV function will be the future value of the investment, which represents the total amount you'll have at the end of the specified period, including both the initial investment and the accumulated interest.

# PV (Present Value)

## Definition:

- The PV function in Excel calculates the present value of an investment or loan based on a series of future payments and a constant interest rate. It's commonly used in financial analysis to determine the current value of an investment or loan, considering the time value of money.

## Syntax:

```excel
=PV(rate, nper, pmt, [fv], [type])
```

1. **_rate:_** The interest rate per period.

2. **_nper:_** The total number of payment periods.

3. **_pmt:_** The payment made each period. Must remain constant throughout.

4. **_fv(optional):_** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. **_type(optional):_** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

## Explanation:

- The PV function calculates the present value of an investment or loan based on the provided interest rate, the number of payment periods, and the amount of each payment. It assumes that payments are made at regular intervals and remain constant over the entire period. The present value represents the current worth of the investment or loan, considering the time value of money.

## Example:

- Suppose you are considering an investment opportunity that promises to pay you $200 every month for the next 5 years. The interest rate is 6% per annum. You want to know the present value of this investment opportunity.

## Using the PV function:

```excel

=PV(6%/12, 5*12, -200)
```

- This formula calculates the present value of the investment, given a monthly interest rate of 6%/12 (since it's compounded monthly), a total of 5*12 = 60 payment periods (5 years * 12 months per year), and a monthly payment of -$200 (negative because it's an outgoing payment).

- The result will give you the present value of the investment opportunity.

## Output:

- The output of the PV function will be the present value of the investment or loan, representing the current worth of the cash flows, considering the time value of money.

# PMT (Payment)

## Definition:

- The PMT function in Excel calculates the periodic payment for a loan or investment based on constant payments and a constant interest rate. It's commonly used in financial analysis to determine the fixed payment required to pay off a loan or investment over a specified period.

## Syntax:

```excel
=PMT(rate, nper, pv, [fv], [type])
```

1. **_rate:_** The interest rate per period.

2. **_nper:_** The total number of payment periods.

3. **_pv:_** The present value or principal amount of the loan or investment.

4. **_fv(optional):_** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. **_type(optional):_** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

## Explanation:

- The PMT function calculates the periodic payment for a loan or investment based on the provided interest rate, the number of payment periods, and the present value of the loan or investment. It assumes that payments are made at regular intervals and remain constant over the entire period. The payment represents the fixed amount required to pay off the loan or investment over the specified period.

## Example:

- Suppose you take out a loan of $10,000 with an annual interest rate of 5%. The loan term is 5 years. You want to know the monthly payment required to pay off the loan.

## Using the PMT function:

```excel

=PMT(5%/12, 5*12, 10000)
```

- This formula calculates the monthly payment for the loan, given a monthly interest rate of 5%/12 (since it's compounded monthly), a total of 5*12 = 60 payment periods (5 years * 12 months per year), and a present value of -$10,000 (negative because it's an outgoing payment).

- The result will give you the fixed monthly payment required to pay off the loan over the 5-year period.

## Output:

- The output of the PMT function will be the periodic payment required to pay off the loan or investment, representing the fixed amount to be paid at regular intervals over the specified period.

# Interest Rate (RATE)

## Definition:

- The RATE function in Excel calculates the interest rate per period for an investment or loan based on periodic, constant payments and a constant present value (principal). It's commonly used in financial analysis to determine the interest rate required to reach a desired future value or to pay off a loan over a specified period.

## Syntax:

```excel
=RATE(nper, pmt, pv, [fv], [type], [guess])
```

1. **_nper:_** The total number of payment periods.

2. **_pmt:_** The payment made each period. Must remain constant throughout.

3. **_pv:_** The present value or principal amount of the investment or loan.

4. **_fv(optional):_** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. **_type(optional):_** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

6. **_guess(optional):_** An initial guess at the interest rate. If omitted, assumed to be 10%.

## Explanation:

-The RATE function calculates the interest rate per period required for an investment or loan based on the provided number of payment periods, the amount of each payment, the present value of the investment or loan, and the future value (if applicable). It assumes that payments are made at regular intervals and remain constant over the entire period. The interest rate represents the rate at which the investment grows or the loan is paid off.

## Example:

-Suppose you invest $10,000 today in a savings account and plan to make monthly deposits of $200 for the next 5 years. You want to know the annual interest rate required to reach a future value of $20,000.
Using the RATE function:

```excel
=RATE(5*12, -200, -10000, 20000)
```

- This formula calculates the annual interest rate required for the investment, given a total of 5*12 = 60 payment periods (5 years * 12 months per year), a monthly deposit of -$200 (negative because it's an outgoing payment), an initial investment of -$10,000 (also negative because it's an outgoing payment), and a future value of $20,000.

- The result will give you the annual interest rate required to reach the desired future value.

## Output:

- The output of the RATE function will be the interest rate per period required for the investment or loan, representing the rate at which the investment grows or the loan is paid off.

# NPER (Number of Periods)

## Definition:

- The NPER function in Excel calculates the total number of payment periods required to pay off a loan or reach a financial goal based on periodic, constant payments and a constant interest rate. It's commonly used in financial analysis to determine the time required to pay off a loan or reach a savings goal.

## Syntax:

```excel
=NPER(rate, pmt, pv, [fv], [type])
```

1. **_rate:_** The interest rate per period.

2. **_pmt:_** The payment made each period. Must remain constant throughout.

3. **_pv:_** The present value or principal amount of the loan or investment.

4. **_fv(optional):_** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. **_type (optional):_** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

## Explanation:

- The NPER function calculates the total number of payment periods required to pay off a loan or reach a financial goal based on the provided interest rate, the amount of each payment, the present value of the loan or investment, and the future value (if applicable). It assumes that payments are made at regular intervals and remain constant over the entire period. The number of periods represents the time required to achieve the financial goal or pay off the loan.

## Example:

- Suppose you take out a loan of $10,000 with an annual interest rate of 5%. You plan to make monthly payments of $200. You want to know how long it will take to pay off the loan.

## Using the NPER function:

```excel
=NPER(5%/12, -200, -10000)
```

- This formula calculates the total number of payment periods required to pay off the loan, given a monthly interest rate of 5%/12 (since it's compounded monthly), a monthly payment of -$200 (negative because it's an outgoing payment), and an initial loan amount of -$10,000 (also negative because it's an outgoing payment).

- The result will give you the total number of months required to pay off the loan.

## Output:

- The output of the NPER function will be the total number of payment periods required to pay off the loan or reach the financial goal, representing the time required to achieve the goal or pay off the loan.

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


# Advanced:

# Cumulative Principal Paid on a Loan Between Two Periods (CUMPRINC)

## Definition:

- Cumulative principal paid on a loan between two periods represents the total amount of the loan principal that has been repaid from the beginning of the loan term to a specific period. It's essential for borrowers to track the cumulative principal paid to understand how much of the original loan amount has been repaid over time.

## Formula:

To calculate the cumulative principal paid on a loan between two periods, you can use the following formula:

```excel
Cumulative Principal = Total Payments - Cumulative Interest
```
Where:

Total Payments = Total amount paid towards the loan including both principal and interest.

Cumulative Interest = Total interest paid on the loan between the beginning of the loan term and the specified period.
# Example:
Suppose you take out a loan of $10,000 at an annual interest rate of 6%. The loan term is 5 years, and you make monthly payments. You want to find out the cumulative principal paid on the loan after the first 3 years.

## Using the Formula:
Calculate the total payments made towards the loan for the first 3 years.
Determine the cumulative interest paid on the loan for the first 3 years.
Subtract the cumulative interest from the total payments to find the cumulative principal paid.
## Output:
The output of the calculation will be the cumulative principal paid on the loan between the beginning of the loan term and the specified period.

# Double Declining Balance Depreciation

## Definition:

- Double Declining Balance (DDB) Depreciation is an accelerated depreciation method used to calculate the depreciation expense of an asset. It assumes that the asset loses value more rapidly in the early years of its useful life and decreases gradually over time. The DDB method allows for higher depreciation expenses in the early years, which can be beneficial for tax purposes or reflecting the asset's actual usage pattern.

## Formula:

To calculate depreciation using the Double Declining Balance method, you can use the following formula:

```excel
Depreciation Expense = (2 * Straight-line depreciation rate) * Book value at the beginning of the period
```
Where:

Straight-line depreciation rate = 1 / Useful life of the asset (in periods).

Book value at the beginning of the period = Initial cost of the asset - Accumulated depreciation.

# Example:
- Suppose you purchase a piece of equipment for $10,000 with an expected useful life of 5 years and no salvage value. To calculate the depreciation expense using the Double Declining Balance method for the first year:

# Calculate the straight-line depreciation rate:
Straight-line depreciation rate = 1 / 5 years = 0.2 or 20%

Determine the book value at the beginning of the first year:
Book value = Initial cost - Accumulated depreciation = $10,000 - $0 = $10,000

Calculate the depreciation expense for the first year:
Depreciation Expense = (2 * 20%) * $10,000 = $4,000

# Output:
The output of the calculation will be the depreciation expense for the specified period using the Double Declining Balance method.


# SLN (Straight-Line Depreciation)

## Definition:

- SLN (Straight-Line Depreciation) is a method used to calculate the depreciation expense of an asset uniformly over its useful life. It assumes that the asset loses value at a constant rate each period. This method is straightforward and widely used for financial reporting purposes.

## Formula:

To calculate depreciation using the Straight-Line Depreciation method, you can use the following formula:

```excel
Depreciation Expense = (Cost of the asset - Salvage value) / Useful life of the asset

Where:

Cost of the asset: Initial cost of acquiring the asset
Salvage value: Estimated value of the asset at the end of its useful life

Useful life of the asset: Total number of periods over which the asset is expected to be used
```
## Example:
- Suppose you purchase a piece of machinery for $50,000 with a salvage value of $5,000 and an expected useful life of 5 years. To calculate the depreciation expense using the Straight-Line Depreciation method for each year:

- Determine the depreciable cost:
Depreciable Cost = Cost of the asset - Salvage value = $50,000 - $5,000 = $45,000

- Calculate the annual depreciation expense:
Depreciation Expense = Depreciable Cost / Useful life of the asset = $45,000 / 5 years = $9,000 per year

## Output:
The output of the calculation will be the depreciation expense for each period using the Straight-Line Depreciation method.

SYD (Sum of Years' Digits Depreciation)

# SYD (Sum of Years' Digits Depreciation)

## Definition:

- SYD (Sum of Years' Digits Depreciation) is a method used to calculate the depreciation expense of an asset based on its useful life. This method assumes that assets lose value more rapidly in the earlier years of their useful life and slows down as they approach the end of their useful life. SYD depreciation is calculated using a formula that takes into account the total number of years of an asset's useful life.

## Formula:

To calculate depreciation using the Sum of Years' Digits Depreciation method, you can use the following formula:

```excel
Depreciation Expense = (Remaining useful life / Sum of years' digits) * (Cost of the asset - Salvage value)
Where:

Remaining useful life: Number of years left in the asset's useful life at the beginning of the period
Sum of years' digits: The sum of the digits representing the years of the asset's useful life (e.g., for a 5-year asset, the sum of years' digits would be 1 + 2 + 3 + 4 + 5 = 15)
Cost of the asset: Initial cost of acquiring the asset
Salvage value: Estimated value of the asset at the end of its useful life
Example:
Suppose you purchase a piece of machinery for $50,000 with a salvage value of $5,000 and an expected useful life of 5 years. To calculate the depreciation expense using the Sum of Years' Digits Depreciation method for each year:

Determine the sum of years' digits:
Sum of Years' Digits = 1 + 2 + 3 + 4 + 5 = 15

Calculate the depreciation expense for each year:

Year 1: (5 / 15) * ($50,000 - $5,000) = $13,333.33
Year 2: (4 / 15) * ($50,000 - $5,000) = $10,666.67
Year 3: (3 / 15) * ($50,000 - $5,000) = $8,000
Year 4: (2 / 15) * ($50,000 - $5,000) = $5,333.33
Year 5: (1 / 15) * ($50,000 - $5,000) = $2,666.67
```
## Output:
The output of the calculation will be the depreciation expense for each period using the Sum of Years' Digits Depreciation method.

# DB (Fixed Declining Balance Depreciation)

## Definition:

- DB (Fixed Declining Balance Depreciation) is a method used to calculate the depreciation expense of an asset based on a fixed rate of depreciation applied to the asset's book value. This method assumes that assets lose value more rapidly in the earlier years of their useful life and slows down as they approach the end of their useful life. DB depreciation is calculated using a fixed depreciation rate applied to the remaining book value of the asset each period.

## Formula:

To calculate depreciation using the Fixed Declining Balance Depreciation method, you can use the following formula:

```excel
Depreciation Expense = Depreciation Rate * Book Value at Beginning of Period
Where:

Depreciation Rate: Fixed rate of depreciation applied to the asset's book value (expressed as a percentage)
Book Value at Beginning of Period: The remaining book value of the asset at the beginning of the depreciation period
Example:
Suppose you purchase a piece of machinery for $50,000 with a salvage value of $5,000 and an expected useful life of 5 years. You decide to use the Fixed Declining Balance Depreciation method with a depreciation rate of 20% per year. To calculate the depreciation expense for each year:

Calculate the depreciation rate:
Depreciation Rate = 20% (given)

Determine the book value at the beginning of each period:

Year 1: $50,000
Year 2: Book Value at End of Year 1 - Depreciation Expense for Year 1
Year 3: Book Value at End of Year 2 - Depreciation Expense for Year 2
Year 4: Book Value at End of Year 3 - Depreciation Expense for Year 3
Year 5: Book Value at End of Year 4 - Depreciation Expense for Year 4
Calculate the depreciation expense for each year:

Year 1: 20% * $50,000 = $10,000
Year 2: 20% * (Book Value at End of Year 1) = $8,000
Year 3: 20% * (Book Value at End of Year 2) = $6,400
Year 4: 20% * (Book Value at End of Year 3) = $5,120
Year 5: 20% * (Book Value at End of Year 4) = $4,096
```
## Output:
The output of the calculation will be the depreciation expense for each period using the Fixed Declining Balance Depreciation method.