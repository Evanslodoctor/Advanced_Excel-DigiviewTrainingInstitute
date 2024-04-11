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
1. ***nper:*** The total number of payment periods.

2. ***pmt:*** The payment made each period. Must remain constant throughout.

3. ***pv:*** The present value or principal amount of the investment or loan.

4. ***fv(optional):*** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. ***type(optional):*** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

6. ***guess(optional):*** An initial guess at the interest rate. If omitted, assumed to be 10%.

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

1. ***rate:*** The interest rate per period.

2. ***pmt:*** The payment made each period. Must remain constant throughout.

3. ***pv:*** The present value or principal amount of the loan or investment.

4. ***fv(optional):*** The future value or final amount that a series of payments will grow to. If omitted, assumed to be 0.

5. ***type (optional):*** The timing of the payment. 0 for payments at the end of the period (default), 1 for payments at the beginning of the period.

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
