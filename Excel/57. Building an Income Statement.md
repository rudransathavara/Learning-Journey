# Building an Income Statement in Excel

## Overview

An **income statement** summarises a company's revenue and expenses over a specific period to determine its profitability.

It is one of the three core financial statements used in financial modelling.

In this exercise, we use a fictional mail-order cookie business to build an income statement in Excel.

The model uses assumptions about customer numbers, average order value, refunds, discounts, product costs and operating expenses to calculate financial results.

## Understanding the Income Statement Structure

The model begins with three main categories:

1. Revenue
2. Cost of Goods Sold (COGS)
3. Operating Expenses

These categories contain assumptions and calculations used to build the income statement.

## 1. Revenue Assumptions

Revenue assumptions determine how much money the business expects to generate.

Examples include:

- New Customers
- Average Order Value
- Refunds as a Percentage of Revenue
- Discounts as a Percentage of Revenue

### New Customers

The number of new customers represents the customers acquired during the relevant period.

### Average Order Value

Average order value represents the average amount spent per order.

### Refunds

Refunds reduce the revenue retained by the business.

### Discounts

Discounts reduce the amount customers pay and therefore reduce the revenue recognised.

## 2. Cost of Goods Sold (COGS)

**Cost of Goods Sold (COGS)** represents the direct costs associated with producing or delivering the goods sold by a business.

For the cookie company, the assumptions may include:

- Materials Costs
- Fulfilment Costs
- Merchant Services Fees

Materials costs may include ingredients and packaging.

Fulfilment costs may include costs associated with preparing and delivering customer orders.

Merchant services fees may include payment processing charges.

## 3. Operating Expenses

Operating expenses are costs associated with running the business.

The model includes assumptions for categories such as:

- Human Resources
- Marketing
- Miscellaneous Expenses
- Depreciation
- Interest Expense

The exact classification of an expense depends on the accounting presentation used by the business.

Interest expense is generally presented separately from operating expenses in a conventional income statement. Depreciation may be included in COGS or operating expenses, depending on the assets involved and the company's accounting policy.

## 4. Tax Rate

The tax rate is used to estimate the tax expense applied to the relevant measure of profit.

For a simplified model, the tax expense may be estimated using an assumed tax rate.

In a complete financial model, tax calculations may require adjustments for tax rules, losses, allowances and other relevant factors.

## Financial Modelling Formatting

Financial models often distinguish between assumptions and calculated values.

In the practice workbook, hard-coded assumptions are displayed in blue font.

This helps users identify which values can be changed and which values are calculated by formulas.

Common modelling conventions include:

- Blue font for hard-coded inputs.
- Black font for formulas.
- Green font for links to other worksheets.

These conventions can vary between organisations, so consistency is the most important principle.

## Calculating Gross Revenue

Gross revenue is calculated by multiplying the number of customers by the average order value.

**Formula:**

`Gross Revenue = New Customers × Average Order Value`

### Excel Example

Suppose:

- Cell B16 contains New Customers.
- Cell B17 contains Average Order Value.

In the gross revenue cell, enter:

`=B16*B17`

Press **Enter**.

Excel multiplies the number of new customers by the average order value.

### Format as Currency

1. Select the gross revenue cell.
2. Go to the **Home** tab.
3. Apply the appropriate currency or accounting number format.
4. Remove unnecessary decimal places if required.

### Copy the Formula Across Periods

If the model contains multiple years or periods:

1. Select the formula cell.
2. Use the fill handle to drag the formula across the relevant columns.
3. Check that each formula references the correct period.

Excel adjusts relative cell references automatically.

**Note:** The cell references above are illustrative. Use the actual locations of the assumptions in your workbook.

## Calculating Refunds

Refunds reduce the revenue retained by the business.

Suppose:

- Gross revenue is in cell B6.
- The refund percentage is in cell B18.

To calculate refunds as a negative amount, enter:

`=-B6*B18`

Press **Enter**.

This multiplies gross revenue by the refund percentage and makes the result negative.

### Example

If gross revenue is £100,000 and refunds are 5%:

`=-£100,000*5%`

The refund amount is **-£5,000**.

### Using an Absolute Reference

If the refund percentage is stored in a fixed row that should remain constant when copying the formula across periods, use an absolute row reference.

For example:

`=-B$6*B18`

In this example, row 6 remains fixed while the column can change when the formula is copied horizontally.

If the refund assumption is actually in a separate assumptions section, lock the appropriate reference to that cell instead.

To toggle between reference types while editing a formula, select the cell reference and press **F4**. Depending on the Excel version and keyboard, you may need to use **Fn + F4**.

### Copy the Formula Across Periods

1. Select the refund formula.
2. Drag the fill handle across the relevant periods.
3. Check that the gross revenue and refund assumption references are correct.

## Calculating Discounts

Discounts are calculated in a similar way to refunds when they are modelled as a percentage of gross revenue.

For example, if gross revenue is in B6 and the discount assumption is in B19, enter:

`=-B6*B19`

Adjust the cell references to match your workbook.

The result should be negative because discounts reduce revenue.

## Calculating Net Revenue

Net revenue is the amount remaining after deducting refunds and discounts from gross revenue.

**Formula:**

`Net Revenue = Gross Revenue + Refunds + Discounts`

Because refunds and discounts are entered as negative values, they reduce gross revenue when added.

### Excel Example

Suppose:

- Gross revenue is in B6.
- Refunds are in B7.
- Discounts are in B8.

Enter:

`=SUM(B6:B8)`

Press **Enter**.

### Example

| Item | Amount |
|---|---:|
| Gross Revenue | £100,000 |
| Refunds | -£5,000 |
| Discounts | -£10,000 |
| Net Revenue | £85,000 |

The net revenue is £85,000.

### Copy the Formula Across Periods

Use the fill handle to copy the formula across the remaining periods.

### Apply a Top Border

To distinguish net revenue from the preceding calculations:

1. Select the net revenue cells.
2. Open the **Borders** menu on the Home tab.
3. Select **Top Border**.

This provides a clear visual separation between the revenue calculation and the next section.

## Calculating Cost of Goods Sold

The model calculates individual COGS categories using net revenue and the relevant cost percentages.

For example:

**Materials Cost = Net Revenue × Materials Cost Percentage**

Suppose:

- Net revenue is in cell B9.
- Materials cost percentage is in cell B20.

Enter:

`=B$9*B20`

Press **Enter**.

This calculates the materials cost using net revenue and the relevant percentage assumption.

The row reference `$9` remains fixed when the formula is copied vertically, while the column can change when it is copied horizontally.

**Important:** This formula assumes the cost percentage is based on net revenue. If your workbook defines the percentage using another revenue measure, use that measure instead.

### Copy the Formula

1. Select the materials cost formula.
2. Use the fill handle to copy it across the relevant periods.
3. Copy it down for other COGS categories where the same calculation structure applies.
4. Check that each row refers to the correct cost assumption.

## Calculating Total COGS

Once the individual cost categories have been calculated, add them together.

Suppose the individual COGS values are in cells B12:B14.

Enter:

`=SUM(B12:B14)`

Press **Enter**.

This calculates the total COGS for that period.

### Copy the Formula Across Periods

Use the fill handle to copy the formula across the remaining periods.

### Format Total COGS

To make the total easy to identify:

1. Apply the appropriate number format.
2. Bold the total label and values.
3. Consider adding a top or bottom border for visual separation.

Clear formatting helps readers distinguish individual cost categories from the total.

## Calculating Gross Profit and Gross Margin

Once net revenue and total COGS are available, the model can calculate gross profit and gross margin.

### Gross Profit

**Gross Profit = Net Revenue − Total COGS**

If net revenue is in B9 and total COGS is in B15, enter:

`=B9-B15`

The cell references are illustrative and should be adjusted to your workbook.

### Gross Margin

**Gross Margin = Gross Profit ÷ Net Revenue**

For example, if gross profit is in B16 and net revenue is in B9, enter:

`=B16/B9`

Format the result as a percentage.

### Example

Suppose:

- Net revenue is £85,000.
- Total COGS is £34,000.

Gross profit:

`£85,000 − £34,000 = £51,000`

Gross margin:

`£51,000 ÷ £85,000 = 60%`

The business retains £0.60 in gross profit for every £1 of net revenue before operating expenses and other costs below gross profit.

## Excel Skills Used in This Model

### Basic Arithmetic

Multiplication:

`=B16*B17`

Addition and subtraction:

`=B6+B7+B8`

### SUM Function

Adds values across a range:

`=SUM(B12:B14)`

### Absolute and Mixed References

Examples:

- `$B$6` fixes the column and row.
- `B$6` fixes the row.
- `$B6` fixes the column.

These references help control how formulas change when copied to other cells.

### Fill Handle

The fill handle copies formulas across rows or columns, reducing repetitive work.

### Number Formatting

Currency and percentage formatting make the financial model easier to read.

### Borders and Bold Formatting

Borders and bold text help distinguish totals and separate sections.

## Important Tips

- **Keep assumptions separate from calculations:** This makes the model easier to maintain.
- **Use consistent sign conventions:** Revenue deductions and expenses should be treated consistently.
- **Check cell references:** Make sure formulas refer to the correct assumptions and periods.
- **Use absolute references where necessary:** Fixed references help prevent formulas from shifting incorrectly.
- **Check the formula before copying it:** A formula that works in one cell may reference the wrong row or column after being copied.
- **Avoid unnecessary decimal places:** Display a level of precision appropriate for the report.
- **Use clear formatting:** Distinguish assumptions, calculations and totals consistently.
- **Check the revenue basis:** Ensure cost percentages are applied to the intended revenue measure.

## Key Takeaways

- Gross revenue can be calculated by multiplying customer numbers by average order value.
- Refunds and discounts reduce gross revenue to produce net revenue.
- COGS can be calculated using cost assumptions and the appropriate revenue measure.
- Total COGS is calculated by summing the individual cost categories.
- Gross profit is net revenue minus total COGS.
- Gross margin expresses gross profit as a percentage of net revenue.
- Absolute references and the fill handle make financial modelling more efficient.

## Practice Exercise

Using the `Three Statement Model - Part 2`, `Part 3` and `Part 4` practice workbooks:

1. Review the financial assumptions.
2. Identify new customers and average order value.
3. Calculate gross revenue for each period.
4. Calculate refunds as a negative amount.
5. Calculate discounts as a negative amount.
6. Calculate net revenue using the SUM function.
7. Calculate the individual COGS categories.
8. Calculate total COGS.
9. Calculate gross profit.
10. Calculate gross margin as a percentage.
11. Copy formulas across the remaining periods.
12. Check the formulas and formatting throughout the model.

## When to Use an Income Statement Model

Use an income statement model to forecast revenue, understand the relationship between costs and sales, analyse profitability and provide a foundation for a wider three-statement financial model.
