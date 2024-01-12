# Balance Calculator Class

## Description

The `BalanceCalculator` class is designed to calculate the current balance based on a given month and row number. It reads values from a table in an Excel worksheet and sums them up from the first negative value. If the sum becomes positive, it resets to zero.

## Usage

First, create an instance of the `BalanceCalculator` class and set the `monthName` and `rowNumber` properties:

```vba
Dim calc As New BalanceCalculator
calc.monthName = "Január"
calc.rowNumber = 1
```

Then, call the `CurrentBalance` method to calculate the balance:

```vba
Debug.Print calc.CurrentBalance()
```

## Methods
### CurrentBalance
This method calculates the current balance. It first checks if the month name and row number are valid. If they are, it gets the values from the table and calculates the sum from the first negative value.

### IsValidMonth
This private method checks if the month name is valid. It returns `True` if the month name is found in the list of months, and `False` otherwise.

### GetColumnNumber
This private method finds the column number for the given month name. If the month name is not found, it raises an error.

### GetValuesFromTable
This private method gets the values from the table for the given column number. It returns an array of values.

### SumFromFirstNegative
This private method calculates the sum from the first negative value in the given array of values. If the sum becomes positive, it resets to zero. It returns the final sum.

## Error Handling
If an invalid month name or row number is provided, or if the month name is not found in the worksheet, an error is raised.