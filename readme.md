# Balance Calculator Class

## Description

The `BalanceCalculator` class is designed to calculate the current balance based on a given month and row number. It reads values from a table in an Excel worksheet and sums them up from the first negative value. If the sum becomes positive, it resets to zero.

## Usage

First, create an instance of the `BalanceCalculator` class and set the `monthName` and `rowNumber` properties:

'''vba
Dim calc As New BalanceCalculator
calc.monthName = "Január"
calc.rowNumber = 1
'''
Then, call the `CurrentBalance` method to calculate the balance:

'''vba
Debug.Print calc.CurrentBalance()
'''

## Methods
### CurrentBalance

This method calculates the current balance. It first checks if the month name and row number are valid. If they are, it gets the values from the table and calculates the sum from the first negative value.

### IsValidMonth

This private method checks if the month name is valid. It returns `True` if the month name is found in the list of months, and `False` otherwise.
