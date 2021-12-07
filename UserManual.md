# Initial Setup

1. Create the "Non Recurring" sheet, which hosts a table starting at cell A1.
  Add the headers "Date", "Description", "Amount", "Account". Add non-recurring
  transactions, chronologically organized.
2. Create the "Balances" sheet, which hosts the balances of all accounts, at
  the beginning and end of the period. Cell A1 should contain "Account", cell
  A2 should contain the date that the balance was held by the account, and cell
  A3 should contain some text (it only matters that it's not empty).
3. Create the "Front" sheet, which mostly contains configuration for the
  script. Configuration sections are titled with the particular function they
  configure (see below) and the section following is configured in a
  function-specific way. An empty row must succeed a configuration
  section--this is how the script differentiates between configuration
  sections.
4. Run the script. The burndown table should be generated in the "Burndown
  Table" sheet.

# Configuration Sections

## Burndown

If the first cell in a record starts with `total:`, the line is taken to be a
total that the Burndown Calculator should calculate. The second cell is a
comma-separated list of configured accounts to sum to get the total. See the
following example:

```
Burndown
total: Assets | Bank of America,Citibank
```
