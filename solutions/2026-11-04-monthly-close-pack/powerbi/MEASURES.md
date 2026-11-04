# Power BI measures — MonthlyClosePack model

Rebuild the same model from the Journal table (`data/journal.csv`) plus
the Accounts table (`data/accounts.csv`, opening balances) in Power BI.
Import both tables; add the MonthKey and SignedAmount calculated
columns to Journal, then the measures below. No DAX in this repo is
executed — these are the definitions the `.pbix` build uses.

## Calculated columns (Journal)

```dax
MonthKey = YEAR(Journal[Date]) * 100 + MONTH(Journal[Date])
SignedAmount = Journal[Debit] - Journal[Credit]
```

## Trial balance

```dax
Opening Signed = SUMX(Accounts, Accounts[Opening Debit] - Accounts[Opening Credit])
Movements Signed = SUM(Journal[SignedAmount])
Closing Signed =
    VAR Acct = SELECTEDVALUE(Accounts[Account])
    RETURN
        CALCULATE(
            SUM(Accounts[Opening Debit]) - SUM(Accounts[Opening Credit])
                + SUM(Journal[SignedAmount]),
            Accounts[Account] = Acct
        )
Closing Dr = MAX([Closing Signed], 0)
Closing Cr = MAX(-[Closing Signed], 0)
```

## Profit & loss

```dax
Revenue = -CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "Revenue")
COGS = CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "COGS")
Gross Profit = [Revenue] - [COGS]
Operating Expenses = CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "Expense")
Net Income = [Gross Profit] - [Operating Expenses]
```

## Balance sheet

```dax
Total Assets =
    CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "Asset")
Total Liabilities =
    -CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "Liability")
Retained Earnings Closing =
    -CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Account] = "Retained Earnings")
        + [Net Income]
Total Equity =
    -CALCULATE(SUM(Accounts[Closing Signed]), Accounts[Type] = "Equity")
        + [Net Income]
Balance Check = [Total Assets] - ([Total Liabilities] + [Total Equity])
```

`Balance Check` must be 0. Drawings is a debit balance inside the
Equity type, so it is already netted inside the `-SUM(Equity)` term —
do not subtract it again.

## Monthly trend

```dax
Monthly Revenue = CALCULATE(SUM(Journal[Credit]), Journal[Category] = "Revenue")
Monthly COGS = CALCULATE(SUM(Journal[Debit]), Journal[Category] = "COGS")
Monthly Opex = CALCULATE(SUM(Journal[Debit]), Journal[Category] = "Opex")
Monthly Net = [Monthly Revenue] - [Monthly COGS] - [Monthly Opex]
```

Slice by `Journal[MonthKey]` (Apr–Oct 2026). Monthly totals must match
the P&L measures; the Budget table joins on MonthKey for variance
visuals (`Sales Variance = [Monthly Revenue] - SUM(Budget[Sales Budget])`).
