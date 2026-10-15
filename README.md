# Bank Statement Analyzer

Turn a bank statement `.xlsx` (including **password-protected** files) into a clean, categorized dataset, an Excel report and an interactive dashboard. It shows where your money went, who you paid, who paid you, and how your spending changes month to month.

Everything runs locally. Your statement is never uploaded anywhere by this script.

## Features

- Opens password-protected Excel files (via `msoffcrypto-tool`).
- Finds the transaction table automatically, reads every sheet, joins narrations that wrap onto a second line, and removes duplicates from overlapping sheets.
- Works with separate Debit/Credit columns or a single Amount column with a Dr/Cr marker.
- Categorizes each transaction: food, groceries, shopping, travel, fuel, bills, subscriptions, rent, EMI, insurance, investments, cash withdrawals, bank charges, salary, interest, refunds and more.
- Separates UPI **merchant payments** from **person-to-person** transfers.
- Extracts the counterparty (who you paid or who paid you) from UPI, IMPS, NEFT, RTGS and card narrations.
- Leaves self-transfers and credit-card bill payments out of the totals so money moving between your own accounts isn't counted twice.
- Detects recurring payments (subscriptions, EMIs, rent).

## Output

Everything is written to the output folder (default `bank_report/`):

| File | What it is |
|---|---|
| `bank_report.xlsx` | Excel workbook with filterable sheets: Summary, Transactions, Monthly, Spend by Category, Income by Category, Category x Month, Paid To (Payees), Received From (Payers), By Channel, Recurring Payments, Review & Refine |
| `dashboard.html` | Interactive dashboard (month filter, charts, searchable transaction table). Works offline in any browser |
| `transactions_clean.csv` | The cleaned, categorized data |

## Requirements

- Python 3.8+
- `pandas`, `openpyxl`, `msoffcrypto-tool`

```bash
pip install -r requirements.txt
```

## Usage

```bash
python bank_analyzer.py statement.xlsx -o my_report
```

You will be asked for the file password (input is hidden). Other ways to provide it:

```bash
python bank_analyzer.py statement.xlsx -p "yourPassword"
export BANK_XLSX_PASSWORD="yourPassword" && python bank_analyzer.py statement.xlsx
```

| Option | Description |
|---|---|
| `file` | Path to the statement `.xlsx` |
| `-p`, `--password` | File password (otherwise the env var `BANK_XLSX_PASSWORD`, otherwise a hidden prompt) |
| `-o`, `--outdir` | Output folder (default `bank_report`) |
| `--sheet` | Read only this sheet name |

After it finishes, open `dashboard.html` in your browser, or open `bank_report.xlsx` in Excel.

### Google Colab

```python
!pip install -q msoffcrypto-tool
from google.colab import files
files.upload()   # upload bank_analyzer.py and your statement
```

```python
!python bank_analyzer.py "statement.xlsx" -p "YOUR_PASSWORD" -o my_report
!zip -r my_report.zip my_report
files.download("my_report.zip")
```

Extract the zip on your computer before opening `dashboard.html`. If you re-upload a changed `bank_analyzer.py`, delete the old copy first (`!rm bank_analyzer*.py`), because Colab keeps the old file and renames the new one to `bank_analyzer (1).py`.

## Improving the categories

The built-in rules are heuristics, so some transactions will land in the wrong category. To fix that:

1. Run the script and open the **Review & Refine** sheet. It lists the payees the script couldn't place confidently.
2. Edit the config block at the top of `bank_analyzer.py`:

```python
MY_NAMES = ["RAHUL SHARMA"]          # your own names -> tagged "Self Transfer"

CUSTOM_RULES = [
    # (category, regex matched against the narration, side)  side: debit | credit | any
    ("Rent & Housing", r"MR LANDLORD NAME", "debit"),
    ("Salary",         r"YOUR EMPLOYER NAME", "credit"),
]
```

Custom rules run first and override the built-in ones. You can also edit `DEBIT_RULES`, `CREDIT_RULES` and `BRANDS` directly.

3. Re-run the script.

## Privacy

Bank statements are sensitive. Before you push this repo to GitHub:

- Never commit your statement, the generated reports, or your password.
- Add this `.gitignore`:

```gitignore
*.xlsx
*.csv
*.html
bank_report/
my_report/
my_report.zip
```

If you share the dashboard or report with anyone, remember it contains your full transaction history.

## Limitations

- Category and counterparty detection is rule-based and tuned for common Indian bank formats (UPI, IMPS, NEFT, RTGS, card/POS narrations). Other banks and countries will need custom rules.
- The UPI merchant vs person split is a guess based on the name and UPI ID.
- Tested on synthetic statements in an HDFC-style layout. Real statements vary, so check the totals against your bank's own summary.
- Only `.xlsx` is supported (not PDF or CSV statements).

## Troubleshooting

| Problem | Fix |
|---|---|
| `Could not decrypt the file (wrong password?)` | Check the password. Some banks use a pattern such as date of birth or part of your phone number |
| `No transaction table found` | The header row needs a date, a narration/description column and debit/credit/amount columns. Try `--sheet` or share the header row (with private details removed) in an issue |
| `could not map columns` | Column names in your file differ from the usual ones. The skipped sheet's column list is printed so you can adjust the patterns in `load_statement` |
| Dashboard is blank | Re-download the latest script, delete old output, re-run, then extract the zip before opening. A red banner at the top of the page shows any error |

## License

MIT. Add a `LICENSE` file to your repo if you want it to apply.
