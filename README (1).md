# RAAST QTBS Transaction Report Generator

A Python script that processes RAAST payment transaction CSV exports and generates:

1. A **full report** (`RAAST_REPORT_*.xlsx`) with raw, incoming, outgoing, and pivoted transaction data.
2. A **QTBS (Query the Bank/Sender) report** (`QTBS_REPORT_*.xlsx`) flagging accounts that meet suspicious-activity thresholds, complete with an auto-generated narrative description per account.

## How It Works

### 1. Input

The script expects a CSV file named in the format:

```
fds_pmts_SADAPKKA_DDMMYYYYHHMM_DDMMYYYYHHMM.csv
```

The two timestamps in the filename (start and end of the reporting period) are parsed via regex and used to name the output files and populate the transaction date span.

Update the `input_file_path` variable at the top of the script to point to your CSV file.

**Required columns in the CSV:**

| Column | Used For |
|---|---|
| `STATUS` | Filtering out `Declined` transactions |
| `CREDITOR PARTICIPANT` / `SENDER PARTICIPANT` | Splitting into incoming vs. outgoing transactions (relative to `SADAPKKA`) |
| `CREDITOR ACCOUNT` / `DEBTOR ACCOUNT` | Pivot grouping key |
| `CREDITOR NAME` / `DEBTOR NAME` | Customer name lookup |
| `AMOUNT` | Sum per account |
| `RISK WEIGHT` | Max risk weight per account |
| `REFERENCE` | Total transaction count |

### 2. Processing Steps

- **Declined count**: Counts all transactions with `STATUS == 'Declined'`.
- **Incoming / Outgoing split**: 
  - Incoming = transactions where `CREDITOR PARTICIPANT == 'SADAPKKA'`
  - Outgoing = transactions where `SENDER PARTICIPANT == 'SADAPKKA'`
  - Declined transactions are excluded from both before pivoting.
- **Pivot tables**: Each side is grouped by account, aggregating summed `AMOUNT`, max `RISK WEIGHT`, and transaction count.
- **QTBS flagging rule** — an account is flagged `QTBS` if **any** of the following is true:
  - Total amount ≥ 10,000 **and** transaction count ≥ 30
  - Max risk weight == 200 **and** transaction count > 5
  - Total amount ≥ 30,000 **and** max risk weight == 200

  *(Note: the incoming rule uses `>= 10000`, the outgoing rule uses `> 10000` — this is intentional in the current script but worth confirming against your compliance requirements.)*

- **Summary stats printed to console**:
  - Total references (`TOTAL`)
  - Sum of transaction counts for flagged incoming accounts (`Incoming SUS`)
  - Sum of transaction counts for flagged outgoing accounts (`Outgoing SUS`)
  - Combined flagged count (`Total SUS`)
  - Non-flagged remainder (`FP`)
  - Declined transaction count

### 3. Outputs

**`RAAST_REPORT_<start>_TO_<end>.xlsx`** — sheets:
- `Original` — full unmodified input data
- `Incoming` / `Outgoing` — raw split data (including declined)
- `Incoming Pivot` / `Outgoing Pivot` — aggregated per-account data with `Comments` column marking `QTBS`

**`QTBS_REPORT_<start>_TO_<end>.xlsx`** — sheets:
- `Incoming QTBS` / `Outgoing QTBS` — only flagged accounts, with columns:
  - `PAN`, `Customer's Name`, `Number of Transactions`, `Billing amount`, `Transaction Date span (start)`, `Description`

  The `Description` column contains an **Excel formula** (not static text) that concatenates the row's own cells into a sentence, e.g.:
  > "*[Name]* received a sum of PKR *[Amount]* from multiple account holders in *[Count]* transactions on the *[fixed_date]*. Please find out the relationship..."

  > ⚠️ The date used inside this description (`fixed_date`) is currently **hardcoded** to `21/06/2026` in the script and does not automatically match the input file's date range — update it manually or parametrize it before each run.

## Requirements

```bash
pip install pandas openpyxl
```

## Usage

1. Place your input CSV in the working directory (or update the path).
2. Edit `input_file_path` at the top of the script to match your filename.
3. Run:

```bash
python generate_qtbs_report.py
```

4. Check the working directory for the two generated `.xlsx` files; console output will show the summary counts.

## Known Limitations / TODO

- `fixed_date` in the Description formulas is hardcoded and should be derived from the filename's date range instead.
- Incoming vs. outgoing QTBS amount thresholds are inconsistent (`>=` vs `>` for the 10,000 check) — verify this is intentional.
- No error handling for missing/malformed columns beyond the filename regex check.
- Assumes a single fixed institution code (`SADAPKKA`) — not currently configurable via CLI args.
