# IC++ Pricing Breakdown Calculator with Transparent Cost Analysis

This script calculates the IC++ pricing breakdown for merchant funding transactions across all regions (HK, MY, TH) with **complete transparency** into all cost components.

## Formula

```
MDR = IC (Interchange) + 1st Plus (Scheme Fee) + 2nd Plus (Acquirer Markup)

Where 2nd Plus includes:
- Gateway Fee (3DS, Capture, Debit, Refund events per NomuPay docs)
- Authorization Fee
- Clearing Fee
- Cross-Border Fee
- Cross-Currency Fee
- Preauthorization Fee
- 3DS & Non-3DS Fees
- Tax Components (VAT, WHT, GRT, ST)
- Net Acquirer Markup (residual after all known costs)

Therefore:
2nd Plus = MDR - IC - 1st Plus
Net Acquirer Markup = 2nd Plus - (Gateway Fee + Auth + Clearing + ... + Taxes)
```

**Special Case:** Malaysia (MY) region has **no scheme fees**, so 1st Plus = 0 for MY transactions.

## Requirements

1. **Python 3.x** (standard library only, no external packages needed)
2. **Transaction Excel file** - Already provided: `Merchant Funding Transactions 2025-12-15,2026-01-08.xlsx`
3. **Fee breakdown CSV** - **YOU NEED TO EXPORT THIS** from TotalProcessing portal

## How to Get Fee Data CSV

### Option 1: Export from TotalProcessing Portal (Recommended)

1. Go to https://support.totalprocessing.com/
2. Navigate to the Funding Report / Fee Reconciliation section
3. Export fee breakdown report for date range: **2025-12-15 to 2026-01-08**
4. Ensure the export includes these columns:
   - `Transaction ID` or `NP Transaction ID` or `Gateway UUID`
   - `MDR Amount` and `MDR Currency`
   - `Interchange Amount` and `Interchange Currency`
   - `Scheme Fee Bucket Amount` and `Scheme Fee Bucket Currency`
   - `Gateway Fee Amount` (optional, for detailed breakdown)
5. Save the CSV file (e.g., as `fees_export.csv`)

### Option 2: Use NomuPay Daily Funding Report

If you have access to NomuPay portal:
1. Download the Daily Funding Report for EID-8520028455
2. Date range: 2025-12-15 to 2026-01-08
3. This report already has the correct format with all fee columns

**Reference format:** See `/Users/manik.soin/Desktop/transaction_reports/daily_funding_report__EID-8520028455_ronaldlam@heroplusgroup.com_POW-344000000003747_2025-12-15 (1).csv`

## Usage

```bash
python3 icpp_calculator.py <excel_file> <fee_csv_file>
```

### Example

```bash
python3 icpp_calculator.py "Merchant Funding Transactions 2025-12-15,2026-01-08.xlsx" fees_export.csv
```

## Required CSV Columns

The fee CSV must contain these columns (column names must match exactly):

| Column Name | Required | Description |
|-------------|----------|-------------|
| `NP Transaction ID` or `Transaction ID` or `Gateway UUID` | ✓ | Transaction identifier for matching |
| `MDR Amount` | ✓ | Total merchant discount rate amount |
| `MDR Currency` | ✓ | Currency for MDR |
| `Interchange Amount` | ✓ | IC (Interchange Fee) |
| `Interchange Currency` | ✓ | Currency for IC |
| `Scheme Fee Bucket Amount` | ✓ | 1st Plus (Scheme Fee) |
| `Scheme Fee Bucket Currency` | ✓ | Currency for scheme fee |
| `Gateway Fee Amount` | Optional | Gateway processing fee |
| `Cross Border Amount` | Optional | Cross-border fee |
| `Authorization Amount` | Optional | Authorization fee |
| `Clearing Amount` | Optional | Clearing fee |

## Output

The script generates two outputs:

### 1. Console Report with Transparent Cost Breakdown

Formatted breakdown by region and card type with **complete transparency**:

```
═══════════════════════════════════════════════════════════
IC++ PRICING BREAKDOWN ANALYSIS
═══════════════════════════════════════════════════════════
Total Transactions Processed: 150

REGION: HONG KONG (HK)
───────────────────────────────────────────────────────────
Card Type: CC
  Count: 65 | Volume: HK$125,000.00

  Fee Breakdown:
    IC (Interchange):           HK$1,250.00  (1.00%)
    1st Plus (Scheme):            HK$250.00  (0.20%)
    2nd Plus (Acquirer):          HK$625.00  (0.50%)
      │
      ├─ 2nd Plus Breakdown:
      │  ├─ Gateway Fee                  HK$390.00  (0.312%)
      │  ├─ Authorization Fee            HK$80.00   (0.064%)
      │  ├─ Clearing Fee                 HK$60.00   (0.048%)
      │  ├─ 3DS Fee                      HK$45.00   (0.036%)
      │  └─ Net Acquirer Markup          HK$50.00   (0.040%)
      │
    ───────────────────────────────────────────────────────
    Total MDR:                  HK$2,125.00  (1.70%)

REGION: MALAYSIA (MY)
───────────────────────────────────────────────────────────
Card Type: CC
  Count: 38 | Volume: RM50,000.00

  Fee Breakdown:
    IC (Interchange):             RM500.00  (1.00%)
    1st Plus (Scheme):              RM0.00  (0.00%)  ⚠️ MY Region
    2nd Plus (Acquirer):           RM300.00  (0.60%)
      │
      ├─ 2nd Plus Breakdown:
      │  ├─ Gateway Fee                  RM180.00   (0.360%)
      │  ├─ Authorization Fee            RM40.00    (0.080%)
      │  ├─ Clearing Fee                 RM30.00    (0.060%)
      │  ├─ 3DS Fee                      RM20.00    (0.040%)
      │  └─ Net Acquirer Markup          RM30.00    (0.060%)
      │
    ───────────────────────────────────────────────────────
    Total MDR:                    RM800.00  (1.60%)
```

**Key Features:**
- ✓ **Tree structure** showing exactly where costs come from
- ✓ **Gateway Fee** transparently broken out (3DS, Capture, Debit events)
- ✓ **Operational fees** (Authorization, Clearing) shown separately
- ✓ **Tax components** (VAT, WHT, GRT, ST) itemized when present
- ✓ **Net Acquirer Markup** calculated as residual after all known costs
- ✓ **Only non-zero fees displayed** - clean, relevant output

### 2. Enhanced CSV Export

File: `icpp_breakdown_report.csv`

Contains **complete transparency** with all cost components:

**Columns:**
- Region, CardType, TxnCount, TotalVolume, Currency
- IC_Total, IC_Avg_Pct
- FirstPlus_Total, FirstPlus_Avg_Pct (will be 0 for MY region)
- SecondPlus_Total, SecondPlus_Avg_Pct
- **GatewayFee_Total** - 3DS, Capture, Debit, Refund events
- **AuthorizationFee_Total** - Authorization costs
- **ClearingFee_Total** - Clearing costs
- **CrossBorderFee_Total** - Cross-border transaction fees
- **CrossCurrencyFee_Total** - Currency conversion fees
- **PreauthFee_Total** - Preauthorization fees
- **ThreeDSFee_Total** - 3D Secure fees
- **NonThreeDSFee_Total** - Non-3DS fees
- **VAT_Total** - Value Added Tax
- **WHT_Total** - Withholding Tax
- **GRT_Total** - Gross Receipt Tax
- **ST_Total** - Sales Tax
- **NetAcquirerMarkup_Total** - Pure acquirer profit after all costs
- MDR_Total, MDR_Avg_Pct

## Transaction Filtering

The script automatically excludes:
- Refund transactions
- Zero-amount transactions
- Declined/Failed transactions
- Transactions missing fee data

All skipped transactions are reported in the summary.

## Malaysia (MY) Special Handling

The script **automatically sets 1st Plus (Scheme Fee) = 0** for all Malaysia transactions, as per the business rule that MY region has no scheme fees. This is indicated with a ⚠️ MY Region marker in the console output.

## Troubleshooting

### "Missing fee data" for many transactions

**Problem:** The fee CSV doesn't cover all transactions in the Excel file.

**Solution:** Export a complete fee breakdown CSV from TotalProcessing portal that includes ALL transactions from 2025-12-15 to 2026-01-08.

### "Invalid amount" errors

**Problem:** Some transactions have missing or invalid amount values.

**Solution:** This is normal for Gateway Events or internal fee records. The script will skip these and report them in the summary.

### Transaction ID mismatch

**Problem:** Transaction IDs in Excel don't match IDs in fee CSV.

**Solution:** Check that both files use the same ID format. The script tries multiple ID column names:
- `NP Transaction ID`
- `Transaction ID`
- `Gateway UUID`
- `Gateway Reference`

## Test Run Results

When tested with the sample CSV (only 3 records), the script correctly reported:

```
Total Transactions Processed: 0
Transactions Skipped: 173

Skipped Breakdown:
  - Missing fee data: 68
  - Status: DECLINED: 13
  - Invalid amount: 90
  - Refund transaction: 2
```

This is **expected** because the sample CSV only contains 3 gateway fee records, not the full transaction set.

## Next Steps

1. **Export complete fee data CSV** from TotalProcessing portal (175 transactions)
2. **Run the script** with both files
3. **Review the console output** to verify calculations
4. **Use the CSV export** for further analysis or reporting

## Understanding the Gateway Fee

Based on NomuPay's documentation, the Gateway Fee includes charges for these event types:

| Event | Description |
|-------|-------------|
| three_d | 3D Secure authentication |
| cp | Capture |
| db | Debit |
| rf | Refund |
| rg | Register |
| pa | Preauthorization |
| rb | Rebill |
| rv | Reversal |
| sd | Schedule |
| dr | Deregister |
| ds | Deschedule |
| rr | Reregister |
| rs | Reschedule |

The script shows the total Gateway Fee from the daily funding report. To reconcile individual events, refer to the `Remarks` column in your funding report which shows event counts like: `"gatewayfee for 100 three_d 100 db 2 rf"`.

## Transparent Cost Analysis Benefits

1. **Identify Cost Drivers** - See which fees contribute most to your 2nd Plus
2. **Compare Regions** - Understand how costs differ between HK, MY, TH
3. **Validate Billing** - Verify gateway events match your transaction volume
4. **Negotiate Better** - Armed with detailed breakdown for acquirer discussions
5. **MY Region Clarity** - Confirm zero scheme fees for Malaysia merchants
6. **Track Net Markup** - Know the pure acquirer profit after all operational costs

## Support

For questions about:
- **Fee data export:** Contact TotalProcessing/NomuPay support
- **Script issues:** Check transaction ID matching and CSV column names
- **IC++ calculation:** Formula is MDR = IC + 1st Plus + 2nd Plus
- **Gateway events:** Refer to NomuPay's Funding Report Gateway Fee Reconciliation docs

## Files in This Directory

- `icpp_calculator.py` - Main script with transparent cost analysis (executable)
- `Merchant Funding Transactions 2025-12-15,2026-01-08.xlsx` - Transaction data (175 records)
- `icpp_breakdown_report.csv` - Output CSV with detailed fee breakdown (generated after run)
- `demo_output_example.txt` - Example of transparent cost output
- `README.md` - This file
