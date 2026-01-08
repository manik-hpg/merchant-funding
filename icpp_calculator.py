#!/usr/bin/env python3
"""
IC++ Pricing Breakdown Calculator with Transparent Cost Analysis

Calculates Interchange++ pricing breakdown for merchant funding transactions.
Formula: MDR = IC (Interchange) + 1st Plus (Scheme Fee) + 2nd Plus (Acquirer Markup)

2nd Plus breakdown includes:
- Gateway Fee (3DS, Capture, Debit, Refund, etc.)
- Authorization Fee
- Clearing Fee
- Cross-Border Fee
- Cross-Currency Fee
- Preauthorization Fee
- Tax fees (VAT, WHT, GRT, ST)
- Net Acquirer Markup (residual)

Special handling: Malaysia (MY) region has no scheme fees.

Usage:
    python3 icpp_calculator.py <excel_file> <fee_csv_file>

Example:
    python3 icpp_calculator.py "Merchant Funding Transactions 2025-12-15,2026-01-08.xlsx" fees_export.csv
"""

import zipfile
import xml.etree.ElementTree as ET
import csv
import sys
from collections import defaultdict
from typing import Dict, List, Optional, Tuple


# ============================================================================
# CONFIGURATION
# ============================================================================

REGION_MAPPING = {
    'Hong Kong': 'HK',
    'Malaysia': 'MY',
    'Thailand': 'TH'
}

CURRENCY_SYMBOLS = {
    'HKD': 'HK$',
    'MYR': 'RM',
    'THB': '฿',
    'USD': '$',
    'EUR': '€',
    'GBP': '£'
}

# Excel namespace for parsing
EXCEL_NS = {
    '': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
}


# ============================================================================
# EXCEL READER
# ============================================================================

class ExcelReader:
    """Parse .xlsx files using standard library (zipfile + xml)"""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.shared_strings = []
        self.data = []

    def read(self) -> List[Dict]:
        """
        Read Excel file and return list of transaction dictionaries
        """
        print(f"Reading Excel file: {self.filepath}")

        with zipfile.ZipFile(self.filepath, 'r') as z:
            self._load_shared_strings(z)
            self._load_worksheet(z)

        print(f"Loaded {len(self.data)} transactions from Excel")
        return self.data

    def _load_shared_strings(self, zip_file: zipfile.ZipFile):
        """Extract shared strings from Excel XML"""
        try:
            with zip_file.open('xl/sharedStrings.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # Extract all string values
                for si in root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
                    t = si.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    if t is not None and t.text:
                        self.shared_strings.append(t.text)
                    else:
                        self.shared_strings.append('')

                print(f"Loaded {len(self.shared_strings)} shared strings")
        except KeyError:
            print("No shared strings found in Excel file")

    def _load_worksheet(self, zip_file: zipfile.ZipFile):
        """Parse worksheet XML to extract transaction data"""
        with zip_file.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()

            # Find all rows
            rows = root.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row')

            if len(rows) < 3:
                print("Warning: Excel file has less than 3 rows")
                return

            # Skip first row (title), second row is headers
            header_row = rows[1]
            headers = self._parse_row(header_row)

            print(f"Headers: {headers[:10]}...")  # Print first 10 headers

            # Parse data rows
            for row in rows[2:]:
                row_data = self._parse_row(row)

                # Create transaction dictionary
                transaction = {}
                for i, value in enumerate(row_data):
                    if i < len(headers):
                        transaction[headers[i]] = value

                # Only add if has transaction ID
                if transaction.get('Gateway UUID') or transaction.get('Transaction ID'):
                    self.data.append(transaction)

    def _parse_row(self, row_elem) -> List:
        """Parse a single row element into list of values"""
        cells = row_elem.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c')

        # Create dict of column positions to values
        values_dict = {}

        for cell in cells:
            cell_ref = cell.get('r')  # e.g., "A1", "B1"
            col_letter = ''.join(filter(str.isalpha, cell_ref))
            col_num = self._col_letter_to_num(col_letter)

            value = self._parse_cell_value(cell)
            values_dict[col_num] = value

        # Convert to list (fill missing with empty strings)
        max_col = max(values_dict.keys()) if values_dict else 0
        result = []
        for i in range(max_col + 1):
            result.append(values_dict.get(i, ''))

        return result

    def _parse_cell_value(self, cell_elem) -> str:
        """Parse cell value from XML element"""
        val_elem = cell_elem.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}v')

        if val_elem is None or val_elem.text is None:
            return ''

        cell_type = cell_elem.get('t')

        # String value (reference to shared strings)
        if cell_type == 's':
            idx = int(val_elem.text)
            if 0 <= idx < len(self.shared_strings):
                return self.shared_strings[idx]
            return ''

        # Numeric or other value
        return val_elem.text

    @staticmethod
    def _col_letter_to_num(col_letter: str) -> int:
        """Convert Excel column letter to number (A=0, B=1, etc.)"""
        num = 0
        for char in col_letter:
            num = num * 26 + (ord(char.upper()) - ord('A') + 1)
        return num - 1


# ============================================================================
# FEE DATA LOADER
# ============================================================================

class FeeDataLoader:
    """Load fee breakdown data from CSV"""

    def __init__(self, filepath: str):
        self.filepath = filepath
        self.fee_data = {}

    def load(self) -> Dict[str, Dict]:
        """
        Load fee data from CSV and return lookup dictionary
        Returns: {transaction_id: {mdr, interchange, scheme_fee, ...}}
        """
        print(f"\nLoading fee data from: {self.filepath}")

        with open(self.filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)

            for row in reader:
                # Try different possible transaction ID column names
                txn_id = (row.get('NP Transaction ID') or
                         row.get('Transaction ID') or
                         row.get('Gateway Reference') or
                         row.get('Gateway UUID'))

                if not txn_id:
                    continue

                # Parse fee amounts (handle empty strings)
                self.fee_data[txn_id] = {
                    'mdr_amount': self._parse_float(row.get('MDR Amount', '')),
                    'mdr_currency': row.get('MDR Currency', ''),
                    'interchange_amount': self._parse_float(row.get('Interchange Amount', '')),
                    'interchange_currency': row.get('Interchange Currency', ''),
                    'scheme_fee_amount': self._parse_float(row.get('Scheme Fee Bucket Amount', '')),
                    'scheme_fee_currency': row.get('Scheme Fee Bucket Currency', ''),
                    # Detailed 2nd Plus components
                    'gateway_fee_amount': self._parse_float(row.get('Gateway Fee Amount', '')),
                    'gateway_fee_currency': row.get('Gateway Fee Currency', ''),
                    'authorization_amount': self._parse_float(row.get('Authorization Amount', '')),
                    'authorization_currency': row.get('Authorization Currency', ''),
                    'clearing_amount': self._parse_float(row.get('Clearing Amount', '')),
                    'clearing_currency': row.get('Clearing Currency', ''),
                    'cross_border_amount': self._parse_float(row.get('Cross Border Amount', '')),
                    'cross_border_currency': row.get('Cross Border Currency', ''),
                    'cross_currency_amount': self._parse_float(row.get('Cross Currency Amount', '')),
                    'cross_currency_currency': row.get('Cross Currency Currency', ''),
                    'preauthorization_amount': self._parse_float(row.get('Preauthorization Amount', '')),
                    'preauthorization_currency': row.get('Preauthorization Currency', ''),
                    'three_ds_amount': self._parse_float(row.get('Three Ds Amount', '')),
                    'three_ds_currency': row.get('Three Ds Currency', ''),
                    'non_three_ds_amount': self._parse_float(row.get('Non Three Ds Amount', '')),
                    'non_three_ds_currency': row.get('Non Three Ds Currency', ''),
                    # Tax components
                    'vat_amount': self._parse_float(row.get('VAT Amount', '')),
                    'vat_currency': row.get('VAT Currency', ''),
                    'wht_amount': self._parse_float(row.get('WHT Amount', '')),
                    'wht_currency': row.get('WHT Currency', ''),
                    'grt_amount': self._parse_float(row.get('GRT Amount', '')),
                    'grt_currency': row.get('GRT Currency', ''),
                    'st_amount': self._parse_float(row.get('ST Amount', '')),
                    'st_currency': row.get('ST Currency', '')
                }

        print(f"Loaded fee data for {len(self.fee_data)} transactions")
        return self.fee_data

    @staticmethod
    def _parse_float(value: str) -> float:
        """Parse float from string, return 0.0 if invalid"""
        try:
            return float(value) if value else 0.0
        except (ValueError, TypeError):
            return 0.0


# ============================================================================
# REGION IDENTIFIER
# ============================================================================

def identify_region(merchant: str, card_country: str = '') -> str:
    """
    Determine region from merchant name or card country

    Args:
        merchant: Merchant name
        card_country: Card country code (fallback)

    Returns:
        Region code (HK, MY, TH, or UNKNOWN)
    """
    if not merchant:
        merchant = ''

    for keyword, code in REGION_MAPPING.items():
        if keyword in merchant:
            return code

    # Fallback to card country
    if card_country:
        if card_country in ['HKG', 'HK']:
            return 'HK'
        elif card_country in ['MYS', 'MY']:
            return 'MY'
        elif card_country in ['THA', 'TH']:
            return 'TH'

    return 'UNKNOWN'


# ============================================================================
# IC++ CALCULATOR
# ============================================================================

def calculate_icpp(transaction: Dict, fees: Dict, region: str) -> Optional[Dict]:
    """
    Calculate IC++ breakdown for a transaction with detailed 2nd Plus components

    Args:
        transaction: Transaction dictionary
        fees: Fee data dictionary
        region: Region code

    Returns:
        Dictionary with IC, 1st Plus, 2nd Plus breakdown and detailed sub-components
    """
    # Get primary fee components
    mdr = fees.get('mdr_amount', 0.0)
    interchange = fees.get('interchange_amount', 0.0)
    scheme_fee = fees.get('scheme_fee_amount', 0.0)

    # MY region special case - NO scheme fees
    if region == 'MY':
        scheme_fee = 0.0

    # Get detailed 2nd Plus components
    gateway_fee = fees.get('gateway_fee_amount', 0.0)
    authorization_fee = fees.get('authorization_amount', 0.0)
    clearing_fee = fees.get('clearing_amount', 0.0)
    cross_border_fee = fees.get('cross_border_amount', 0.0)
    cross_currency_fee = fees.get('cross_currency_amount', 0.0)
    preauth_fee = fees.get('preauthorization_amount', 0.0)
    three_ds_fee = fees.get('three_ds_amount', 0.0)
    non_three_ds_fee = fees.get('non_three_ds_amount', 0.0)

    # Tax components
    vat = fees.get('vat_amount', 0.0)
    wht = fees.get('wht_amount', 0.0)
    grt = fees.get('grt_amount', 0.0)
    st = fees.get('st_amount', 0.0)

    # Calculate total 2nd Plus (Acquirer Markup)
    second_plus_total = mdr - interchange - scheme_fee

    # Calculate known components total
    known_components = (gateway_fee + authorization_fee + clearing_fee +
                       cross_border_fee + cross_currency_fee + preauth_fee +
                       three_ds_fee + non_three_ds_fee + vat + wht + grt + st)

    # Net acquirer markup is what's left after all known fees
    net_acquirer_markup = second_plus_total - known_components

    return {
        'ic': interchange,
        'first_plus': scheme_fee,
        'second_plus': second_plus_total,
        'mdr': mdr,
        # Detailed 2nd Plus breakdown
        'second_plus_details': {
            'gateway_fee': gateway_fee,
            'authorization_fee': authorization_fee,
            'clearing_fee': clearing_fee,
            'cross_border_fee': cross_border_fee,
            'cross_currency_fee': cross_currency_fee,
            'preauth_fee': preauth_fee,
            'three_ds_fee': three_ds_fee,
            'non_three_ds_fee': non_three_ds_fee,
            'vat': vat,
            'wht': wht,
            'grt': grt,
            'st': st,
            'net_acquirer_markup': net_acquirer_markup
        }
    }


# ============================================================================
# TRANSACTION FILTER
# ============================================================================

def is_valid_transaction(transaction: Dict) -> Tuple[bool, str]:
    """
    Check if transaction should be included in analysis

    Returns:
        (is_valid, reason_if_invalid)
    """
    # Check payment type
    payment_type = str(transaction.get('Payment Type', '')).upper()
    if 'REFUND' in payment_type or payment_type == 'RF':
        return False, 'Refund transaction'

    # Check transaction amount
    amount_str = str(transaction.get('Amount', '0'))
    try:
        amount = float(amount_str)
        if amount == 0:
            return False, 'Zero amount'
    except (ValueError, TypeError):
        return False, 'Invalid amount'

    # Check processor status
    status = str(transaction.get('Processor Status', '')).upper()
    if status in ['DECLINED', 'FAILED', 'ERROR']:
        return False, f'Status: {status}'

    return True, ''


# ============================================================================
# AGGREGATOR
# ============================================================================

class StatisticsAggregator:
    """Aggregate IC++ statistics by region and card type"""

    def __init__(self):
        self.stats = defaultdict(lambda: defaultdict(lambda: {
            'count': 0,
            'total_volume': 0.0,
            'total_ic': 0.0,
            'total_first_plus': 0.0,
            'total_second_plus': 0.0,
            'total_mdr': 0.0,
            'currency': '',
            'transactions': [],
            # Detailed 2nd Plus components
            'total_gateway_fee': 0.0,
            'total_authorization_fee': 0.0,
            'total_clearing_fee': 0.0,
            'total_cross_border_fee': 0.0,
            'total_cross_currency_fee': 0.0,
            'total_preauth_fee': 0.0,
            'total_three_ds_fee': 0.0,
            'total_non_three_ds_fee': 0.0,
            'total_vat': 0.0,
            'total_wht': 0.0,
            'total_grt': 0.0,
            'total_st': 0.0,
            'total_net_acquirer_markup': 0.0
        }))

        self.skipped_count = 0
        self.skipped_reasons = defaultdict(int)

    def add_transaction(self, transaction: Dict, icpp: Dict, region: str):
        """Add a transaction to statistics"""
        card_type = str(transaction.get('Card Type', 'UNKNOWN')).upper()
        currency = str(transaction.get('Currency', ''))
        amount = float(transaction.get('Amount', 0))

        # Get or create bucket
        bucket = self.stats[region][card_type]

        # Update primary counters
        bucket['count'] += 1
        bucket['total_volume'] += amount
        bucket['total_ic'] += icpp['ic']
        bucket['total_first_plus'] += icpp['first_plus']
        bucket['total_second_plus'] += icpp['second_plus']
        bucket['total_mdr'] += icpp['mdr']
        bucket['currency'] = currency

        # Update detailed 2nd Plus component counters
        details = icpp['second_plus_details']
        bucket['total_gateway_fee'] += details['gateway_fee']
        bucket['total_authorization_fee'] += details['authorization_fee']
        bucket['total_clearing_fee'] += details['clearing_fee']
        bucket['total_cross_border_fee'] += details['cross_border_fee']
        bucket['total_cross_currency_fee'] += details['cross_currency_fee']
        bucket['total_preauth_fee'] += details['preauth_fee']
        bucket['total_three_ds_fee'] += details['three_ds_fee']
        bucket['total_non_three_ds_fee'] += details['non_three_ds_fee']
        bucket['total_vat'] += details['vat']
        bucket['total_wht'] += details['wht']
        bucket['total_grt'] += details['grt']
        bucket['total_st'] += details['st']
        bucket['total_net_acquirer_markup'] += details['net_acquirer_markup']

        # Store transaction reference
        bucket['transactions'].append({
            'id': transaction.get('Gateway UUID', 'N/A'),
            'amount': amount,
            'icpp': icpp
        })

    def skip_transaction(self, reason: str):
        """Record a skipped transaction"""
        self.skipped_count += 1
        self.skipped_reasons[reason] += 1

    def calculate_percentages(self):
        """Calculate average percentages for all buckets"""
        for region in self.stats:
            for card_type in self.stats[region]:
                bucket = self.stats[region][card_type]
                volume = bucket['total_volume']

                if volume > 0:
                    bucket['avg_ic_pct'] = abs(bucket['total_ic'] / volume * 100)
                    bucket['avg_first_plus_pct'] = abs(bucket['total_first_plus'] / volume * 100)
                    bucket['avg_second_plus_pct'] = abs(bucket['total_second_plus'] / volume * 100)
                    bucket['avg_mdr_pct'] = abs(bucket['total_mdr'] / volume * 100)
                else:
                    bucket['avg_ic_pct'] = 0.0
                    bucket['avg_first_plus_pct'] = 0.0
                    bucket['avg_second_plus_pct'] = 0.0
                    bucket['avg_mdr_pct'] = 0.0

    def get_stats(self) -> Dict:
        """Get aggregated statistics"""
        return dict(self.stats)


# ============================================================================
# REPORT GENERATOR
# ============================================================================

class ReportGenerator:
    """Generate console and CSV reports"""

    @staticmethod
    def format_currency(amount: float, currency: str) -> str:
        """Format currency with proper symbol"""
        symbol = CURRENCY_SYMBOLS.get(currency, currency + ' ')
        return f"{symbol}{abs(amount):,.2f}"

    def print_console_report(self, stats: Dict, aggregator: StatisticsAggregator):
        """Print formatted console report"""
        print('\n')
        print('═' * 65)
        print('IC++ PRICING BREAKDOWN ANALYSIS')
        print('═' * 65)

        # Calculate totals
        total_transactions = sum(
            bucket['count']
            for region in stats.values()
            for bucket in region.values()
        )

        print(f"Total Transactions Processed: {total_transactions}")
        print(f"Transactions Skipped: {aggregator.skipped_count}")

        if aggregator.skipped_reasons:
            print(f"\nSkipped Breakdown:")
            for reason, count in aggregator.skipped_reasons.items():
                print(f"  - {reason}: {count}")

        print('═' * 65)

        # Report by region
        region_order = ['HK', 'MY', 'TH', 'UNKNOWN']
        region_names = {
            'HK': 'HONG KONG',
            'MY': 'MALAYSIA',
            'TH': 'THAILAND',
            'UNKNOWN': 'OTHER/UNKNOWN'
        }

        for region_code in region_order:
            if region_code not in stats:
                continue

            region_name = region_names.get(region_code, region_code)
            print(f'\n\nREGION: {region_name} ({region_code})')
            print('─' * 65)

            for card_type in sorted(stats[region_code].keys()):
                data = stats[region_code][card_type]

                if data['count'] == 0:
                    continue

                currency = data['currency']

                print(f'\nCard Type: {card_type}')
                print(f'  Count: {data["count"]} | Volume: {self.format_currency(data["total_volume"], currency)}')
                print(f'\n  Fee Breakdown:')
                print(f'    IC (Interchange):      {self.format_currency(data["total_ic"], currency):>15}  ({data["avg_ic_pct"]:.2f}%)')

                # Special indicator for MY region
                if region_code == 'MY':
                    print(f'    1st Plus (Scheme):     {self.format_currency(data["total_first_plus"], currency):>15}  ({data["avg_first_plus_pct"]:.2f}%)  ⚠️ MY Region')
                else:
                    print(f'    1st Plus (Scheme):     {self.format_currency(data["total_first_plus"], currency):>15}  ({data["avg_first_plus_pct"]:.2f}%)')

                print(f'    2nd Plus (Acquirer):   {self.format_currency(data["total_second_plus"], currency):>15}  ({data["avg_second_plus_pct"]:.2f}%)')

                # Show detailed breakdown of 2nd Plus
                self._print_second_plus_breakdown(data, currency)

                print(f'    {"─" * 55}')
                print(f'    Total MDR:             {self.format_currency(data["total_mdr"], currency):>15}  ({data["avg_mdr_pct"]:.2f}%)')

    def _print_second_plus_breakdown(self, data: Dict, currency: str):
        """Print detailed breakdown of 2nd Plus components"""
        volume = data['total_volume']

        # Helper to format and show only non-zero components
        def show_component(label: str, amount: float, indent: str = '      '):
            if abs(amount) > 0.001:  # Show if greater than 0.001
                pct = abs(amount / volume * 100) if volume > 0 else 0.0
                print(f'{indent}├─ {label:<26} {self.format_currency(amount, currency):>12}  ({pct:.3f}%)')

        print(f'      │')
        print(f'      ├─ 2nd Plus Breakdown:')

        # Operational fees
        show_component('Gateway Fee', data['total_gateway_fee'], '      │  ')
        show_component('Authorization Fee', data['total_authorization_fee'], '      │  ')
        show_component('Clearing Fee', data['total_clearing_fee'], '      │  ')
        show_component('Cross-Border Fee', data['total_cross_border_fee'], '      │  ')
        show_component('Cross-Currency Fee', data['total_cross_currency_fee'], '      │  ')
        show_component('Preauthorization Fee', data['total_preauth_fee'], '      │  ')
        show_component('3DS Fee', data['total_three_ds_fee'], '      │  ')
        show_component('Non-3DS Fee', data['total_non_three_ds_fee'], '      │  ')

        # Tax components
        tax_total = (data['total_vat'] + data['total_wht'] +
                    data['total_grt'] + data['total_st'])
        if abs(tax_total) > 0.001:
            print(f'      │  ├─ Tax Components:')
            show_component('VAT', data['total_vat'], '      │  │  ')
            show_component('WHT (Withholding Tax)', data['total_wht'], '      │  │  ')
            show_component('GRT (Gross Receipt Tax)', data['total_grt'], '      │  │  ')
            show_component('ST (Sales Tax)', data['total_st'], '      │  │  ')

        # Net acquirer markup (residual)
        markup = data['total_net_acquirer_markup']
        if abs(markup) > 0.001:
            markup_pct = abs(markup / volume * 100) if volume > 0 else 0.0
            print(f'      │  └─ {"Net Acquirer Markup":<26} {self.format_currency(markup, currency):>12}  ({markup_pct:.3f}%)')

        print(f'      │')

        print('\n' + '═' * 65)

    def export_csv(self, stats: Dict, output_path: str):
        """Export statistics to CSV file with detailed breakdown"""
        print(f"\nExporting CSV report to: {output_path}")

        with open(output_path, 'w', newline='', encoding='utf-8') as f:
            fieldnames = [
                'Region', 'CardType', 'TxnCount', 'TotalVolume', 'Currency',
                'IC_Total', 'IC_Avg_Pct',
                'FirstPlus_Total', 'FirstPlus_Avg_Pct',
                'SecondPlus_Total', 'SecondPlus_Avg_Pct',
                # Detailed 2nd Plus components
                'GatewayFee_Total', 'AuthorizationFee_Total', 'ClearingFee_Total',
                'CrossBorderFee_Total', 'CrossCurrencyFee_Total', 'PreauthFee_Total',
                'ThreeDSFee_Total', 'NonThreeDSFee_Total',
                'VAT_Total', 'WHT_Total', 'GRT_Total', 'ST_Total',
                'NetAcquirerMarkup_Total',
                'MDR_Total', 'MDR_Avg_Pct'
            ]

            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()

            for region in sorted(stats.keys()):
                for card_type in sorted(stats[region].keys()):
                    data = stats[region][card_type]

                    if data['count'] == 0:
                        continue

                    writer.writerow({
                        'Region': region,
                        'CardType': card_type,
                        'TxnCount': data['count'],
                        'TotalVolume': f"{data['total_volume']:.2f}",
                        'Currency': data['currency'],
                        'IC_Total': f"{data['total_ic']:.2f}",
                        'IC_Avg_Pct': f"{data['avg_ic_pct']:.2f}",
                        'FirstPlus_Total': f"{data['total_first_plus']:.2f}",
                        'FirstPlus_Avg_Pct': f"{data['avg_first_plus_pct']:.2f}",
                        'SecondPlus_Total': f"{data['total_second_plus']:.2f}",
                        'SecondPlus_Avg_Pct': f"{data['avg_second_plus_pct']:.2f}",
                        # Detailed 2nd Plus components
                        'GatewayFee_Total': f"{data['total_gateway_fee']:.2f}",
                        'AuthorizationFee_Total': f"{data['total_authorization_fee']:.2f}",
                        'ClearingFee_Total': f"{data['total_clearing_fee']:.2f}",
                        'CrossBorderFee_Total': f"{data['total_cross_border_fee']:.2f}",
                        'CrossCurrencyFee_Total': f"{data['total_cross_currency_fee']:.2f}",
                        'PreauthFee_Total': f"{data['total_preauth_fee']:.2f}",
                        'ThreeDSFee_Total': f"{data['total_three_ds_fee']:.2f}",
                        'NonThreeDSFee_Total': f"{data['total_non_three_ds_fee']:.2f}",
                        'VAT_Total': f"{data['total_vat']:.2f}",
                        'WHT_Total': f"{data['total_wht']:.2f}",
                        'GRT_Total': f"{data['total_grt']:.2f}",
                        'ST_Total': f"{data['total_st']:.2f}",
                        'NetAcquirerMarkup_Total': f"{data['total_net_acquirer_markup']:.2f}",
                        'MDR_Total': f"{data['total_mdr']:.2f}",
                        'MDR_Avg_Pct': f"{data['avg_mdr_pct']:.2f}"
                    })

        print(f"CSV export complete")

    def export_html(self, stats: Dict, aggregator: StatisticsAggregator, output_path: str):
        """Export visual HTML report with actual data"""
        print(f"\nExporting HTML visual report to: {output_path}")

        # Calculate totals
        total_transactions = sum(
            bucket['count']
            for region in stats.values()
            for bucket in region.values()
        )

        total_regions = len(stats)

        # Region metadata
        region_meta = {
            'HK': {'name': 'Hong Kong', 'flag': '🇭🇰', 'currency_symbol': 'HK$'},
            'MY': {'name': 'Malaysia', 'flag': '🇲🇾', 'currency_symbol': 'RM'},
            'TH': {'name': 'Thailand', 'flag': '🇹🇭', 'currency_symbol': '฿'},
            'UNKNOWN': {'name': 'Other', 'flag': '🌐', 'currency_symbol': '$'}
        }

        html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>IC++ Pricing Breakdown - Regional Analysis</title>
    <style>
        * {{ margin: 0; padding: 0; box-sizing: border-box; }}
        body {{
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 40px 20px;
            color: #333;
        }}
        .container {{ max-width: 1400px; margin: 0 auto; }}
        .header {{
            text-align: center;
            color: white;
            margin-bottom: 50px;
        }}
        .header h1 {{
            font-size: 42px;
            font-weight: 700;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
        }}
        .header p {{ font-size: 18px; opacity: 0.9; }}
        .stats-overview {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            margin-bottom: 40px;
        }}
        .stat-card {{
            background: white;
            padding: 25px;
            border-radius: 15px;
            box-shadow: 0 10px 30px rgba(0,0,0,0.1);
            text-align: center;
        }}
        .stat-card h3 {{
            color: #667eea;
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 10px;
        }}
        .stat-card .value {{
            font-size: 36px;
            font-weight: 700;
            color: #333;
        }}
        .regions-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(400px, 1fr));
            gap: 30px;
            margin-bottom: 40px;
        }}
        .region-card {{
            background: white;
            border-radius: 20px;
            padding: 35px;
            box-shadow: 0 15px 40px rgba(0,0,0,0.15);
            transition: transform 0.3s ease, box-shadow 0.3s ease;
        }}
        .region-card:hover {{
            transform: translateY(-5px);
            box-shadow: 0 20px 50px rgba(0,0,0,0.2);
        }}
        .region-header {{
            display: flex;
            align-items: center;
            justify-content: space-between;
            margin-bottom: 25px;
            padding-bottom: 20px;
            border-bottom: 3px solid #f0f0f0;
        }}
        .region-title {{
            display: flex;
            align-items: center;
            gap: 15px;
        }}
        .region-flag {{ font-size: 36px; }}
        .region-name {{
            font-size: 28px;
            font-weight: 700;
            color: #333;
        }}
        .region-code {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 8px 16px;
            border-radius: 20px;
            font-size: 14px;
            font-weight: 600;
        }}
        .region-stats {{
            display: flex;
            justify-content: space-around;
            margin-bottom: 25px;
            padding: 20px;
            background: #f8f9ff;
            border-radius: 12px;
        }}
        .region-stat {{ text-align: center; }}
        .region-stat-label {{
            font-size: 12px;
            color: #888;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 5px;
        }}
        .region-stat-value {{
            font-size: 24px;
            font-weight: 700;
            color: #333;
        }}
        .fee-breakdown {{ margin-top: 25px; }}
        .fee-item {{ margin-bottom: 15px; }}
        .fee-header {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 8px;
        }}
        .fee-label {{
            font-size: 16px;
            font-weight: 600;
            color: #333;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .fee-badge {{
            background: #ffd700;
            color: #333;
            padding: 4px 10px;
            border-radius: 12px;
            font-size: 11px;
            font-weight: 700;
        }}
        .fee-amount {{
            display: flex;
            gap: 15px;
            align-items: center;
        }}
        .fee-value {{
            font-size: 18px;
            font-weight: 700;
            color: #333;
        }}
        .fee-percentage {{
            font-size: 14px;
            color: #888;
            background: #f0f0f0;
            padding: 4px 10px;
            border-radius: 8px;
        }}
        .fee-bar {{
            height: 8px;
            background: #f0f0f0;
            border-radius: 4px;
            overflow: hidden;
        }}
        .fee-bar-fill {{
            height: 100%;
            border-radius: 4px;
            transition: width 0.8s ease;
        }}
        .fee-bar-ic {{ background: linear-gradient(90deg, #667eea 0%, #764ba2 100%); }}
        .fee-bar-first {{ background: linear-gradient(90deg, #f093fb 0%, #f5576c 100%); }}
        .fee-bar-second {{ background: linear-gradient(90deg, #4facfe 0%, #00f2fe 100%); }}
        .fee-tree {{
            margin-left: 30px;
            border-left: 3px solid #e0e0e0;
            padding-left: 20px;
            margin-top: 15px;
        }}
        .fee-tree-item {{
            margin-bottom: 12px;
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 10px;
            background: #fafafa;
            border-radius: 8px;
            font-size: 14px;
        }}
        .fee-tree-label {{
            color: #666;
            display: flex;
            align-items: center;
            gap: 8px;
        }}
        .fee-tree-icon {{ color: #999; }}
        .fee-tree-amount {{
            font-weight: 600;
            color: #333;
        }}
        .total-section {{
            margin-top: 25px;
            padding: 20px;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 12px;
            color: white;
        }}
        .total-label {{
            font-size: 14px;
            text-transform: uppercase;
            letter-spacing: 1px;
            opacity: 0.9;
            margin-bottom: 5px;
        }}
        .total-value {{
            font-size: 32px;
            font-weight: 700;
        }}
        .footer {{
            text-align: center;
            color: white;
            margin-top: 40px;
            padding-top: 30px;
            border-top: 1px solid rgba(255,255,255,0.2);
        }}
        .footer p {{
            opacity: 0.8;
            font-size: 14px;
        }}
        .alert {{
            background: #fff3cd;
            border: 2px solid #ffc107;
            padding: 20px;
            border-radius: 12px;
            margin-bottom: 30px;
            color: #856404;
        }}
        .alert h3 {{
            margin-bottom: 10px;
            color: #856404;
        }}
        .loading-overlay {{
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.7);
            display: none;
            justify-content: center;
            align-items: center;
            z-index: 2000;
        }}
        .loading-content {{
            background: white;
            padding: 40px 60px;
            border-radius: 15px;
            text-align: center;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }}
        .spinner {{
            border: 4px solid #f3f3f3;
            border-top: 4px solid #667eea;
            border-radius: 50%;
            width: 50px;
            height: 50px;
            animation: spin 1s linear infinite;
            margin: 0 auto 20px auto;
        }}
        @keyframes spin {{
            0% {{ transform: rotate(0deg); }}
            100% {{ transform: rotate(360deg); }}
        }}
        .loading-text {{
            font-size: 18px;
            font-weight: 600;
            color: #333;
        }}
        .clickable {{
            cursor: pointer;
            position: relative;
        }}
        .clickable:hover {{
            background: #f8f9ff;
            border-radius: 8px;
        }}
        .info-icon {{
            display: inline-block;
            width: 18px;
            height: 18px;
            background: #667eea;
            color: white;
            border-radius: 50%;
            text-align: center;
            line-height: 18px;
            font-size: 12px;
            margin-left: 5px;
            cursor: help;
        }}
        .modal {{
            display: none;
            position: fixed;
            z-index: 1000;
            left: 0;
            top: 0;
            width: 100%;
            height: 100%;
            background: rgba(0,0,0,0.5);
            animation: fadeIn 0.3s;
        }}
        .modal-content {{
            background: white;
            margin: 5% auto;
            padding: 0;
            border-radius: 15px;
            max-width: 700px;
            max-height: 85vh;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            animation: slideDown 0.3s;
            display: flex;
            flex-direction: column;
        }}
        .modal-header {{
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px 30px;
            border-radius: 15px 15px 0 0;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }}
        .modal-header h2 {{
            margin: 0;
            font-size: 24px;
        }}
        .close {{
            color: white;
            font-size: 32px;
            font-weight: bold;
            cursor: pointer;
            line-height: 1;
            transition: transform 0.2s;
        }}
        .close:hover {{
            transform: scale(1.2);
        }}
        .modal-body {{
            padding: 30px;
            overflow-y: auto;
            flex: 1;
        }}
        .calculation-section {{
            margin-bottom: 25px;
        }}
        .calculation-section h3 {{
            color: #667eea;
            font-size: 16px;
            margin-bottom: 15px;
            text-transform: uppercase;
            letter-spacing: 1px;
        }}
        .calc-row {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 12px 15px;
            background: #f8f9ff;
            border-radius: 8px;
            margin-bottom: 8px;
        }}
        .calc-label {{
            font-weight: 600;
            color: #333;
        }}
        .calc-value {{
            font-family: 'Courier New', monospace;
            color: #667eea;
            font-weight: 600;
        }}
        .column-tag {{
            display: inline-block;
            background: #e8eaf6;
            color: #5e35b1;
            padding: 4px 10px;
            border-radius: 6px;
            font-size: 13px;
            font-family: 'Courier New', monospace;
            margin: 4px 4px 4px 0;
        }}
        .formula-box {{
            background: #f5f5f5;
            border-left: 4px solid #667eea;
            padding: 15px;
            border-radius: 8px;
            font-family: 'Courier New', monospace;
            margin: 15px 0;
        }}
        @keyframes fadeIn {{
            from {{ opacity: 0; }}
            to {{ opacity: 1; }}
        }}
        @keyframes slideDown {{
            from {{ transform: translateY(-50px); opacity: 0; }}
            to {{ transform: translateY(0); opacity: 1; }}
        }}
        @media (max-width: 768px) {{
            .regions-grid {{ grid-template-columns: 1fr; }}
            .header h1 {{ font-size: 32px; }}
            .modal-content {{ margin: 10% 15px; }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>💳 IC++ Pricing Breakdown</h1>
            <p>Transparent Cost Analysis - Actual Data</p>
        </div>

        <!-- File Upload Section -->
        <div id="uploadSection" class="stat-card" style="max-width: 800px; margin: 0 auto 40px auto; padding: 40px;">
            <h3 style="font-size: 20px; margin-bottom: 20px; color: #667eea;">📂 Load Daily Funding Report</h3>

            <!-- Preset Selector -->
            <div style="margin-bottom: 25px;">
                <label style="display: block; font-weight: 600; margin-bottom: 10px; color: #333;">
                    🎯 Quick Load Presets:
                </label>
                <select id="presetSelector" style="width: 100%; padding: 12px; border: 2px solid #667eea; border-radius: 8px; font-size: 15px; background: white; cursor: pointer;">
                    <option value="">-- Select a preset report --</option>
                    <option value="daily_funding_report__EID-8520028455_ theresalam@heroplusgroup.com_POW-344000000003747_2026-01-06.csv">2026-01-06 (Theresa) - Latest</option>
                    <option value="daily_funding_report__EID-8520028455_ronaldlam@heroplusgroup.com_POW-344000000003747_2025-12-23.csv">2025-12-23 (Ronald)</option>
                    <option value="daily_funding_report__EID-8520028455_ronaldlam@heroplusgroup.com_POW-344000000003747_2025-12-22.csv">2025-12-22 (Ronald)</option>
                    <option value="daily_funding_report__EID-8520028455_ronaldlam@heroplusgroup.com_POW-344000000003747_2025-12-18.csv">2025-12-18 (Ronald)</option>
                    <option value="daily_funding_report__EID-8520028455_ronaldlam@heroplusgroup.com_POW-344000000003747_2025-12-15.csv">2025-12-15 (Ronald)</option>
                    <option value="daily_funding_report__EID-600004504_ theresalam@heroplusgroup.com_POW-458000000000274_2026-01-06.csv">2026-01-06 (EID-600004504)</option>
                </select>
                <button id="loadPresetBtn" style="margin-top: 10px; width: 100%; padding: 12px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; border-radius: 8px; font-size: 15px; font-weight: 600; cursor: pointer; transition: transform 0.2s;">
                    Load Selected Report
                </button>
            </div>

            <div style="text-align: center; margin: 20px 0; color: #888;">
                <span style="background: white; padding: 0 15px;">OR</span>
            </div>

            <!-- File Upload -->
            <div id="dropZone" style="border: 3px dashed #667eea; border-radius: 12px; padding: 40px; text-align: center; background: #f8f9ff; cursor: pointer; transition: all 0.3s;">
                <div style="font-size: 48px; margin-bottom: 15px;">📄</div>
                <p style="font-size: 16px; color: #333; margin-bottom: 10px;">
                    <strong>Drag & Drop CSV here</strong> or click to browse
                </p>
                <p style="font-size: 14px; color: #888;">
                    NomuPay Daily Funding Report with 65 columns
                </p>
                <input type="file" id="fileInput" accept=".csv" style="display: none;">
            </div>
            <div id="uploadStatus" style="margin-top: 20px; display: none;">
                <div style="background: #4caf50; color: white; padding: 15px; border-radius: 8px; text-align: center;">
                    <strong>✓ File loaded:</strong> <span id="fileName"></span>
                    <br>
                    <span id="fileStats"></span>
                </div>
            </div>
        </div>

        <div class="stats-overview">
            <div class="stat-card">
                <h3>Total Transactions</h3>
                <div class="value">{total_transactions}</div>
            </div>
            <div class="stat-card">
                <h3>Regions Analyzed</h3>
                <div class="value">{total_regions}</div>
            </div>
            <div class="stat-card">
                <h3>Transactions Skipped</h3>
                <div class="value">{aggregator.skipped_count}</div>
            </div>
            <div class="stat-card">
                <h3>Transparency</h3>
                <div class="value">100%</div>
            </div>
        </div>
"""

        # Generate region cards
        html += '        <div class="regions-grid">\n'

        for region_code in sorted(stats.keys()):
            region_data = stats[region_code]
            meta = region_meta.get(region_code, region_meta['UNKNOWN'])

            for card_type in sorted(region_data.keys()):
                data = region_data[card_type]

                if data['count'] == 0:
                    continue

                currency = data['currency']
                currency_symbol = meta['currency_symbol']

                # Calculate bar widths (as percentage of MDR)
                mdr_abs = abs(data['total_mdr'])
                ic_width = abs(data['total_ic']) / mdr_abs * 100 if mdr_abs > 0 else 0
                first_width = abs(data['total_first_plus']) / mdr_abs * 100 if mdr_abs > 0 else 0
                second_width = abs(data['total_second_plus']) / mdr_abs * 100 if mdr_abs > 0 else 0

                html += f'''            <div class="region-card">
                <div class="region-header">
                    <div class="region-title">
                        <span class="region-flag">{meta['flag']}</span>
                        <span class="region-name">{meta['name']}</span>
                    </div>
                    <div class="region-code">{region_code}</div>
                </div>

                <div class="region-stats">
                    <div class="region-stat">
                        <div class="region-stat-label">Volume</div>
                        <div class="region-stat-value">{self.format_currency(data['total_volume'], currency)}</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-label">Transactions</div>
                        <div class="region-stat-value">{data['count']}</div>
                    </div>
                    <div class="region-stat">
                        <div class="region-stat-label">Card Type</div>
                        <div class="region-stat-value">{card_type}</div>
                    </div>
                </div>

                <div class="fee-breakdown">
                    <div class="fee-item clickable" onclick="showModal('ic', '{region_code}', '{card_type}')">
                        <div class="fee-header">
                            <div class="fee-label">
                                💱 IC (Interchange)
                                <span class="info-icon">i</span>
                            </div>
                            <div class="fee-amount">
                                <span class="fee-value">{self.format_currency(data['total_ic'], currency)}</span>
                                <span class="fee-percentage">{data['avg_ic_pct']:.2f}%</span>
                            </div>
                        </div>
                        <div class="fee-bar">
                            <div class="fee-bar-fill fee-bar-ic" style="width: {ic_width:.1f}%"></div>
                        </div>
                    </div>

                    <div class="fee-item clickable" onclick="showModal('first', '{region_code}', '{card_type}')">
                        <div class="fee-header">
                            <div class="fee-label">
                                🏦 1st Plus (Scheme)
                                {' <span class="fee-badge">⚠️ MY Region</span>' if region_code == 'MY' else ''}
                                <span class="info-icon">i</span>
                            </div>
                            <div class="fee-amount">
                                <span class="fee-value">{self.format_currency(data['total_first_plus'], currency)}</span>
                                <span class="fee-percentage">{data['avg_first_plus_pct']:.2f}%</span>
                            </div>
                        </div>
                        <div class="fee-bar">
                            <div class="fee-bar-fill fee-bar-first" style="width: {first_width:.1f}%"></div>
                        </div>
                    </div>

                    <div class="fee-item clickable" onclick="showModal('second', '{region_code}', '{card_type}')">
                        <div class="fee-header">
                            <div class="fee-label">
                                🔧 2nd Plus (Acquirer)
                                <span class="info-icon">i</span>
                            </div>
                            <div class="fee-amount">
                                <span class="fee-value">{self.format_currency(data['total_second_plus'], currency)}</span>
                                <span class="fee-percentage">{data['avg_second_plus_pct']:.2f}%</span>
                            </div>
                        </div>
                        <div class="fee-bar">
                            <div class="fee-bar-fill fee-bar-second" style="width: {second_width:.1f}%"></div>
                        </div>

                        <div class="fee-tree">
'''

                # Add detailed breakdown (only non-zero items)
                volume = data['total_volume']

                def add_tree_item(label, amount, is_last=False):
                    if abs(amount) > 0.001:
                        pct = abs(amount / volume * 100) if volume > 0 else 0.0
                        icon = '└─' if is_last else '├─'
                        return f'''                            <div class="fee-tree-item">
                                <div class="fee-tree-label">
                                    <span class="fee-tree-icon">{icon}</span> {label}
                                </div>
                                <div class="fee-tree-amount">{self.format_currency(amount, currency)} ({pct:.3f}%)</div>
                            </div>
'''
                    return ''

                html += add_tree_item('Gateway Fee', data['total_gateway_fee'])
                html += add_tree_item('Authorization Fee', data['total_authorization_fee'])
                html += add_tree_item('Clearing Fee', data['total_clearing_fee'])
                html += add_tree_item('Cross-Border Fee', data['total_cross_border_fee'])
                html += add_tree_item('Cross-Currency Fee', data['total_cross_currency_fee'])
                html += add_tree_item('Preauth Fee', data['total_preauth_fee'])
                html += add_tree_item('3DS Fee', data['total_three_ds_fee'])
                html += add_tree_item('Non-3DS Fee', data['total_non_three_ds_fee'])
                html += add_tree_item('VAT', data['total_vat'])
                html += add_tree_item('WHT', data['total_wht'])
                html += add_tree_item('GRT', data['total_grt'])
                html += add_tree_item('ST', data['total_st'])
                html += add_tree_item('Net Acquirer Markup', data['total_net_acquirer_markup'], is_last=True)

                html += f'''                        </div>
                    </div>
                </div>

                <div class="total-section">
                    <div class="total-label">Total MDR</div>
                    <div class="total-value">{self.format_currency(data['total_mdr'], currency)} ({data['avg_mdr_pct']:.2f}%)</div>
                </div>
            </div>
'''

        html += '''        </div>

        <div class="footer">
            <p>🤖 Generated with IC++ Pricing Breakdown Calculator</p>
            <p>100% Real Data • Complete Transparency • Detailed Cost Analysis</p>
        </div>
    </div>

    <!-- Loading Overlay -->
    <div id="loadingOverlay" class="loading-overlay">
        <div class="loading-content">
            <div class="spinner"></div>
            <div class="loading-text">Loading CSV data...</div>
        </div>
    </div>

    <!-- Modal -->
    <div id="calculationModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h2 id="modalTitle">Calculation Details</h2>
                <span class="close" onclick="closeModal()">&times;</span>
            </div>
            <div class="modal-body" id="modalBody">
                <!-- Content will be inserted by JavaScript -->
            </div>
        </div>
    </div>

    <script>
        let csvData = null;

        // File upload handling
        const dropZone = document.getElementById('dropZone');
        const fileInput = document.getElementById('fileInput');
        const loadPresetBtn = document.getElementById('loadPresetBtn');
        const presetSelector = document.getElementById('presetSelector');

        // Show/hide loading indicator
        function showLoading() {
            document.getElementById('loadingOverlay').style.display = 'flex';
        }

        function hideLoading() {
            document.getElementById('loadingOverlay').style.display = 'none';
        }

        // Load preset button handler
        loadPresetBtn.addEventListener('click', () => {
            const selectedFile = presetSelector.value;
            if (!selectedFile) {
                alert('Please select a report from the dropdown first');
                return;
            }

            showLoading();

            // Try to fetch the file
            fetch(selectedFile)
                .then(response => {
                    if (!response.ok) throw new Error('Could not load preset');
                    return response.text();
                })
                .then(text => {
                    parseCSV(text, selectedFile);
                    hideLoading();
                })
                .catch(error => {
                    hideLoading();
                    alert('Could not auto-load preset. Please use drag & drop or click to browse instead.\\n\\nNote: Preset loading only works via local server (python3 serve.py)');
                });
        });

        // Highlight button on hover
        loadPresetBtn.addEventListener('mouseenter', () => {
            loadPresetBtn.style.transform = 'scale(1.02)';
        });
        loadPresetBtn.addEventListener('mouseleave', () => {
            loadPresetBtn.style.transform = 'scale(1)';
        });

        dropZone.addEventListener('click', () => fileInput.click());

        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#764ba2';
            dropZone.style.background = '#e8eaf6';
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.style.borderColor = '#667eea';
            dropZone.style.background = '#f8f9ff';
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = '#667eea';
            dropZone.style.background = '#f8f9ff';
            const file = e.dataTransfer.files[0];
            if (file && file.name.endsWith('.csv')) {
                processFile(file);
            } else {
                alert('Please upload a CSV file');
            }
        });

        fileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                processFile(file);
            }
        });

        function processFile(file) {
            showLoading();
            const reader = new FileReader();
            reader.onload = function(e) {
                const text = e.target.result;
                parseCSV(text, file.name);
                hideLoading();
            };
            reader.onerror = function() {
                hideLoading();
                alert('Error reading file');
            };
            reader.readAsText(file);
        }

        function parseCSV(text, fileName) {
            const lines = text.split('\\n');
            const headers = lines[0].split(',');

            const data = [];
            for (let i = 1; i < lines.length; i++) {
                if (lines[i].trim()) {
                    const values = lines[i].split(',');
                    const row = {};
                    headers.forEach((header, index) => {
                        row[header.trim()] = values[index] ? values[index].trim() : '';
                    });
                    data.push(row);
                }
            }

            csvData = data;

            // Show upload status
            document.getElementById('fileName').textContent = fileName;
            document.getElementById('fileStats').textContent = \`\${data.length} rows • \${headers.length} columns\`;
            document.getElementById('uploadStatus').style.display = 'block';

            // Process and update display
            processTransactions();
        }

        function processTransactions() {
            if (!csvData) return;

            // Filter and group transactions
            const stats = {};
            let totalProcessed = 0;
            let totalSkipped = 0;

            csvData.forEach(row => {
                const txnType = row['Transaction Type'] || '';
                const status = row['Transaction Status'] || '';
                const amount = parseFloat(row['Transaction Amount'] || 0);

                // Skip gateway events and non-approved transactions
                if (txnType === 'Gateway_Events' || status !== 'Approved' || amount === 0) {
                    totalSkipped++;
                    return;
                }

                // Identify region from account name
                const merchant = row['Account name'] || '';
                let region = 'UNKNOWN';
                if (merchant.includes('Hong Kong')) region = 'HK';
                else if (merchant.includes('Malaysia')) region = 'MY';
                else if (merchant.includes('Thailand')) region = 'TH';

                const cardType = row['Card Type'] || 'UNKNOWN';
                const currency = row['Transaction Currency'] || '';

                // Parse fees
                const mdr = parseFloat(row['MDR Amount'] || 0);
                const ic = parseFloat(row['Interchange Amount'] || 0);
                let schemeRaw = parseFloat(row['Scheme Fee Bucket Amount'] || 0);
                const scheme = region === 'MY' ? 0 : schemeRaw; // MY region special case

                const gateway = parseFloat(row['Gateway Fee Amount'] || 0);
                const auth = parseFloat(row['Authorization Amount'] || 0);
                const clearing = parseFloat(row['Clearing Amount'] || 0);
                const crossBorder = parseFloat(row['Cross Border Amount'] || 0);
                const crossCurrency = parseFloat(row['Cross Currency Amount'] || 0);
                const preauth = parseFloat(row['Preauthorization Amount'] || 0);
                const threeDS = parseFloat(row['Three Ds Amount'] || 0);
                const nonThreeDS = parseFloat(row['Non Three Ds Amount'] || 0);
                const vat = parseFloat(row['VAT Amount'] || 0);
                const wht = parseFloat(row['WHT Amount'] || 0);
                const grt = parseFloat(row['GRT Amount'] || 0);
                const st = parseFloat(row['ST Amount'] || 0);

                const secondPlus = mdr - ic - scheme;
                const knownComponents = gateway + auth + clearing + crossBorder + crossCurrency +
                                      preauth + threeDS + nonThreeDS + vat + wht + grt + st;
                const netMarkup = secondPlus - knownComponents;

                // Initialize region/card type if needed
                if (!stats[region]) stats[region] = {};
                if (!stats[region][cardType]) {
                    stats[region][cardType] = {
                        count: 0,
                        volume: 0,
                        ic: 0,
                        scheme: 0,
                        secondPlus: 0,
                        mdr: 0,
                        gateway: 0,
                        auth: 0,
                        clearing: 0,
                        crossBorder: 0,
                        crossCurrency: 0,
                        preauth: 0,
                        threeDS: 0,
                        nonThreeDS: 0,
                        vat: 0,
                        wht: 0,
                        grt: 0,
                        st: 0,
                        netMarkup: 0,
                        currency: currency
                    };
                }

                // Aggregate
                const bucket = stats[region][cardType];
                bucket.count++;
                bucket.volume += amount;
                bucket.ic += ic;
                bucket.scheme += scheme;
                bucket.secondPlus += secondPlus;
                bucket.mdr += mdr;
                bucket.gateway += gateway;
                bucket.auth += auth;
                bucket.clearing += clearing;
                bucket.crossBorder += crossBorder;
                bucket.crossCurrency += crossCurrency;
                bucket.preauth += preauth;
                bucket.threeDS += threeDS;
                bucket.nonThreeDS += nonThreeDS;
                bucket.vat += vat;
                bucket.wht += wht;
                bucket.grt += grt;
                bucket.st += st;
                bucket.netMarkup += netMarkup;

                totalProcessed++;
            });

            // Update display
            updateDisplay(stats, totalProcessed, totalSkipped);
        }

        function updateDisplay(stats, totalProcessed, totalSkipped) {
            // Update overview stats
            const totalRegions = Object.keys(stats).length;
            const totalInCSV = totalProcessed + totalSkipped;
            const coveragePct = totalInCSV > 0 ? ((totalProcessed / totalInCSV) * 100).toFixed(0) : 0;

            document.querySelector('.stats-overview').innerHTML = \`
                <div class="stat-card">
                    <h3>Transactions Analyzed</h3>
                    <div class="value">\${totalProcessed}</div>
                </div>
                <div class="stat-card">
                    <h3>Regions Found</h3>
                    <div class="value">\${totalRegions}</div>
                </div>
                <div class="stat-card">
                    <h3>CSV Coverage</h3>
                    <div class="value">\${coveragePct}%</div>
                </div>
                <div class="stat-card">
                    <h3>Skipped</h3>
                    <div class="value" style="color: \${totalSkipped > totalProcessed ? '#f5576c' : '#43e97b'};">\${totalSkipped}</div>
                </div>
            \`;

            // Generate region cards
            const regionsGrid = document.querySelector('.regions-grid');
            regionsGrid.innerHTML = '';

            const regionMeta = {
                'HK': { name: 'Hong Kong', flag: '🇭🇰', symbol: 'HK$' },
                'MY': { name: 'Malaysia', flag: '🇲🇾', symbol: 'RM' },
                'TH': { name: 'Thailand', flag: '🇹🇭', symbol: '฿' },
                'UNKNOWN': { name: 'Other', flag: '🌐', symbol: '$' }
            };

            Object.keys(stats).sort().forEach(region => {
                Object.keys(stats[region]).sort().forEach(cardType => {
                    const data = stats[region][cardType];
                    const meta = regionMeta[region] || regionMeta['UNKNOWN'];

                    const icPct = data.volume ? Math.abs(data.ic / data.volume * 100) : 0;
                    const schemePct = data.volume ? Math.abs(data.scheme / data.volume * 100) : 0;
                    const secondPct = data.volume ? Math.abs(data.secondPlus / data.volume * 100) : 0;
                    const mdrPct = data.volume ? Math.abs(data.mdr / data.volume * 100) : 0;

                    const mdrAbs = Math.abs(data.mdr);
                    const icWidth = mdrAbs ? Math.abs(data.ic) / mdrAbs * 100 : 0;
                    const schemeWidth = mdrAbs ? Math.abs(data.scheme) / mdrAbs * 100 : 0;
                    const secondWidth = mdrAbs ? Math.abs(data.secondPlus) / mdrAbs * 100 : 0;

                    const card = document.createElement('div');
                    card.className = 'region-card';
                    card.innerHTML = \`
                        <div class="region-header">
                            <div class="region-title">
                                <span class="region-flag">\${meta.flag}</span>
                                <span class="region-name">\${meta.name}</span>
                            </div>
                            <div class="region-code">\${region}</div>
                        </div>

                        <div class="region-stats">
                            <div class="region-stat">
                                <div class="region-stat-label">Volume</div>
                                <div class="region-stat-value">\${meta.symbol}\${Math.abs(data.volume).toFixed(2)}</div>
                            </div>
                            <div class="region-stat">
                                <div class="region-stat-label">Transactions</div>
                                <div class="region-stat-value">\${data.count}</div>
                            </div>
                            <div class="region-stat">
                                <div class="region-stat-label">Card Type</div>
                                <div class="region-stat-value">\${cardType}</div>
                            </div>
                        </div>

                        <div class="fee-breakdown">
                            <div class="fee-item clickable" onclick="showModal('ic', '\${region}', '\${cardType}')">
                                <div class="fee-header">
                                    <div class="fee-label">
                                        💱 IC (Interchange)
                                        <span class="info-icon">i</span>
                                    </div>
                                    <div class="fee-amount">
                                        <span class="fee-value">\${meta.symbol}\${Math.abs(data.ic).toFixed(2)}</span>
                                        <span class="fee-percentage">\${icPct.toFixed(2)}%</span>
                                    </div>
                                </div>
                                <div class="fee-bar">
                                    <div class="fee-bar-fill fee-bar-ic" style="width: \${icWidth.toFixed(1)}%"></div>
                                </div>
                            </div>

                            <div class="fee-item clickable" onclick="showModal('first', '\${region}', '\${cardType}')">
                                <div class="fee-header">
                                    <div class="fee-label">
                                        🏦 1st Plus (Scheme)
                                        \${region === 'MY' ? '<span class="fee-badge">⚠️ MY Region</span>' : ''}
                                        <span class="info-icon">i</span>
                                    </div>
                                    <div class="fee-amount">
                                        <span class="fee-value">\${meta.symbol}\${Math.abs(data.scheme).toFixed(2)}</span>
                                        <span class="fee-percentage">\${schemePct.toFixed(2)}%</span>
                                    </div>
                                </div>
                                <div class="fee-bar">
                                    <div class="fee-bar-fill fee-bar-first" style="width: \${schemeWidth.toFixed(1)}%"></div>
                                </div>
                            </div>

                            <div class="fee-item clickable" onclick="showModal('second', '\${region}', '\${cardType}')">
                                <div class="fee-header">
                                    <div class="fee-label">
                                        🔧 2nd Plus (Acquirer)
                                        <span class="info-icon">i</span>
                                    </div>
                                    <div class="fee-amount">
                                        <span class="fee-value">\${meta.symbol}\${Math.abs(data.secondPlus).toFixed(2)}</span>
                                        <span class="fee-percentage">\${secondPct.toFixed(2)}%</span>
                                    </div>
                                </div>
                                <div class="fee-bar">
                                    <div class="fee-bar-fill fee-bar-second" style="width: \${secondWidth.toFixed(1)}%"></div>
                                </div>

                                <div class="fee-tree">
                                    \${generateTreeItem('Gateway Fee', data.gateway, data.volume, meta.symbol)}
                                    \${generateTreeItem('Authorization Fee', data.auth, data.volume, meta.symbol)}
                                    \${generateTreeItem('Clearing Fee', data.clearing, data.volume, meta.symbol)}
                                    \${generateTreeItem('Cross-Border Fee', data.crossBorder, data.volume, meta.symbol)}
                                    \${generateTreeItem('Cross-Currency Fee', data.crossCurrency, data.volume, meta.symbol)}
                                    \${generateTreeItem('Preauth Fee', data.preauth, data.volume, meta.symbol)}
                                    \${generateTreeItem('3DS Fee', data.threeDS, data.volume, meta.symbol)}
                                    \${generateTreeItem('Non-3DS Fee', data.nonThreeDS, data.volume, meta.symbol)}
                                    \${generateTreeItem('VAT', data.vat, data.volume, meta.symbol)}
                                    \${generateTreeItem('WHT', data.wht, data.volume, meta.symbol)}
                                    \${generateTreeItem('GRT', data.grt, data.volume, meta.symbol)}
                                    \${generateTreeItem('ST', data.st, data.volume, meta.symbol)}
                                    \${generateTreeItem('Net Acquirer Markup', data.netMarkup, data.volume, meta.symbol, true)}
                                </div>
                            </div>
                        </div>

                        <div class="total-section">
                            <div class="total-label">Total MDR</div>
                            <div class="total-value">\${meta.symbol}\${Math.abs(data.mdr).toFixed(2)} (\${mdrPct.toFixed(2)}%)</div>
                        </div>
                    \`;
                    regionsGrid.appendChild(card);
                });
            });

            // Animate bars
            setTimeout(() => {
                const bars = document.querySelectorAll('.fee-bar-fill');
                bars.forEach(bar => {
                    const width = bar.style.width;
                    bar.style.width = '0%';
                    setTimeout(() => {
                        bar.style.width = width;
                    }, 100);
                });
            }, 100);
        }

        function generateTreeItem(label, amount, volume, symbol, isLast = false) {
            if (Math.abs(amount) < 0.001) return '';
            const pct = volume ? Math.abs(amount / volume * 100) : 0;
            const icon = isLast ? '└─' : '├─';
            return \`
                <div class="fee-tree-item">
                    <div class="fee-tree-label">
                        <span class="fee-tree-icon">\${icon}</span> \${label}
                    </div>
                    <div class="fee-tree-amount">\${symbol}\${Math.abs(amount).toFixed(2)} (\${pct.toFixed(3)}%)</div>
                </div>
            \`;
        }

        // Animation for bars (initial load)
        window.addEventListener('load', function() {
            const bars = document.querySelectorAll('.fee-bar-fill');
            bars.forEach(bar => {
                const width = bar.style.width;
                bar.style.width = '0%';
                setTimeout(() => {
                    bar.style.width = width;
                }, 100);
            });
        });

        // Modal functions
        function showModal(feeType, region, cardType) {
            const modal = document.getElementById('calculationModal');
            const modalTitle = document.getElementById('modalTitle');
            const modalBody = document.getElementById('modalBody');

            let content = '';
            let title = '';

            if (feeType === 'ic') {
                title = '💱 IC (Interchange Fee) Calculation';
                content = `
                    <div class="calculation-section">
                        <h3>📊 Data Source</h3>
                        <p>This value comes from the NomuPay Daily Funding Report CSV file.</p>
                        <div class="calc-row">
                            <span class="calc-label">CSV Column:</span>
                            <span class="calc-value"><span class="column-tag">Interchange Amount</span></span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Region Filter:</span>
                            <span class="calc-value">${region}</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Card Type Filter:</span>
                            <span class="calc-value">${cardType}</span>
                        </div>
                    </div>

                    <div class="calculation-section">
                        <h3>🔢 Calculation Method</h3>
                        <div class="formula-box">
                            IC Total = SUM(Interchange Amount)<br>
                            WHERE Merchant contains "${region}" AND Card Type = "${cardType}"
                        </div>
                        <p><strong>What is IC (Interchange)?</strong></p>
                        <p>The Interchange Fee is paid to the card-issuing bank. It varies based on:</p>
                        <ul>
                            <li>Card type (Credit vs Debit)</li>
                            <li>Card brand (Visa, Mastercard, etc.)</li>
                            <li>Transaction type (ECOM, POS, etc.)</li>
                            <li>Issuer country vs Merchant country</li>
                        </ul>
                    </div>

                    <div class="calculation-section">
                        <h3>📁 CSV File Structure</h3>
                        <p>The script reads from:</p>
                        <span class="column-tag">daily_funding_report_*.csv</span>
                        <p style="margin-top: 10px;">Column 38 in the CSV contains the Interchange Amount per transaction.</p>
                    </div>
                `;
            } else if (feeType === 'first') {
                title = '🏦 1st Plus (Scheme Fee) Calculation';
                const isMyRegion = region === 'MY';
                content = `
                    <div class="calculation-section">
                        <h3>📊 Data Source</h3>
                        <p>This value comes from the NomuPay Daily Funding Report CSV file.</p>
                        <div class="calc-row">
                            <span class="calc-label">CSV Column:</span>
                            <span class="calc-value"><span class="column-tag">Scheme Fee Bucket Amount</span></span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Region Filter:</span>
                            <span class="calc-value">${region} ${isMyRegion ? '⚠️' : ''}</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Card Type Filter:</span>
                            <span class="calc-value">${cardType}</span>
                        </div>
                    </div>

                    ${isMyRegion ? `
                    <div class="calculation-section">
                        <h3>⚠️ Malaysia Special Case</h3>
                        <div class="formula-box" style="border-color: #ffc107; background: #fff3cd;">
                            IF Region = "MY" THEN Scheme Fee = 0<br>
                            <br>
                            Malaysia merchants do NOT pay scheme fees.<br>
                            This value is ALWAYS forced to 0 regardless of CSV data.
                        </div>
                    </div>
                    ` : `
                    <div class="calculation-section">
                        <h3>🔢 Calculation Method</h3>
                        <div class="formula-box">
                            1st Plus Total = SUM(Scheme Fee Bucket Amount)<br>
                            WHERE Merchant contains "${region}" AND Card Type = "${cardType}"
                        </div>
                    </div>
                    `}

                    <div class="calculation-section">
                        <h3>💡 What is 1st Plus (Scheme Fee)?</h3>
                        <p>The Scheme Fee is charged by the card networks (Visa, Mastercard, etc.) for:</p>
                        <ul>
                            <li>Processing transactions through their network</li>
                            <li>Authorization and clearing services</li>
                            <li>Network maintenance and security</li>
                            <li>Cross-border transaction processing</li>
                        </ul>
                        ${isMyRegion ? '<p><strong>Note:</strong> Malaysia has negotiated zero scheme fees with card networks.</p>' : ''}
                    </div>

                    <div class="calculation-section">
                        <h3>📁 CSV File Structure</h3>
                        <p>Column 50 in the CSV: <span class="column-tag">Scheme Fee Bucket Amount</span></p>
                    </div>
                `;
            } else if (feeType === 'second') {
                title = '🔧 2nd Plus (Acquirer Markup) Calculation';
                content = `
                    <div class="calculation-section">
                        <h3>📊 Data Source</h3>
                        <p>This is a <strong>calculated value</strong> derived from multiple CSV columns.</p>
                    </div>

                    <div class="calculation-section">
                        <h3>🔢 Calculation Formula</h3>
                        <div class="formula-box">
                            2nd Plus = MDR - IC - 1st Plus
                        </div>
                        <p>Where:</p>
                        <div class="calc-row">
                            <span class="calc-label">MDR:</span>
                            <span class="calc-value"><span class="column-tag">MDR Amount</span> (Column 36)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">IC:</span>
                            <span class="calc-value"><span class="column-tag">Interchange Amount</span> (Column 38)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">1st Plus:</span>
                            <span class="calc-value"><span class="column-tag">Scheme Fee Bucket Amount</span> (Column 50)</span>
                        </div>
                    </div>

                    <div class="calculation-section">
                        <h3>🔧 2nd Plus Components Breakdown</h3>
                        <p>The 2nd Plus is further broken down into these CSV columns:</p>
                        <div class="calc-row">
                            <span class="calc-label">Gateway Fee:</span>
                            <span class="calc-value"><span class="column-tag">Gateway Fee Amount</span> (Col 56)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Authorization Fee:</span>
                            <span class="calc-value"><span class="column-tag">Authorization Amount</span> (Col 44)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Clearing Fee:</span>
                            <span class="calc-value"><span class="column-tag">Clearing Amount</span> (Col 46)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Cross-Border Fee:</span>
                            <span class="calc-value"><span class="column-tag">Cross Border Amount</span> (Col 40)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">Cross-Currency Fee:</span>
                            <span class="calc-value"><span class="column-tag">Cross Currency Amount</span> (Col 42)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">3DS Fee:</span>
                            <span class="calc-value"><span class="column-tag">Three Ds Amount</span> (Col 54)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">VAT:</span>
                            <span class="calc-value"><span class="column-tag">VAT Amount</span> (Col 60)</span>
                        </div>
                        <div class="calc-row">
                            <span class="calc-label">WHT:</span>
                            <span class="calc-value"><span class="column-tag">WHT Amount</span> (Col 64)</span>
                        </div>
                    </div>

                    <div class="calculation-section">
                        <h3>💰 Net Acquirer Markup</h3>
                        <div class="formula-box">
                            Net Markup = 2nd Plus - (Gateway + Auth + Clearing + ... + Taxes)
                        </div>
                        <p>This is the <strong>pure profit margin</strong> retained by the acquirer after all operational costs.</p>
                    </div>

                    <div class="calculation-section">
                        <h3>💡 What is 2nd Plus (Acquirer Markup)?</h3>
                        <p>The Acquirer Markup covers:</p>
                        <ul>
                            <li>Gateway processing fees (3DS, Capture, Debit, Refund events)</li>
                            <li>Authorization and clearing operational costs</li>
                            <li>Risk management and fraud prevention</li>
                            <li>Cross-border and currency conversion handling</li>
                            <li>Taxes (VAT, WHT, GRT, ST)</li>
                            <li>Acquirer profit margin</li>
                        </ul>
                    </div>
                `;
            }

            modalTitle.textContent = title;
            modalBody.innerHTML = content;
            modal.style.display = 'block';
        }

        function closeModal() {
            document.getElementById('calculationModal').style.display = 'none';
        }

        // Close modal when clicking outside
        window.onclick = function(event) {
            const modal = document.getElementById('calculationModal');
            if (event.target == modal) {
                closeModal();
            }
        }
    </script>
</body>
</html>
'''

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html)

        print(f"HTML report complete")


# ============================================================================
# MAIN
# ============================================================================

def main():
    """Main execution flow"""

    # Check command line arguments
    if len(sys.argv) < 3:
        print("Usage: python3 icpp_calculator.py <excel_file> <fee_csv_file>")
        print("\nExample:")
        print('  python3 icpp_calculator.py "Merchant Funding Transactions 2025-12-15,2026-01-08.xlsx" fees_export.csv')
        sys.exit(1)

    excel_file = sys.argv[1]
    fee_csv_file = sys.argv[2]
    output_csv = 'icpp_breakdown_report.csv'

    print("IC++ PRICING BREAKDOWN CALCULATOR")
    print("=" * 65)

    # Step 1: Read Excel file
    try:
        excel_reader = ExcelReader(excel_file)
        transactions = excel_reader.read()
    except Exception as e:
        print(f"ERROR: Failed to read Excel file: {e}")
        sys.exit(1)

    # Step 2: Load fee data
    try:
        fee_loader = FeeDataLoader(fee_csv_file)
        fee_data = fee_loader.load()
    except Exception as e:
        print(f"ERROR: Failed to load fee data: {e}")
        sys.exit(1)

    # Step 3: Process transactions
    print("\nProcessing transactions...")
    aggregator = StatisticsAggregator()

    for transaction in transactions:
        # Get transaction ID - Try Transaction ID first, then Gateway UUID
        # IMPORTANT: NomuPay CSVs use "NP Transaction ID" which matches Excel's "Transaction ID" column
        txn_id = (transaction.get('Transaction ID') or
                 transaction.get('Gateway UUID') or
                 transaction.get('Gateway Reference'))

        # Validate transaction
        is_valid, reason = is_valid_transaction(transaction)
        if not is_valid:
            aggregator.skip_transaction(reason)
            continue

        # Check if we have fee data
        if txn_id not in fee_data:
            aggregator.skip_transaction('Missing fee data')
            continue

        fees = fee_data[txn_id]

        # Skip if no MDR data
        if fees['mdr_amount'] == 0.0:
            aggregator.skip_transaction('Zero MDR')
            continue

        # Identify region
        merchant = transaction.get('Merchant', '')
        card_country = transaction.get('Card Country', '')
        region = identify_region(merchant, card_country)

        # Calculate IC++ breakdown
        icpp = calculate_icpp(transaction, fees, region)

        # Add to statistics
        aggregator.add_transaction(transaction, icpp, region)

    # Step 4: Calculate percentages
    aggregator.calculate_percentages()
    stats = aggregator.get_stats()

    # Step 5: Generate reports
    output_html = 'icpp_breakdown_report.html'

    report_gen = ReportGenerator()
    report_gen.print_console_report(stats, aggregator)
    report_gen.export_csv(stats, output_csv)
    report_gen.export_html(stats, aggregator, output_html)

    print(f"\n✓ Analysis complete!")
    print(f"✓ CSV report saved to: {output_csv}")
    print(f"✓ HTML report saved to: {output_html}")
    print(f"\nOpen the HTML report in your browser to view the visual breakdown.")


if __name__ == '__main__':
    main()
