# Financial Report Generator

Generates financial reports from HTML bank statements with Russian Excel format.

## Installation

### Local Installation
```bash
pip install beautifulsoup4 openpyxl
```

### Docker Installation
```bash
docker-compose up --build
```

## Usage

### Local Usage
1. Place HTML statement as `statement_example.html`
2. (Optional) Create `category_mapping.csv` for categorization
3. Run:
```bash
python financial_report.py
```

### Docker Usage
1. Place files in `data/` directory:
   - `statement_example.html`
   - `category_mapping.csv` (optional)
2. Run:
```bash
docker-compose up
```

## Environment Variables

Customize budget values via environment variables:
- `FIN_INCOME_VALUE` (default: 6453)
- `FIN_TAX_VALUE` (default: 1652)
- `FIN_FOOD_VALUE` (default: 900)
- `FIN_UTILITY_VALUE` (default: 1333.88)
- `FIN_SAVINGS_VALUE` (default: 700)

## Files

- `financial_report.py` - Main script
- `statement_example.html` - Input HTML file
- `category_mapping.csv` - Optional mapping (Description,Category)
- `financial_report.xlsx` - Output Excel file
- `processed_periods` - Tracks processed date ranges

## Category Mapping Example

```csv
Description,Category
PEKARNA,Food
KONZUM,Food
A1 HRVATSKA,Utility bills
```

## Output

Creates monthly Excel sheets with:
- Fixed budget values (Income: 6453, Taxes: 1652, Food: 900, Utilities: 1333.88, Savings: 700)
- Formula-based calculations for remaining amounts
- C3: Food remaining (900 - spent)
- D3: Utility bills remaining (1333.88 - spent)
- A7: Other remaining (budget - spent)
- Automatic backup on update
- Period overlap detection
- Detailed log file (logs_{end_date}.log) with transaction processing details