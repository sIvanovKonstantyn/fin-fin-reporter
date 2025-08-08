from bs4 import BeautifulSoup
import re
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font
import os
import shutil
from datetime import datetime

def extract_dates(html_file):
    with open(html_file, 'r', encoding='windows-1250') as f:
        content = f.read()
    match = re.search(r'Za razdoblje \(po datumu valute\): (\d{2}\.\d{2}\.\d{4})\. do (\d{2}\.\d{2}\.\d{4})\.', content)
    return (match.group(1), match.group(2)) if match else (None, None)

def check_period_overlap(start_date, end_date, log_file):
    periods_file = 'processed_periods'
    if os.path.exists(periods_file):
        with open(periods_file, 'r') as f:
            for line in f:
                if line.strip():
                    existing_start, existing_end = line.strip().split(',')
                    if start_date <= existing_end and end_date >= existing_start:
                        with open(log_file, 'a', encoding='utf-8') as lf:
                            lf.write(f"ERROR: Period {start_date}-{end_date} overlaps with {existing_start}-{existing_end}\n")
                        return True
    
    with open(periods_file, 'a') as f:
        f.write(f"{start_date},{end_date}\n")
    return False

def log_transaction(log_file, description, amount, category, skipped=False):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    status = "SKIPPED" if skipped else "PROCESSED"
    with open(log_file, 'a', encoding='utf-8') as f:
        f.write(f"{timestamp} | {description} | {amount} | {category} | {status}\n")

def parse_transactions(html_file, log_file):
    with open(html_file, 'r', encoding='windows-1250') as f:
        content = f.read()
    
    soup = BeautifulSoup(content, 'html.parser')
    table = soup.find('table', {'style': lambda x: x and 'border-collapse: collapse' in x})
    transactions = []
    
    with open(log_file, 'w', encoding='utf-8') as f:
        f.write("Timestamp | Description | Amount | Category | Status\n")
        f.write("-" * 80 + "\n")
    
    for row in table.find_all('tr'):
        cells = row.find_all('td')
        if len(cells) >= 6 and not row.get('bgcolor'):
            date_text = cells[0].get_text(strip=True)
            if re.match(r'\d{2}\.\d{2}\.\d{4}', date_text):
                description = cells[1].get_text(strip=True).split('\n')[0].strip()
                amount_text = cells[4].get_text(strip=True)
                if amount_text and amount_text != '&nbsp;':
                    try:
                        amount = float(amount_text.replace(',', '.').replace(' ', ''))
                        transactions.append({'description': description, 'amount': amount})
                    except ValueError:
                        log_transaction(log_file, description, amount_text, "N/A", skipped=True)
                else:
                    log_transaction(log_file, description, "0.00", "N/A", skipped=True)
    return transactions

def load_mapping():
    if os.path.exists('category_mapping.csv'):
        mapping = {}
        with open('category_mapping.csv', 'r', encoding='utf-8') as f:
            lines = f.readlines()[1:]
            for line in lines:
                parts = line.strip().split(',')
                if len(parts) >= 2:
                    mapping[parts[0]] = parts[1]
        return mapping
    return {}

def categorize(description, mapping):
    for keyword, category in mapping.items():
        if keyword.lower() in description.lower():
            return category
    return 'Other'

def save_excel(totals, end_date, output_file):
    if os.path.exists(output_file):
        backup_date = datetime.now().strftime("%d_%m_%Y")
        shutil.copy2(output_file, f"{output_file}.backup_{backup_date}")
        wb = load_workbook(output_file)
    else:
        wb = Workbook()
        wb.remove(wb.active)
    
    if end_date in wb.sheetnames:
        ws = wb[end_date]
    else:
        ws = wb.create_sheet(end_date)
    
    bold = Font(bold=True)
    
    # Check for manual expenses in C6-C100 and add to other category
    manual_expenses = 0
    for row in range(6, 101):
        cell_value = ws[f'C{row}'].value
        if cell_value and isinstance(cell_value, (int, float)):
            manual_expenses += cell_value
            ws[f'C{row}'] = None  # Clear the cell
    
    # Add manual expenses to other category
    other_total = totals.get('Other', 0) + manual_expenses
    
    # Set headers and fixed values from environment variables
    ws['A1'] = 'Доход'
    ws['A1'].font = bold
    if not ws['A2'].value:
        ws['A2'] = float(os.getenv('FIN_INCOME_VALUE', '6453'))
    
    ws['B1'] = 'Налоги'
    ws['B1'].font = bold
    if not ws['B2'].value:
        ws['B2'] = float(os.getenv('FIN_TAX_VALUE', '1652'))
    
    ws['C1'] = 'Еда'
    ws['C1'].font = bold
    if not ws['C2'].value:
        ws['C2'] = float(os.getenv('FIN_FOOD_VALUE', '900'))
    
    ws['D1'] = 'Комы'
    ws['D1'].font = bold
    if not ws['D2'].value:
        ws['D2'] = float(os.getenv('FIN_UTILITY_VALUE', '1333.88'))
    
    ws['E1'] = 'Отложить'
    ws['E1'].font = bold
    if not ws['E2'].value:
        ws['E2'] = float(os.getenv('FIN_SAVINGS_VALUE', '700'))
    
    ws['A5'] = 'Бюджет'
    ws['A5'].font = bold
    if not ws['A6'].value:
        ws['A6'] = f'=A2-B2-C2-D2-E2'
    
    ws['C5'] = 'Наличка'
    ws['C5'].font = bold
    
    ws['A12'] = 'Заполнено по'
    ws['A12'].font = bold
    ws['A13'] = end_date
    
    # Update category totals with formulas
    food_spent = totals.get('Food', 0)
    utility_spent = totals.get('Utility bills', 0)
    
    if food_spent > 0:
        current_c3 = ws['C3'].value or 0
        ws['C3'] = f'=C2-{current_c3}-{food_spent}'
    
    if utility_spent > 0:
        current_d3 = ws['D3'].value or 0
        ws['D3'] = f'=D2-{current_d3}-{utility_spent}'
    
    if other_total > 0:
        current_a7 = ws['A7'].value or 0
        ws['A7'] = f'=A6-{current_a7}-{other_total}'
    
    wb.save(output_file)

def main():
    # Check if running in Docker (data directory exists)
    if os.path.exists('/app/data'):
        os.chdir('/app/data')
    
    html_file = 'statement_example.html'
    output_file = 'financial_report.xlsx'
    
    if not os.path.exists(html_file):
        print(f"Error: {html_file} not found")
        return
    
    start_date, end_date = extract_dates(html_file)
    if not start_date or not end_date:
        print("Could not extract dates")
        return
    
    log_file = f"logs_{end_date.replace('.', '_')}.log"
    
    if check_period_overlap(start_date, end_date, log_file):
        print(f"Period overlap detected. Check {log_file}")
        return
    
    transactions = parse_transactions(html_file, log_file)
    mapping = load_mapping()
    
    totals = {}
    for t in transactions:
        category = categorize(t['description'], mapping)
        totals[category] = totals.get(category, 0) + t['amount']
        log_transaction(log_file, t['description'], f"€{t['amount']:.2f}", category)
    
    save_excel(totals, end_date, output_file)
    print(f"Processed {len(transactions)} transactions")
    for category, total in totals.items():
        print(f"{category}: €{total:.2f}")
    print(f"Files saved in: {os.getcwd()}")

if __name__ == "__main__":
    main()