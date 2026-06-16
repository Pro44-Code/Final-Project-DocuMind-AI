from datetime import datetime
import openpyxl

def check_expiry_from_excel(filepath, days_threshold=30):
    alerts = []
    today = datetime.now()

    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        for sheet in wb.worksheets:
            headers = []
            for row in sheet.iter_rows(values_only=True):
                if not headers:
                    headers = [str(h).lower() if h else "" for h in row]
                    continue

                if not any(row):
                    continue

                row_dict = dict(zip(headers, row))
                deadline = row_dict.get('deadlines')

                if not deadline:
                    continue

                # datetime ობიექტია პირდაპირ
                if isinstance(deadline, datetime):
                    days_left = (deadline - today).days
                    if 0 <= days_left <= days_threshold:
                        alerts.append({
                            'filename': filepath.split('\\')[-1],
                            'employee': row_dict.get('employee', 'უცნობი'),
                            'date': deadline.strftime('%Y-%m-%d'),
                            'days_left': days_left
                        })

    except Exception as e:
        print(f"Error: {e}")

    return alerts