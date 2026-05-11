from dashboard import calculate_salary
from datetime import date

result = calculate_salary(date(2026, 4, 1), date(2026, 4, 30))
if 'error' in result:
    print('ERROR:', result['error'])
else:
    print('Total orders:', result['total_orders'])
    print('Employees:')
    for name, emp in result['employees'].items():
        print(f"  {name}: total={emp['total']}, orders={emp['orders_count']}, done={emp['orders_done']}")
