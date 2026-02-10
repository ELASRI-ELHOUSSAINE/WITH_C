# WITH_C Repository

This repository contains multiple projects:

## 📊 Sales & Inventory Dashboard (Excel)

**NEW!** Professional Excel dashboard for business analytics and inventory management.

### Quick Start
1. Open `Sales_Inventory_Dashboard.xlsx` in Microsoft Excel 2021
2. View the Dashboard sheet for KPIs and charts
3. Explore Products and Sales sheets for sample data
4. Read `DASHBOARD_README.md` for detailed documentation
5. See `QUICK_START_GUIDE.md` for 5-minute setup

### Features
- ✓ **5 KPI Cards**: Total Sales, Profit, Orders, Avg Order Value, Low Stock
- ✓ **Interactive Charts**: Sales trends, Top products, Stock status
- ✓ **Auto-calculations**: Current stock, profit margins, revenue
- ✓ **Excel 2021 Compatible**: No VBA, pure formulas & pivot tables
- ✓ **50 Sample Products**: Pre-loaded with realistic business data
- ✓ **200 Sales Transactions**: Full year of sample data

### Files
- `Sales_Inventory_Dashboard.xlsx` - Ready-to-use Excel dashboard
- `DASHBOARD_README.md` - Complete documentation
- `QUICK_START_GUIDE.md` - 5-minute setup guide
- `create_dashboard.py` - Python script to regenerate dashboard

---

## 🔧 C Stack Implementation (Linked List)

Simple stack (pile) implementation in C using a linked list.

### Features

- Create a stack
- Push / Pop
- Peek (top)
- Size and empty checks
- Print and free the stack

### Project Structure

- file.h: Stack types and function declarations
- file.c: Stack implementation
- main.c: Example usage

### Build

Using GCC:

```bash
gcc -Wall -Wextra -O2 -o main main.c file.c
```

### Run

```bash
./main
```

### Notes

- The stack is LIFO.
- On empty stack, pop/top exit with an error message.
