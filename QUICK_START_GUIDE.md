# 🚀 Quick Start Guide - Sales & Inventory Dashboard

## 📖 5-Minute Setup

### Step 1: Open the Dashboard
```
1. Double-click: Sales_Inventory_Dashboard.xlsx
2. Click: Enable Content (if prompted)
3. Navigate to: Dashboard sheet
```

### Step 2: Explore the Data
```
✓ Products sheet - 50 sample products
✓ Sales sheet - 200 sample transactions
✓ Calculations sheet - Auto-calculated metrics
✓ Dashboard sheet - Visual overview
```

### Step 3: Add Interactivity (Optional)

#### Quick Slicer Setup:
```
1. Click any cell in Products sheet (in the table)
2. Table Design → Insert Slicer
3. Select: Category ☑ Product Name ☑ Supplier ☑
4. Click OK
5. Drag slicers to Dashboard sheet
```

#### Timeline Slicer (Date Filter):
```
1. Click any cell in Sales sheet
2. Table Design → Insert Timeline
3. Select: Date
4. Drag to Dashboard sheet
```

## 🎯 Common Tasks

### Task 1: Add New Product
```
1. Go to Products sheet
2. Scroll to last row of table
3. Press Tab (table expands automatically)
4. Enter product details
5. Dashboard updates automatically
```

### Task 2: Record New Sale
```
1. Go to Sales sheet
2. Add new row at bottom
3. Enter: Order ID, Product ID, Quantity, Price, Date
4. Total Sale calculates automatically
5. All charts update
```

### Task 3: Check Stock Levels
```
1. Go to Calculations sheet
2. Look at "Current Stock" column (E)
3. Red = Low stock (< 50)
4. Yellow = Medium (50-100)
5. Green = Good stock (> 100)
```

### Task 4: Export to PDF
```
1. Go to Dashboard sheet
2. File → Print → Save as PDF
3. Adjust print area if needed
4. Click Save
```

## 📊 Understanding the KPIs

| KPI | What It Means | How It's Calculated |
|-----|---------------|---------------------|
| **Total Sales** | Revenue from all orders | Sum of Total Sale column |
| **Total Profit** | Net profit after costs | (Selling Price - Purchase Price) × Qty |
| **Total Orders** | Number of transactions | Count of Order IDs |
| **Avg Order Value** | Typical order size | Total Sales ÷ Total Orders |
| **Low Stock Products** | Items needing reorder | Products with stock < 50 |

## 🔍 Analysis Tips

### Monthly Trends
- Check "Sales by Month" line chart
- Look for seasonal patterns
- Identify growth/decline periods

### Best Sellers
- Review "Top 10 Selling Products" bar chart
- Focus inventory on these items
- Consider promotions for slow movers

### Inventory Management
- Monitor Calculations sheet weekly
- Reorder when stock turns yellow/red
- Adjust Initial Quantity as needed

## ⚡ Power User Tips

### Tip 1: Custom Pivot Table
Create a Category Performance analysis:
```
1. Insert → PivotTable from Sales
2. Rows: Category
3. Values: Total Sale (Sum)
4. Place on Dashboard sheet
```

### Tip 2: Conditional Formatting
Highlight top performers:
```
1. Select data range
2. Home → Conditional Formatting
3. Top 10 Items
4. Choose green fill
```

### Tip 3: Multiple Slicers
Connect one slicer to both tables:
```
1. Right-click slicer
2. Report Connections
3. Check: tbl_Products ☑ tbl_Sales ☑
4. Click OK
```

## ⚠️ Important Notes

### DO:
✓ Keep Product IDs consistent between tables
✓ Use proper date formats (MM/DD/YYYY)
✓ Backup before major changes
✓ Save regularly

### DON'T:
✗ Delete table headers
✗ Convert tables to ranges
✗ Break table structure
✗ Change table names without updating formulas

## 🐛 Troubleshooting

### Problem: Charts not showing data
**Solution**: Right-click chart → Select Data → Verify range

### Problem: #REF! errors
**Solution**: Check Product IDs match in both tables

### Problem: Slicer not filtering
**Solution**: Right-click → Report Connections → Select tables

### Problem: KPIs showing zero
**Solution**: Ensure data is in correct table format

## 📞 Need More Help?

1. **In-Excel Help**: Press F1
2. **Documentation**: See DASHBOARD_README.md
3. **Regenerate**: Run create_dashboard.py script
4. **Microsoft Support**: support.microsoft.com/excel

## ✅ Success Checklist

Before going live with your data:

- [ ] Backed up the template file
- [ ] Deleted sample data
- [ ] Added your real products
- [ ] Added your real sales
- [ ] Verified calculations are correct
- [ ] Tested slicers work
- [ ] Customized colors/branding
- [ ] Added company logo (optional)
- [ ] Trained team on usage
- [ ] Set up regular update schedule

---

**You're ready to go!** 🎉

Start by exploring the sample data, then replace it with your own when comfortable.
