# 📊 Sales & Inventory Dashboard

## Overview
This is a professional, fully functional Sales & Inventory Dashboard created for Microsoft Excel 2021. The dashboard provides comprehensive insights into sales performance, inventory levels, and profitability metrics.

## ✨ Features

### 📁 Data Structure
The Excel file contains the following sheets:

1. **Products Sheet** (`tbl_Products`)
   - Product ID
   - Product Name
   - Category (Electronics, Clothing, Home & Garden, Sports, Books)
   - Purchase Price
   - Selling Price
   - Initial Quantity
   - Supplier
   - Date Added

2. **Sales Sheet** (`tbl_Sales`)
   - Order ID
   - Product ID
   - Quantity Sold
   - Selling Price
   - Total Sale
   - Date

3. **Calculations Sheet**
   - Product-level metrics
   - Current Stock (Initial Quantity - Total Sold)
   - Stock Status (Low/Medium/Good)
   - Total Revenue by Product
   - Total Profit by Product
   - Color-coded stock alerts

4. **Dashboard Sheet**
   - KPI Cards
   - Interactive Charts
   - Instructions

### 📈 Key Performance Indicators (KPIs)

The dashboard displays the following KPIs:

1. **Total Sales** - Sum of all sales revenue
2. **Total Profit** - Total profit (Revenue - Cost)
3. **Total Orders** - Number of orders processed
4. **Average Order Value** - Average sales per order
5. **Low Stock Products** - Number of products with stock < 50 units

### 📊 Visualizations

1. **Sales by Month** - Line chart showing sales trends over time
2. **Top 10 Selling Products** - Bar chart of best-performing products
3. **Stock Status** - Color-coded inventory levels (Red: Low, Yellow: Medium, Green: Good)

## 🚀 How to Use

### Getting Started
1. Open `Sales_Inventory_Dashboard.xlsx` in Microsoft Excel 2021
2. Navigate to the **Dashboard** sheet for an overview
3. Use **Products** and **Sales** sheets to view/edit raw data
4. Check **Calculations** sheet for detailed product metrics

### Adding Interactive Slicers (Excel 2021)

To add slicers for interactive filtering:

1. **Select a Table**
   - Click anywhere in `tbl_Products` or `tbl_Sales`

2. **Insert Slicer**
   - Go to: `Table Design` → `Insert Slicer`
   - Check desired fields:
     - Date (for timeline filtering)
     - Category
     - Product Name
     - Supplier

3. **Position the Slicer**
   - Drag the slicer to the Dashboard sheet
   - Resize as needed

4. **Connect to Multiple Tables** (Advanced)
   - Right-click slicer → `Report Connections`
   - Check all tables you want to filter

### Creating Pivot Tables

To create additional pivot tables:

1. **Select Data Source**
   - Click in `tbl_Sales` or `tbl_Products`

2. **Insert Pivot Table**
   - Go to: `Insert` → `PivotTable`
   - Choose location (new worksheet or Dashboard)

3. **Configure Fields**
   - Drag fields to Rows, Columns, Values areas
   - Example: Category (Rows) + Total Sale (Values)

4. **Add Charts**
   - Select Pivot Table
   - Go to: `PivotTable Analyze` → `PivotChart`

### Creating Data Relationships

To link Products and Sales tables:

1. Go to: `Data` → `Relationships`
2. Click: `New...`
3. Set up relationship:
   - Table: `tbl_Sales`
   - Column: `Product ID`
   - Related Table: `tbl_Products`
   - Related Column: `Product ID`
4. Click: `OK`

## 📊 Sample Data

The dashboard includes **50 products** and **200 sales transactions** spanning a full year (2025). This sample data allows you to:
- See the dashboard in action immediately
- Understand the data structure
- Test filtering and interactivity
- Customize for your needs

### Data Breakdown
- **5 Categories**: Electronics, Clothing, Home & Garden, Sports, Books
- **5 Suppliers**: Supplier A through E
- **Date Range**: January 1, 2025 - December 31, 2025
- **Price Range**: $10 - $400 (realistic business values)

## 🎨 Design Features

### Color Scheme
- **Primary Blue**: `#4472C4` (Headers, titles, KPI values)
- **Light Gray**: `#E7E6E6` (KPI backgrounds)
- **Status Colors**:
  - Red `#FF6B6B` - Low stock (< 50 units)
  - Yellow `#FFD93D` - Medium stock (50-100 units)
  - Green `#6BCF7F` - Good stock (> 100 units)

### Layout
- Clean, modern design
- No gridlines on Dashboard
- Professional formatting
- Proper spacing and alignment
- Color-coded visual alerts

## 🔧 Customization

### Adding Your Own Data

1. **Replace Sample Data**
   - Open Products sheet
   - Select all data (keep headers)
   - Delete and paste your product data
   - Repeat for Sales sheet

2. **Maintain Table Structure**
   - Keep the same column names
   - Ensure Product IDs match between tables
   - Use consistent date formats

3. **Update Calculations**
   - Calculations sheet will auto-update
   - Charts will refresh automatically

### Modifying KPIs

To change KPI calculations:
1. Go to Dashboard sheet
2. Modify formulas in cells B4, D4, F4, H4, J4
3. Adjust thresholds (e.g., low stock threshold)

### Adding More Charts

1. Create new Pivot Table from data
2. Insert Pivot Chart
3. Place on Dashboard sheet
4. Connect to slicers for interactivity

## 📋 Technical Specifications

- **Excel Version**: Compatible with Excel 2021
- **File Format**: .xlsx (Office Open XML)
- **NO VBA**: Pure Excel formulas and features
- **NO Office 365-only functions**: Uses standard Excel 2021 features
- **Tables**: Structured Excel Tables (not ranges)
- **Charts**: Native Excel charts (Line, Bar)

## 🎯 Use Cases

This dashboard is perfect for:
- Small to medium businesses tracking inventory
- Retail stores monitoring sales performance
- E-commerce businesses analyzing product trends
- Sales teams tracking KPIs
- Inventory managers monitoring stock levels
- Business analysts creating reports

## 📚 Learning Resources

To get the most out of this dashboard:

1. **Excel Tables**: [Microsoft Documentation](https://support.microsoft.com/excel)
2. **Pivot Tables**: Practice with the included data
3. **Slicers**: Experiment with filtering
4. **Charts**: Customize chart styles and types

## ⚙️ Maintenance

### Best Practices
1. **Regular Backups**: Save copies before major changes
2. **Data Validation**: Ensure Product IDs are consistent
3. **Date Formats**: Use standard date formats
4. **Calculations**: Verify formulas after data updates

### Troubleshooting

**Charts not updating?**
- Right-click chart → Refresh Data
- Check data source ranges

**Slicers not working?**
- Verify slicer connections
- Ensure tables are properly formatted

**Formulas showing errors?**
- Check for missing Product IDs
- Verify table names haven't changed

## 📝 Notes

- The dashboard is designed to be maintained without programming knowledge
- All features use native Excel functionality
- Charts and KPIs update automatically when data changes
- The file is ready for immediate business use
- Easily extensible for additional metrics

## 🤝 Support

For questions or customization help:
1. Review Excel's built-in help (F1)
2. Check Microsoft Excel documentation
3. Modify the Python script (`create_dashboard.py`) to regenerate with different data

## 📄 License

This dashboard template is provided as-is for business and educational use.

---

**Created**: February 2026
**Compatible with**: Microsoft Excel 2021
**No VBA Required**: ✓
**Interactive**: ✓
**Professional**: ✓
