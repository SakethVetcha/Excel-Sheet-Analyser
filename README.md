# Amazon Store Sales Analysis Tool

This tool helps analyze sales data from your Amazon store and generates comprehensive reports and visualizations.

## Setup

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

2. Prepare your sales data in an Excel file named 'amazon_sales_data.xlsx' with the following columns:
   - Date: The date of the sale (YYYY-MM-DD format)
   - Product: Product name
   - Category: Product category
   - Sales: Sale amount
   - Quantity: Number of units sold
   - Price: Unit price

3. Run the analysis:
```bash
python amazon_sales_analysis.py
```

## Output

The tool will generate:
1. A comprehensive Excel report ('sales_analysis_report.xlsx') containing:
   - Basic Statistics
   - Category Analysis
   - Top Products
   - Monthly Trends
2. A visualization of monthly sales trends ('monthly_sales_trend.png')

## Features

- Basic sales statistics (total, average, highest, lowest sales)
- Category-wise analysis
- Monthly sales trends
- Top performing products
- Comprehensive Excel report generation
- Sales visualization 
