# Amazon Store Sales Analysis Tool

This tool helps analyze data from Excel Files and generates comprehensive visualizations in PowerPoint.

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
Pie charts of data required into PowerPoint