import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib as mpl
import os
import time 
from datetime import datetime

class AmazonSalesAnalysis:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None

    def load_data(self):
        try:
            self.df = pd.read_excel(self.excel_file, engine="openpyxl")
            self.df['Date'] = pd.to_datetime(self.df['Date'])
            print("Data loaded successfully!")
            print(f"Total records: {len(self.df)}")
            print("\nData Overview:")
            print(f"Date Range: {self.df['Date'].min().strftime('%Y-%m-%d')} to {self.df['Date'].max().strftime('%Y-%m-%d')}")
            print(f"Total Categories: {len(self.df['Category'].unique())}")
            print(f"Total Products: {len(self.df['Product'].unique())}")
            return True
        except Exception as e:
            print(f"Error loading as {str(e)}")
            return False

    def basic_statistics(self):
        if self.df is None:
            return "First load excel data"
        
        stats = {
            "Total Sales Revenue": f"${self.df['Sales'].sum():,.2f}",
            "Average Sales Amount": f"${self.df['Sales'].mean():,.2f}",
            "Highest Single Sale": f"${self.df['Sales'].max():,.2f}",
            "Lowest Single Sale": f"${self.df['Sales'].min():,.2f}",
            "Total Products Sold": f"{self.df['Quantity'].sum():,}",
            "Total Unique Products": f"{len(self.df['Product'].unique()):,}",
            "Average Items Per Sale": f"{self.df['Quantity'].mean():.1f}"
        }
        return pd.Series(stats)
    
    def sales_by_category(self):
        if self.df is None:
            return "First load excel data"
        
        category_sales = self.df.groupby('Category').agg({
            'Sales': ['sum', 'mean', 'count'],
            'Quantity': 'sum'
        }).round(2)

        category_sales.columns = ['Total Sales', 'Average Sales', 'Number of Orders', 'Units Sold']
        category_sales = category_sales.sort_values('Total Sales', ascending=False)
        category_sales['Sales(%)'] = (category_sales['Total Sales'] / category_sales['Total Sales'].sum() * 100).round(2)
        category_sales['Orders(%)'] = (category_sales['Number of Orders'] / category_sales['Number of Orders'].sum() * 100).round(2)

        return category_sales
    
    def monthly_trends(self):
        if self.df is None:
            return "First load Excel sheet"
        
        monthly_data = self.df.set_index('Date').resample('M').agg({
            'Sales': 'sum',
            'Quantity': 'sum'
        }).reset_index()

        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 12))

        ax1.plot(monthly_data['Date'], monthly_data['Sales'], marker='o', linewidth=2, color='#1f77b4')
        ax1.set_title('Monthly Sales Trends', pad=20)
        ax1.set_xlabel('Month')
        ax1.set_ylabel('Total Sales ($)')
        ax1.grid(True, linestyle='--', alpha=0.7)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        ax2.plot(monthly_data['Date'], monthly_data['Quantity'], marker='s', linewidth=2, color='#ff7f0e')
        ax2.set_title('Monthly Units Sold Trend', pad=20)
        ax2.set_xlabel('Month')
        ax2.set_ylabel('Units Sold')
        ax2.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()
        plt.savefig('monthly_sales_trend(1).png', dpi=300, bbox_inches='tight')
        plt.close()

    def top_products(self, n=10):
        if self.df is None:
            return "First load Excel data"
        top_products = self.df.groupby('Product').agg({
            'Sales': 'sum',
            'Quantity': 'sum'
        }).round(2)
        top_products['Average Price'] = (top_products['Sales'] / top_products['Quantity']).round(2)
        top_products = top_products.sort_values('Sales', ascending=False)
        return top_products.head(n)
    
    def generate_excel_report(self):
        if self.df is None:
            return "Please load data first."
        
        output_file = 'sales_analysis_report.xlsx'
        try:
            with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
                # Basic Statistics
                pd.DataFrame(self.basic_statistics()).to_excel(writer, sheet_name='Basic Statistics')
                
                # Category Analysis
                self.sales_by_category().to_excel(writer, sheet_name='Category Analysis')
                
                # Top Products
                self.top_products(n=10).to_excel(writer, sheet_name='Top Products')
                
                # Monthly Trends
                monthly_data = self.df.set_index('Date').resample('M').agg({
                    'Sales': 'sum',
                    'Quantity': 'sum'
                }).reset_index()
                monthly_data.to_excel(writer, sheet_name='Monthly Trends', index=False)
                
            print(f"\nExcel report generated successfully: {output_file}")
            return True
        except Exception as e:
            print(f"Error loading as {str(e)}")
            return False
        
#main function
def main():
    
    analyzer = AmazonSalesAnalysis('sample_data.xlsx')
    
    if analyzer.load_data():
        print("\nGenerating Analysis...")
        # Print basic statistics
        print("\nBasic Statistics:")
        print(analyzer.basic_statistics())
        # Print category analysis
        print("\nCategory Analysis:")
        print(analyzer.sales_by_category())
        # Generate monthly trend graph
        print("\nGenerating monthly trends visualization...")
        analyzer.monthly_trends()
        # Print top products
        print("\nTop 10 Products by Sales:")
        print(analyzer.top_products())
        # Generate Excel report
        print("\nGenerating Excel report...")
        analyzer.generate_excel_report()
        print("\nAnalysis complete! Check 'sales_analysis_report.xlsx' for detailed report.")
        print("Monthly trends visualization saved as 'monthly_sales_trend(1).png'")

if __name__ == "__main__":
    main()









