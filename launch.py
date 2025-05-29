import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from datetime import datetime

# Your analysis class (slightly modified for Streamlit)
class AmazonSalesAnalysis:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None

    def load_data(self):
        try:
            self.df = pd.read_excel(self.excel_file, engine="openpyxl")
            self.df['Date'] = pd.to_datetime(self.df['Date'])
            return True
        except Exception as e:
            st.error(f"Error loading file: {str(e)}")
            return False

    def basic_statistics(self):
        if self.df is None:
            return None
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
            return None
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
            return None
        monthly_data = self.df.set_index('Date').resample('M').agg({
            'Sales': 'sum',
            'Quantity': 'sum'
        }).reset_index()
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
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
        return fig

    def top_products(self, n=10):
        if self.df is None:
            return None
        top_products = self.df.groupby('Product').agg({
            'Sales': 'sum',
            'Quantity': 'sum'
        }).round(2)
        top_products['Average Price'] = (top_products['Sales'] / top_products['Quantity']).round(2)
        top_products = top_products.sort_values('Sales', ascending=False)
        return top_products.head(n)
    
    def generate_excel_report(self):
        if self.df is None:
            return None
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            pd.DataFrame(self.basic_statistics()).to_excel(writer, sheet_name='Basic Statistics')
            self.sales_by_category().to_excel(writer, sheet_name='Category Analysis')
            self.top_products(n=10).to_excel(writer, sheet_name='Top Products')
            monthly_data = self.df.set_index('Date').resample('M').agg({
                'Sales': 'sum',
                'Quantity': 'sum'
            }).reset_index()
            monthly_data.to_excel(writer, sheet_name='Monthly Trends', index=False)
        output.seek(0)
        return output

# Streamlit UI
st.set_page_config(page_title="Amazon Sales Excel Analyzer", layout="wide")
st.title("📊 Amazon Sales Excel Sheet Analyzer")

uploaded_file = st.file_uploader("Upload your Amazon sales Excel file", type=["xlsx"])

if uploaded_file:
    analyzer = AmazonSalesAnalysis(uploaded_file)
    if analyzer.load_data():
        st.success("Data loaded successfully!")
        st.subheader("Basic Statistics")
        stats = analyzer.basic_statistics()
        if stats is not None:
            st.table(stats)
        
        st.subheader("Sales by Category")
        cat_sales = analyzer.sales_by_category()
        if cat_sales is not None:
            st.dataframe(cat_sales)
        
        st.subheader("Monthly Trends")
        fig = analyzer.monthly_trends()
        if fig is not None:
            st.pyplot(fig)
        
        st.subheader("Top 10 Products by Sales")
        top_prods = analyzer.top_products()
        if top_prods is not None:
            st.dataframe(top_prods)
        
        st.subheader("Download Excel Report")
        excel_report = analyzer.generate_excel_report()
        if excel_report is not None:
            st.download_button(
                label="Download Analysis Excel Report",
                data=excel_report,
                file_name="sales_analysis_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
else:
    st.info("Please upload an Excel file to begin analysis.")

