import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from io import BytesIO
from datetime import datetime

class FlexibleDataAnalysis:
    def __init__(self, df):
        self.df = df
        self.has_date = 'Date' in df.columns
        self.has_type = 'Type' in df.columns

    def basic_statistics(self):
        if self.df is None:
            return None
        stats = {}
        
        # Price statistics (always available)
        stats["Total Revenue"] = f"${self.df['Price'].sum():,.2f}"
        stats["Average Price"] = f"${self.df['Price'].mean():,.2f}"
        stats["Highest Price"] = f"${self.df['Price'].max():,.2f}"
        stats["Lowest Price"] = f"${self.df['Price'].min():,.2f}"
        
        # Item statistics
        stats["Total Unique Items"] = f"{len(self.df['Item'].unique()):,}"
        stats["Total Transactions"] = f"{len(self.df):,}"
        
        # Type statistics if available
        if self.has_type:
            stats["Number of Categories"] = f"{len(self.df['Type'].unique()):,}"
        
        return pd.Series(stats)
    
    def type_analysis(self):
        if self.df is None or not self.has_type:
            return None
            
        type_analysis = self.df.groupby('Type').agg({
            'Price': ['sum', 'mean', 'count'],
            'Item': 'nunique'
        }).round(2)
        
        type_analysis.columns = ['Total Revenue', 'Average Price', 'Number of Sales', 'Unique Items']
        type_analysis = type_analysis.sort_values('Total Revenue', ascending=False)
        type_analysis['Revenue(%)'] = (type_analysis['Total Revenue'] / type_analysis['Total Revenue'].sum() * 100).round(2)
        
        return type_analysis
    
    def time_trends(self):
        if self.df is None or not self.has_date:
            return None
            
        # Convert to datetime if not already
        if not pd.api.types.is_datetime64_any_dtype(self.df['Date']):
            try:
                self.df['Date'] = pd.to_datetime(self.df['Date'])
            except:
                return None
        
        time_data = self.df.set_index('Date').resample('M').agg({
            'Price': 'sum',
            'Item': 'count'
        }).reset_index()
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 8))
        
        # Revenue trend
        ax1.plot(time_data['Date'], time_data['Price'], marker='o', linewidth=2, color='#1f77b4')
        ax1.set_title('Monthly Revenue Trends', pad=20)
        ax1.set_xlabel('Month')
        ax1.set_ylabel('Total Revenue ($)')
        ax1.grid(True, linestyle='--', alpha=0.7)
        ax1.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
        
        # Transaction count trend
        ax2.plot(time_data['Date'], time_data['Item'], marker='s', linewidth=2, color='#ff7f0e')
        ax2.set_title('Monthly Transaction Count', pad=20)
        ax2.set_xlabel('Month')
        ax2.set_ylabel('Number of Transactions')
        ax2.grid(True, linestyle='--', alpha=0.7)
        
        plt.tight_layout()
        return fig

    def top_items(self, n=10):
        if self.df is None:
            return None
            
        top_items = self.df.groupby('Item').agg({
            'Price': ['sum', 'mean', 'count']
        }).round(2)
        
        top_items.columns = ['Total Revenue', 'Average Price', 'Times Sold']
        top_items = top_items.sort_values('Total Revenue', ascending=False)
        
        if self.has_type:
            # Add most common type for each item
            item_types = self.df.groupby('Item')['Type'].agg(lambda x: x.mode()[0])
            top_items['Primary Type'] = item_types
            
        return top_items.head(n)
    
    def generate_excel_report(self):
        if self.df is None:
            return None
            
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # Basic Statistics
            pd.DataFrame(self.basic_statistics()).to_excel(writer, sheet_name='Basic Statistics')
            
            # Type Analysis if available
            if self.has_type:
                type_analysis = self.type_analysis()
                if type_analysis is not None:
                    type_analysis.to_excel(writer, sheet_name='Type Analysis')
            
            # Top Items
            top = self.top_items(n=10)
            if top is not None:
                top.to_excel(writer, sheet_name='Top Items')
            
            # Time Analysis if date is available
            if self.has_date:
                try:
                    time_data = self.df.set_index('Date').resample('M').agg({
                        'Price': 'sum',
                        'Item': 'count'
                    }).reset_index()
                    time_data.to_excel(writer, sheet_name='Time Trends', index=False)
                except:
                    pass
                    
        output.seek(0)
        return output

# Streamlit UI
st.set_page_config(page_title="Flexible Data Analyzer", layout="wide")
st.title("📊 Data Analysis Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    st.write("Preview of uploaded data:")
    st.dataframe(df.head())

    st.info("Map your columns to the required fields:")
    columns = df.columns.tolist()
    
    # Required columns
    item_col = st.selectbox("Select the Item column (required)", columns, key="item_col")
    price_col = st.selectbox("Select the Price column (required)", columns, key="price_col")
    
    # Optional columns
    date_col = st.selectbox("Select the Date column (optional)", ["None"] + columns, key="date_col")
    type_col = st.selectbox("Select the Type column (optional)", ["None"] + columns, key="type_col")

    if item_col and price_col:  # Required columns are selected
        # Rename columns for internal use
        col_map = {
            item_col: 'Item',
            price_col: 'Price'
        }
        
        if date_col != "None":
            col_map[date_col] = 'Date'
        if type_col != "None":
            col_map[type_col] = 'Type'
            
        df_renamed = df.rename(columns=col_map)

        # Try to parse date if present
        if 'Date' in df_renamed.columns:
            try:
                df_renamed['Date'] = pd.to_datetime(df_renamed['Date'])
            except Exception as e:
                st.warning(f"Could not parse Date column: {e}")

        analyzer = FlexibleDataAnalysis(df_renamed)
        st.success("Data loaded successfully!")
        
        st.subheader("Basic Statistics")
        stats = analyzer.basic_statistics()
        if stats is not None:
            st.table(stats)
        
        if analyzer.has_type:
            st.subheader("Analysis by Type")
            type_analysis = analyzer.type_analysis()
            if type_analysis is not None:
                st.dataframe(type_analysis)
        
        if analyzer.has_date:
            st.subheader("Time Trends")
            fig = analyzer.time_trends()
            if fig is not None:
                st.pyplot(fig)
        
        st.subheader("Top 10 Items by Revenue")
        top_items = analyzer.top_items()
        if top_items is not None:
            st.dataframe(top_items)
        
        st.subheader("Download Excel Report")
        excel_report = analyzer.generate_excel_report()
        if excel_report is not None:
            st.download_button(
                label="Download Analysis Excel Report",
                data=excel_report,
                file_name="analysis_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please select the required Item and Price columns to begin analysis.")
else:
    st.info("Please upload an Excel file to begin analysis.")

