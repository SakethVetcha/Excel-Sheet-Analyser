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
        # Convert Price column to numeric, replacing any errors with NaN
        if 'Price' in self.df.columns:
            self.df['Price'] = pd.to_numeric(self.df['Price'], errors='coerce')

    def basic_statistics(self):
        if self.df is None:
            return None
        stats = {}
        
        try:
            # Price statistics (always available)
            total_revenue = self.df['Price'].sum()
            avg_price = self.df['Price'].mean()
            max_price = self.df['Price'].max()
            min_price = self.df['Price'].min()
            
            # Format statistics with error handling
            stats["Total Revenue"] = f"${total_revenue:,.2f}" if pd.notnull(total_revenue) else "N/A"
            stats["Average Price"] = f"${avg_price:,.2f}" if pd.notnull(avg_price) else "N/A"
            stats["Highest Price"] = f"${max_price:,.2f}" if pd.notnull(max_price) else "N/A"
            stats["Lowest Price"] = f"${min_price:,.2f}" if pd.notnull(min_price) else "N/A"
            
            # Item statistics
            stats["Total Unique Items"] = str(len(self.df['Item'].unique()))
            stats["Total Transactions"] = str(len(self.df))
            
            # Type statistics if available
            if self.has_type:
                stats["Number of Categories"] = str(len(self.df['Type'].unique()))
        except Exception as e:
            st.error(f"Error calculating statistics: {str(e)}")
            return pd.Series({"Error": "Could not calculate statistics"})
        
        return pd.Series(stats)
    
    def type_analysis(self):
        if self.df is None or not self.has_type:
            return None
            
        try:
            type_analysis = self.df.groupby('Type').agg({
                'Price': ['sum', 'mean', 'count'],
                'Item': 'nunique'
            }).round(2)
            
            type_analysis.columns = ['Total Revenue', 'Average Price', 'Number of Sales', 'Unique Items']
            type_analysis = type_analysis.sort_values('Total Revenue', ascending=False)
            type_analysis['Revenue(%)'] = (type_analysis['Total Revenue'] / type_analysis['Total Revenue'].sum() * 100).round(2)
            
            return type_analysis
        except Exception as e:
            st.error(f"Error in type analysis: {str(e)}")
            return None
    
    def time_trends(self):
        if self.df is None:
            return None
            
        if not self.has_date:
            return None
            
        try:
            # Convert to datetime if not already
            if not pd.api.types.is_datetime64_any_dtype(self.df['Date']):
                self.df['Date'] = pd.to_datetime(self.df['Date'])
            
            time_data = self.df.set_index('Date').resample('M').agg({
                'Price': 'sum',
                'Item': 'count'
            }).reset_index()
            
            fig, ax = plt.subplots(figsize=(10, 6))
            
            # Revenue trend
            ax.plot(time_data['Date'], time_data['Price'], marker='o', linewidth=2, color='#1f77b4')
            ax.set_title('Monthly Revenue Trends', pad=20)
            ax.set_xlabel('Month')
            ax.set_ylabel('Total Revenue ($)')
            ax.grid(True, linestyle='--', alpha=0.7)
            ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda x, p: f'${x:,.0f}'))
            
            # Add number of transactions as text annotations
            for i, (date, revenue, count) in enumerate(zip(time_data['Date'], time_data['Price'], time_data['Item'])):
                ax.annotate(f'{count} sales', (date, revenue),textcoords="offset points",xytext=(0,10),ha='center',fontsize=8)
            
            plt.tight_layout()
            return fig
        except Exception as e:
            st.error(f"Error generating time trends: {str(e)}")
            return None

    def top_items(self, n=10):
        if self.df is None:
            return None
            
        try:
            top_items = self.df.groupby('Item').agg({
                'Price': ['sum', 'mean', 'count']
            }).round(2)
            
            top_items.columns = ['Total Revenue', 'Average Price', 'Times Sold']
            top_items = top_items.sort_values('Total Revenue', ascending=False)
            
            if self.has_type:
                # Add most common type for each item
                item_types = self.df.groupby('Item')['Type'].agg(lambda x: x.mode()[0] if len(x.mode()) > 0 else 'Unknown')
                top_items['Primary Type'] = item_types
                
            return top_items.head(n)
        except Exception as e:
            st.error(f"Error analyzing top items: {str(e)}")
            return None
    
    def generate_excel_report(self):
        if self.df is None:
            return None
            
        try:
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Basic Statistics
                pd.DataFrame(self.basic_statistics()).to_excel(writer, sheet_name='Basic Statistics')
                ws = writer.sheets['Basic Statistics']
                ws.column_dimensions['A'].width = 25
                ws.column_dimensions['B'].width = 15
                
                # Type Analysis if available
                if self.has_type:
                    type_analysis = self.type_analysis()
                    if type_analysis is not None:
                        type_analysis.to_excel(writer, sheet_name='Type Analysis')
                        ws = writer.sheets['Type Analysis']
                        ws.column_dimensions['A'].width = 30
                        for col in ws.columns:
                            if col[0].column_letter != 'A':
                                ws.column_dimensions[col[0].column_letter].width = 15
                
                # Top Items
                top = self.top_items(n=10)
                if top is not None:
                    top.to_excel(writer, sheet_name='Top Items')
                    ws = writer.sheets['Top Items']
                    ws.column_dimensions['A'].width = 35
                    for col in ws.columns:
                        if col[0].column_letter != 'A':
                            ws.column_dimensions[col[0].column_letter].width = 15

                # Monthly Analysis
                if self.has_date:
                    try:
                        # Convert to datetime if not already
                        if not pd.api.types.is_datetime64_any_dtype(self.df['Date']):
                            self.df['Date'] = pd.to_datetime(self.df['Date'])
                        
                        # Create monthly summary
                        monthly_data = self.df.set_index('Date').resample('M').agg({
                            'Price': 'sum',
                            'Item': 'count'
                        }).reset_index()
                        
                        # Create Monthly Trends sheet with chart
                        monthly_data.to_excel(writer, sheet_name='Monthly Trends', index=False)
                        ws = writer.sheets['Monthly Trends']
                        ws.column_dimensions['A'].width = 20
                        ws.column_dimensions['B'].width = 15
                        ws.column_dimensions['C'].width = 15
                        
                        # Create chart
                        from openpyxl.chart import LineChart, Reference
                        chart = LineChart()
                        chart.title = "Monthly Revenue Trends"
                        chart.y_axis.title = "Revenue ($)"
                        chart.x_axis.title = "Month"
                        
                        data = Reference(ws, min_col=2, min_row=1, max_row=len(monthly_data)+1, max_col=2)
                        cats = Reference(ws, min_col=1, min_row=2, max_row=len(monthly_data)+1)
                        
                        chart.add_data(data, titles_from_data=True)
                        chart.set_categories(cats)
                        ws.add_chart(chart, "E2")
                        
                        # Create detailed monthly sheets
                        from openpyxl.chart import LineChart, Reference
                        
                        for name, group in self.df.groupby(self.df['Date'].dt.strftime('%Y-%m')):
                            sheet_name = f"Month_{name}"
                            # Group by item and calculate stats
                            monthly_items = group.groupby('Item').agg({
                                'Price': ['sum', 'mean', 'count']
                            }).round(2)
                            monthly_items.columns = ['Total Revenue', 'Average Price', 'Number of Sales']
                            monthly_items = monthly_items.sort_values('Total Revenue', ascending=False)
                            
                            monthly_items.to_excel(writer, sheet_name=sheet_name)
                            ws = writer.sheets[sheet_name]
                            ws.column_dimensions['A'].width = 35
                            for col in ws.columns:
                                if col[0].column_letter != 'A':
                                    ws.column_dimensions[col[0].column_letter].width = 15
                            
                            # Add conditional formatting to highlight top performers
                            from openpyxl.styles import PatternFill
                            green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
                            yellow_fill = PatternFill(start_color='FFD700', end_color='FFD700', fill_type='solid')
                            
                            # Highlight top 3 revenue items
                            for i in range(2, min(5, len(monthly_items) + 2)):
                                for j in range(1, 5):
                                    cell = ws.cell(row=i, column=j)
                                    if i == 2:
                                        cell.fill = green_fill
                                    elif i == 3 or i == 4:
                                        cell.fill = yellow_fill
                            
                            # Create comparison line chart for each month
                            chart = LineChart()
                            chart.title = f"Revenue vs Quantity - {name}"
                            chart.style = 10
                            chart.y_axis.title = "Revenue ($) / Quantity"
                            chart.x_axis.title = "Items"
                            
                            # Get top 10 items by revenue for the chart
                            top_10_items = monthly_items.head(10)
                            
                            # Add revenue data
                            revenue_data = Reference(ws, min_col=2, min_row=1, max_row=len(top_10_items)+1, max_col=2)
                            cats = Reference(ws, min_col=1, min_row=2, max_row=len(top_10_items)+1)
                            chart.add_data(revenue_data, titles_from_data=True)
                            
                            # Add quantity data
                            quantity_data = Reference(ws, min_col=4, min_row=1, max_row=len(top_10_items)+1, max_col=4)
                            chart.add_data(quantity_data, titles_from_data=True)
                            
                            # Customize line chart
                            chart.set_categories(cats)
                            chart.height = 15  # Height in cm
                            chart.width = 25   # Width in cm
                            
                            # Configure legend position
                            chart.legend.position = 'r'
                            
                            # Make lines more visible
                            for series in chart.series:
                                series.smooth = True  # Make lines smooth
                                series.marker.symbol = "circle"  # Add markers
                                series.marker.size = 8  # Make markers visible
                            
                            # Position the chart below the data
                            ws.add_chart(chart, f"A{len(monthly_items) + 5}")
                                        
                    except Exception as e:
                        st.warning(f"Could not create monthly analysis: {str(e)}")
                        
            output.seek(0)
            return output
        except Exception as e:
            st.error(f"Error generating Excel report: {str(e)}")
            return None

# Streamlit UI
st.set_page_config(page_title="Flexible Data Analyzer", layout="wide")
st.title("📊 Data Analysis Dashboard")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine="openpyxl")
        st.write("Preview of uploaded data:")
        st.dataframe(df.head())

        # Filter out index-like columns
        valid_columns = [col for col in df.columns if not (
            col == "Unnamed: 0" or 
            col == "index" or 
            str(col).isdigit() or 
            str(col).startswith('Unnamed:')
        )]

        st.info("Map your columns to the required fields:")
        
        # Required columns
        item_col = st.selectbox("Select the Item column (required)", valid_columns, key="item_col")
        price_col = st.selectbox("Select the Price column (required)", valid_columns, key="price_col")
        
        # Optional columns
        date_col = st.selectbox("Select the Date column (optional)", ["None"] + valid_columns, key="date_col")
        type_col = st.selectbox("Select the Type column (optional)", ["None"] + valid_columns, key="type_col")

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
                    st.warning(f"Could not parse Date column: {str(e)}")

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
    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
else:
    st.info("Please upload an Excel file to begin analysis.")

