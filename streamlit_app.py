import streamlit as st
import pandas as pd
import os
from main import AmazonSalesAnalysis
import tempfile

st.set_page_config(
    page_title="Excel Sheet Analyzer",
    page_icon="📊",
    layout="wide"
)

# Custom CSS to improve the look
st.markdown("""
    <style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .upload-header {
        text-align: center;
        padding: 2rem 0;
    }
    .success-message {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #d4edda;
        color: #155724;
        margin: 1rem 0;
    }
    .error-message {
        padding: 1rem;
        border-radius: 0.5rem;
        background-color: #f8d7da;
        color: #721c24;
        margin: 1rem 0;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    st.title("📊 Excel Sheet Analyzer")
    st.markdown("### Upload your Excel file for detailed analysis")
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # Create a temporary file
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_path = tmp_file.name
            
            # Process the file
            with st.spinner('Analyzing your file...'):
                analyzer = AmazonSalesAnalysis(tmp_path)
                if analyzer.load_data():
                    # Show basic statistics
                    st.success("File analyzed successfully!")
                    
                    # Display statistics in columns
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.subheader("📈 Basic Statistics")
                        stats = analyzer.basic_statistics()
                        for key, value in stats.items():
                            st.metric(label=key, value=value)
                    
                    with col2:
                        st.subheader("📊 Category Analysis")
                        cat_analysis = analyzer.sales_by_category()
                        st.dataframe(cat_analysis)
                    
                    # Generate and offer download of report
                    analyzer.generate_excel_report()
                    
                    # Read the generated report
                    with open('sales_analysis_report.xlsx', 'rb') as report_file:
                        report_data = report_file.read()
                    
                    st.download_button(
                        label="📥 Download Full Analysis Report",
                        data=report_data,
                        file_name="sales_analysis_report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Show monthly trends visualization
                    st.subheader("📈 Monthly Sales Trends")
                    analyzer.monthly_trends()
                    st.image('monthly_sales_trend(1).png')
                    
                    # Show top products
                    st.subheader("🏆 Top Products")
                    top_products = analyzer.top_products()
                    st.dataframe(top_products)
                
                else:
                    st.error("Failed to analyze the file. Please check if it's in the correct format.")
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
        
        finally:
            # Clean up temporary files
            try:
                os.unlink(tmp_path)
                os.unlink('sales_analysis_report.xlsx')
                os.unlink('monthly_sales_trend(1).png')
            except:
                pass

if __name__ == "__main__":
    main() 