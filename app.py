import os
import csv
import requests
import pandas as pd
import io
from datetime import datetime
import urllib.parse
import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

def read_keywords(file):
    """Read keywords from uploaded file"""
    try:
        df = pd.read_csv(file)
        # Assuming keywords are in the first column
        keywords = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
        return keywords
    except Exception as e:
        st.error(f"Error reading parameters file: {e}")
        return []

def download_csv_for_keyword(keyword):
    """Download CSV for a specific keyword"""
    try:
        encoded_keyword = urllib.parse.quote(keyword)
        url = f"https://find-and-update.company-information.service.gov.uk/advanced-search/download?companyNameIncludes=&companyNameExcludes=&registeredOfficeAddress=LE5&incorporationFromDay=&incorporationFromMonth=&incorporationFromYear=&incorporationToDay=&incorporationToMonth=&incorporationToYear=&sicCodes={encoded_keyword}&dissolvedFromDay=&dissolvedFromMonth=&dissolvedFromYear=&dissolvedToDay=&dissolvedToMonth=&dissolvedToYear="
        response = requests.get(url)
        response.raise_for_status()
        df = pd.read_csv(io.StringIO(response.text))
        return df
    except Exception as e:
        st.warning(f"Error downloading CSV for keyword {keyword}: {e}")
        return None

def process_dissolution_date(date_str):
    """Process dissolution date string"""
    if pd.isna(date_str) or date_str == '':
        return None
    try:
        return datetime.strptime(date_str, '%d/%m/%Y')
    except:
        try:
            return datetime.strptime(date_str, '%Y-%m-%d')
        except:
            return None

def filter_and_append_data(input_df):
    """Filter data based on dissolution date"""
    filtered_df = input_df[
        (input_df['dissolution_date'].apply(process_dissolution_date).isna()) | 
        (input_df['dissolution_date'].apply(process_dissolution_date) > datetime(2019, 1, 1))
    ]
    return filtered_df

def process_company_stats(data_df):
    """Generate statistics about companies"""
    stats = {
        'Active Companies': len(data_df[data_df['company_status'] == 'Active']),
        'Dissolved Companies': len(data_df[data_df['company_status'] == 'Dissolved']),
        'Companies in Liquidation': len(data_df[data_df['company_status'] == 'Liquidation'])
    }
    
    data_df['dissolution_year'] = data_df['dissolution_date'].apply(
        lambda x: process_dissolution_date(x).year if process_dissolution_date(x) else None
    )
    
    dissolution_by_year = data_df[data_df['dissolution_year'].notna()]['dissolution_year'].value_counts().to_dict()
    stats['Dissolution by Year'] = dissolution_by_year
    
    return stats

def main():
    st.title("Company Data Processor")
    st.write("Upload a CSV file containing SIC codes to search for companies.")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Parameters CSV", type=['csv'])
    
    if uploaded_file is not None:
        with st.spinner('Processing...'):
            # Read keywords
            keywords = read_keywords(uploaded_file)
            
            if not keywords:
                st.error("No valid keywords found in the uploaded file.")
                return
            
            # Accumulated data
            all_data = pd.DataFrame()
            active_company_addresses = pd.DataFrame(columns=['company_name', 'registered_office_address'])
            
            # Process each keyword
            progress_bar = st.progress(0)
            for idx, keyword in enumerate(keywords):
                progress_bar.progress((idx + 1) / len(keywords))
                
                # Download CSV
                keyword_df = download_csv_for_keyword(keyword)
                
                if keyword_df is not None:
                    # Filter data
                    filtered_df = filter_and_append_data(keyword_df)
                    all_data = pd.concat([all_data, filtered_df], ignore_index=True)
                    
                    # Collect active company addresses
                    active_addresses = filtered_df[filtered_df['company_status'] == 'Active'][['company_name', 'registered_office_address']]
                    active_company_addresses = pd.concat([active_company_addresses, active_addresses], ignore_index=True)
            
            # Remove duplicates
            all_data.drop_duplicates(inplace=True)
            active_company_addresses.drop_duplicates(inplace=True)
            
            # Generate stats
            stats = process_company_stats(all_data)
            
            # Create Excel file in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                all_data.to_excel(writer, sheet_name='Master_Data', index=False)
                pd.DataFrame.from_dict(stats, orient='index').to_excel(writer, sheet_name='Stats')
                active_company_addresses.to_excel(writer, sheet_name='Active_Addresses', index=False)
                
                # Highlight dissolved companies
                workbook = writer.book
                data_sheet = workbook['Master_Data']
                red_fill = PatternFill(start_color='FFCCCB', end_color='FFCCCB', fill_type='solid')
                
                for row in data_sheet.iter_rows(min_row=2, max_row=data_sheet.max_row, min_col=1, max_col=data_sheet.max_column):
                    if row[data_sheet.max_column - 1].value == 'Dissolved':
                        for cell in row:
                            cell.fill = red_fill
            
            output.seek(0)
            
            # Show results
            st.success("Processing complete!")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="Download Excel Report",
                    data=output,
                    file_name="Company_Data.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with col2:
                st.download_button(
                    label="Download Active Addresses (CSV)",
                    data=active_company_addresses.to_csv(index=False).encode('utf-8'),
                    file_name="Active_Addresses.csv",
                    mime="text/csv"
                )
            
            # Show previews
            st.subheader("Preview of Master Data")
            st.dataframe(all_data.head())
            
            st.subheader("Preview of Active Company Addresses")
            st.dataframe(active_company_addresses.head())
            
            st.subheader("Statistics")
            st.write(stats)

if __name__ == "__main__":
    main()
