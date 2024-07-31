import streamlit as st
import pandas as pd
from io import BytesIO

def load_data(file):
    # Load the Excel file
    excel_data = pd.ExcelFile(file)
    # Load the main data starting from row 7 and use row 6 as headers
    main_data = pd.read_excel(file, sheet_name='Results', header=6)
    return main_data

def process_data(main_data):
    # Explicitly cast numerical columns to a compatible dtype before replacing NaN values with empty strings
    for column in main_data.select_dtypes(include=[float, int]).columns:
        main_data[column] = main_data[column].astype('object')
    # Replace NaN values with empty strings
    main_data.fillna('', inplace=True)
    # Convert relevant columns back to numeric for comparison
    main_data['Quantity'] = pd.to_numeric(main_data['Quantity'], errors='coerce')
    main_data['Subtotal'] = pd.to_numeric(main_data['Subtotal'], errors='coerce')
    main_data['Charge Rate'] = pd.to_numeric(main_data['Charge Rate'], errors='coerce')

    # Display the first few rows of the main data
    st.write("Main Data (first 20 rows):")
    st.dataframe(main_data.head(20))

    # Group the data by 'Consignment Number', using 'OTHER' for those with no 'Consignment Number'
    grouped_data = main_data.groupby(main_data['Consignment Number'].apply(lambda x: x if x else 'OTHER'))

    # Calculate the number of transactions and number of consignments
    num_transactions = len(main_data)
    num_consignments = len(grouped_data)

    # Display summary
    st.write(f"Number of Transactions: {num_transactions}")
    st.write(f"Number of Consignments: {num_consignments}")

    # Display the first group
    first_group_key = next(iter(grouped_data.groups))
    first_group = grouped_data.get_group(first_group_key)
    st.write(f"First Consignment: {first_group_key}")
    st.dataframe(first_group.head())

    # Function to remove cancelling transactions
    def remove_cancelling_transactions(group):
        group['keep'] = True
        for idx, row in group.iterrows():
            if row['keep']:
                cancelling_idx = group[
                    (group['Quantity'] == -row['Quantity']) &
                    (group['Subtotal'] == -row['Subtotal']) &
                    (group.index != idx) &
                    (group['keep'])
                ].index
                if not cancelling_idx.empty:
                    group.at[idx, 'keep'] = False
                    group.at[cancelling_idx[0], 'keep'] = False
        return group[group['keep']]

    # Apply the function to each group and concatenate results
    cleaned_groups = [remove_cancelling_transactions(group) for name, group in grouped_data]
    cleaned_data = pd.concat(cleaned_groups).reset_index(drop=True)

    # Calculate the number of transactions and number of consignments after cleaning
    num_transactions_cleaned = len(cleaned_data)
    num_consignments_cleaned = len(cleaned_data['Consignment Number'].unique())

    # Display summary after cleaning
    st.write(f"Number of Transactions after Cleaning: {num_transactions_cleaned}")
    st.write(f"Number of Consignments after Cleaning: {num_consignments_cleaned}")

    # Display the first cleaned group
    first_cleaned_group_key = next(iter(cleaned_data['Consignment Number'].unique()))
    first_cleaned_group = cleaned_data[cleaned_data['Consignment Number'] == first_cleaned_group_key]
    st.write(f"First Cleaned Consignment: {first_cleaned_group_key}")
    st.dataframe(first_cleaned_group.head())

    return cleaned_data

def create_order_output_data(data):
    output_data = pd.DataFrame()
    output_data['Order ID'] = data['Reference Number'].unique()
    output_data['Unique SKUs'] = 0
    output_data['Total number of items in the order'] = 0
    def get_subtotal(description):
        return data[data['Particulars'].str.startswith(description)].groupby('Reference Number')['Subtotal'].sum().reindex(output_data['Order ID'], fill_value=0).values
    def calculate_unique_skus_and_total_items(ref_num):
        order_data = data[data['Reference Number'] == ref_num]
        first_item = order_data[order_data['Particulars'].str.startswith('Item picking (First item)')]
        additional_items = order_data[order_data['Particulars'].str.startswith('Item picking (Additional Items)')]
        if not additional_items.empty:
            total_items = first_item['Quantity'].sum() + additional_items['Quantity'].sum()
        else:
            total_items = first_item['Quantity'].sum()
        unique_skus = len(set(','.join(first_item['Particulars']).split(',')))
        return int(unique_skus), int(total_items)
    output_data[['Unique SKUs', 'Total number of items in the order']] = output_data['Order ID'].apply(lambda ref_num: calculate_unique_skus_and_total_items(ref_num)).apply(pd.Series)
    output_data['Label Fee'] = get_subtotal('Labelling')
    output_data['Picking - Order First Item'] = get_subtotal('Item picking (First item)')
    output_data['Picking - Additional Items'] = get_subtotal('Item picking (Additional Items)')
    other_charges = data[~data['Particulars'].str.startswith(('Labelling', 'Item picking (First item)', 'Item picking (Additional Items)'))]
    output_data['Other Charges'] = other_charges.groupby('Reference Number')['Subtotal'].sum().reindex(output_data['Order ID'], fill_value=0).values
    output_data['Total'] = output_data[['Label Fee', 'Picking - Order First Item', 'Picking - Additional Items', 'Other Charges']].sum(axis=1)
    return output_data

def create_inwards_lcl_output_data(data):
    output_data = pd.DataFrame()
    output_data['Reference'] = data['Reference Number'].unique()
    output_data['Quantity - Cartons'] = data.groupby('Reference Number')['Quantity'].sum().reindex(output_data['Reference'], fill_value=0).astype(int).values
    output_data['Quantity - Pallets'] = 0  # Assuming no pallet information available
    def get_subtotal(description):
        return data[data['Particulars'].str.startswith(description)].groupby('Reference Number')['Subtotal'].sum().reindex(output_data['Reference'], fill_value=0).values
    output_data['Carton Inbound'] = get_subtotal('Inwards Carton')
    output_data['Pallet Inbound'] = get_subtotal('Pallet Inbound')
    other_charges = data[~data['Particulars'].str.startswith(('Inwards Carton', 'Pallet Inbound'))]
    output_data['Other Charges'] = other_charges.groupby('Reference Number')['Subtotal'].sum().reindex(output_data['Reference'], fill_value=0).values
    output_data['Total'] = output_data[['Carton Inbound', 'Pallet Inbound', 'Other Charges']].sum(axis=1)
    return output_data

def create_ad_hoc_output_data(data):
    output_data = pd.DataFrame()
    output_data['Reference'] = data['Particulars']
    output_data['Quantity'] = data['Quantity']
    output_data['Rate'] = data['Charge Rate']
    output_data['Total'] = data['Subtotal']
    return output_data

def check_unique_consignments(source_data, outbound_data, inwards_data, adhoc_data):
    unique_source = source_data['Consignment Number'].nunique()
    unique_outbound = outbound_data['Order ID'].nunique()
    unique_inwards = inwards_data['Reference'].nunique()
    unique_adhoc = adhoc_data['Reference'].nunique()
    return unique_source, unique_outbound, unique_inwards, unique_adhoc

def check_subtotals(source_data, outbound_data, inwards_data, adhoc_data):
    total_source = source_data['Subtotal'].sum()
    total_outbound = outbound_data['Total'].sum()
    total_inwards = inwards_data['Total'].sum()
    total_adhoc = adhoc_data['Total'].sum()
    total_output = total_outbound + total_inwards + total_adhoc
    return total_source, total_outbound, total_inwards, total_adhoc, total_output

def to_excel(outbound_orders_output, inwards_lcl_output, adhoc_output):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        currency_format = workbook.add_format({'num_format': '$#,##0.00'})
        outbound_orders_output.to_excel(writer, sheet_name='Outbound Orders', index=False)
        inwards_lcl_output.to_excel(writer, sheet_name='Inwards LCL', index=False)
        adhoc_output.to_excel(writer, sheet_name='Ad-Hoc', index=False)
        outbound_orders_sheet = writer.sheets['Outbound Orders']
        inwards_lcl_sheet = writer.sheets['Inwards LCL']
        adhoc_sheet = writer.sheets['Ad-Hoc']
        for column in ['Label Fee', 'Picking - Order First Item', 'Picking - Additional Items', 'Other Charges', 'Total']:
            col_idx = outbound_orders_output.columns.get_loc(column)
            outbound_orders_sheet.set_column(col_idx, col_idx, None, currency_format)
        for column in ['Carton Inbound', 'Pallet Inbound', 'Other Charges', 'Total']:
            col_idx = inwards_lcl_output.columns.get_loc(column)
            inwards_lcl_sheet.set_column(col_idx, col_idx, None, currency_format)
        for column in ['Rate', 'Total']:
            col_idx = adhoc_output.columns.get_loc(column)
            adhoc_sheet.set_column(col_idx, col_idx, None, currency_format)
    return output

def main():
    st.title('Consignly Transaction Schedule to B360 Format Converter')
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        main_data = load_data(uploaded_file)
        cleaned_data = process_data(main_data)
        outbound_orders_data = cleaned_data[cleaned_data['Consignment Number'].str.endswith('-OUT', na=False)]
        inwards_lcl_data = cleaned_data[cleaned_data['Consignment Number'].str.endswith('-IN', na=False)]
        adhoc_data = cleaned_data[cleaned_data['Type'] == 'Ad-hoc']

        outbound_orders_output = create_order_output_data(outbound_orders_data)
        st.write("Example: Outbound Orders Data:")
        st.dataframe(outbound_orders_output.head())

        inwards_lcl_output = create_inwards_lcl_output_data(inwards_lcl_data)
        st.write("Example: Inwards LCL Data:")
        st.dataframe(inwards_lcl_output.head())

        adhoc_output = create_ad_hoc_output_data(adhoc_data)
        st.write("Example: Ad-Hoc Data:")
        st.dataframe(adhoc_output.head())

        # Perform checks
        unique_source, unique_outbound, unique_inwards, unique_adhoc = check_unique_consignments(cleaned_data, outbound_orders_output, inwards_lcl_output, adhoc_output)
        total_source, total_outbound, total_inwards, total_adhoc, total_output = check_subtotals(cleaned_data, outbound_orders_output, inwards_lcl_output, adhoc_output)

        st.write(f"Unique Consignments in Source Data: {unique_source}")
        st.write(f"Unique Consignments in Outbound: {unique_outbound}")
        st.write(f"Unique Consignments in Inwards: {unique_inwards}")
        st.write(f"Unique Consignments in Ad-Hoc: {unique_adhoc}")

        st.write(f"Total Subtotal in Source: {total_source}")
        st.write(f"Total Subtotal in Outbound: {total_outbound}")
        st.write(f"Total Subtotal in Inwards: {total_inwards}")
        st.write(f"Total Subtotal in Ad-Hoc: {total_adhoc}")

        st.write(f"Total Subtotal in all Output: {total_output}")
        st.write(f"Subtotals Match: {total_source == total_output}")

        output_excel = to_excel(outbound_orders_output, inwards_lcl_output, adhoc_output)
        st.download_button(label="Download Output", data=output_excel, file_name="B360-Output.xlsx")

if __name__ == "__main__":
    main()
