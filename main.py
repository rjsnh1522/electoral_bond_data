import pandas as pd
from bound_buyer_file_converter import convert_buyer
from bound_user_file_converter import convert_user


def merge_both_excel_sheet(buyer_file, user_file):
    # Read the Excel sheets into DataFrames
    df1 = pd.read_excel(buyer_file)
    df2 = pd.read_excel(user_file)

    # Merge the DataFrames based on two common columns
    merged_df = pd.merge(df1, df2, on=['Prefix', 'Bond Number'])

    # Write the merged DataFrame to a new Excel file
    merged_df.to_excel('merged_buyer_and_user.xlsx', index=False)


def converter(buyer_output, user_output):
    try:
        convert_buyer(buyer_output)
        convert_user(user_output)
    except Exception as e:
        print("Some error Occured")




def build_markdown():
   # Read the merged Excel data into a Pandas DataFrame
    merged_df = pd.read_excel('merged_buyer_and_user.xlsx')

    # Extract the year from the relevant date column
    merged_df['Year'] = pd.to_datetime(merged_df['Date of Purchase']).dt.year
    merged_df['Denominations_x'] = merged_df['Denominations_x'].str.replace(',', '').astype(float)

    # Group by year and buyer, summing up the bond amounts
    grouped_df = merged_df.groupby(['Year', 'Name of the Purchaser']).agg({'Denominations_x': 'sum'}).reset_index()

    # Rename the columns for clarity
    grouped_df.rename(columns={'Denominations_x': 'Total Amount'}, inplace=True)
        
    # # Table 2: Mean Purchaser Amount by Year, Sorted Descendingly
    # merged_df["Encashment"] = pd.to_datetime(merged_df["Date Of Encashment"]).dt.year
    # merged_df['Denominations_y'] = merged_df['Denominations_y'].str.replace(',', '').astype(float)
    # mean_purchaser_amount = merged_df.groupby(['Encashment', 'Name of Political party'])['Denominations_y'].mean().reset_index()
    # # mean_purchaser_amount = mean_purchaser_amount.sort_values(by=['Date Of Encashment', 'Denominations_y'], ascending=[True, False])
    # most_frequent_purchaser = most_frequent_purchaser.sort_values(by=['Encashment', 'Denominations_y'], ascending=[True, False])
    # most_frequent_purchaser = most_frequent_purchaser.groupby('Date of Purchase').first().reset_index()[['Date of Purchase', 'Name of Political party']]

    # Write tables to separate Markdown files
    # most_frequent_purchaser.to_markdown('most_frequent_purchaser.md', index=False)
    grouped_df.to_markdown('most_frequent_purchaser.md', index=False)
    # mean_purchaser_amount.to_markdown('mean_purchaser_amount.md', index=False)


def main():
    buyer = "bound_buyer.xlsx"
    user = "bound_user.xlsx"
    print("Converting file from pdf to excel to do some magic")
    converter(buyer, user)
    print("File conversion complete........")
    print("Merging both the files On Prefix and Bond Number ..................")
    merge_both_excel_sheet(buyer, user)
    print("File merger complete.....................................")
    print("Check file merged_buyer_and_user to see merged data")
    




if __name__ == "__main__":
    # main()
    build_markdown()