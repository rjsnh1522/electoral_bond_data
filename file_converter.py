import PyPDF2
import pandas as pd
import re

def get_first_int_character_position(s):
    for i, char in enumerate(s):
        if char.isdigit():
            return i
    return -1  # Return -1 if no integer character is found

def separate_amount_from_name(string):
    # Start from the end of the string
    index = len(string) - 1
    amount = ""

    # Find the first space going backwards
    while index >= 0 and string[index] != ' ':
        # Ignore commas
        if string[index] != ',':
            amount = string[index] + amount
        index -= 1

    # After finding the space, the remaining part is the name
    name = string[:index].strip()

    return name, amount


def parse_line(line):
    # Define regular expressions for the patterns
    pattern1 = r'^.{1,11}'
    pattern2 = r'(\d.*)'
    pattern3 = r'(?<=\d)[^\d]*$'

    # Apply regular expressions to split the string
    part1 = line[:11]
    rest = line[11:]
    name, amount = separate_amount_from_name(rest)
    
    return part1, name, amount



def merge_company_name(parsed_data):
    merged_data = []
    company_name = ''
    for item in parsed_data[5:]:
        if item.isdigit() and len(item) >= 4:
            # If the item is a number with 4 or more digits, merge it with the company name
            company_name += ' ' + item
        else:
            # If company_name already contains text, append it to merged_data
            if company_name.strip() != '':
                merged_data.append(company_name.strip())
                company_name = ''  # Reset company_name
            if item == 'TL':
                # Include "TL" in the company name
                company_name += ' ' + item
            else:
                merged_data.append(item)
    if company_name.strip() != '':
        merged_data.append(company_name.strip())
    return merged_data



def data_parse_with_regex(line):
    # pattern = r'^(\d+)\s+(\d+)\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+([\S\s]+?)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\d+)\s+([\d,]+)\s+(\d+)\s+(\d+)\s+(\w+)$'
    # pattern = r'^(\d+)\s+(\d+)\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+(\d{2}/[A-Za-z]{3}/\d{4})\s+([\w\s]+?)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\w+)\s+(\d+)\s+([\d,]+)\s+(\d+)\s+(\d+)\s+(\w+)$'
    pattern = r'^\d+\s+\d+\s+\d{2}\/[A-Za-z]{3}\/\d{4}\s+\d{2}\/[A-Za-z]{3}\/\d{4}\s+\d{2}\/[A-Za-z]{3}\/\d{4}\s+[A-Z\s]+OC\s+\d+\s+\d{1,3}(,\d{3})*\s+\d+\s+\d+\s+\w+$'

    matches = re.match(pattern, line)
    if matches:
        parsed_data = matches.groups()
        print(parsed_data)
        corrected_data = merge_company_name(parsed_data=parsed_data)
        print(corrected_data)

    else:
        print("No match found.")
        print(line)
        print("**"*10)



def extract_table_from_pdf(pdf_path, start_page, end_page):
    table_data = []
    skip_these = ["Sr No. Reference No  (URN) Journal DateDate of ", "PurchaseDate of Expiry Name of the Purchaser PrefixBond ", "NumberDenominations Issue Branch Code Issue Teller Status"]

    with open(pdf_path, 'rb') as pdf_file:
        pdf_reader = PyPDF2.PdfReader(pdf_file)
        for page in pdf_reader.pages:

            text = page.extract_text()
            
            # Split text into lines and extract table-like data
            lines = text.split('\n')
            for line in lines:
                if line not in skip_these:
                    if len(line) >= 6:  # Assuming at least 6 characters for each line
                        table_data.append(data_parse_with_regex(line))
            
    return table_data

def save_to_excel(table_data, output_path):
    df = pd.DataFrame(table_data, columns=['Column1', 'Column2', 'Column3'])
    df.to_excel(output_path, index=False)

def convert_pdf_to_excel(pdf_path, start_page, end_page, output_path):
    table_data = extract_table_from_pdf(pdf_path, start_page, end_page)
    save_to_excel(table_data, output_path)

# Example usage
pdf_path = 'bond_buyer.pdf'  # Path to your PDF file
start_page = 1  # Start page of the table
end_page = 500  # End page of the table
output_path = 'bond_buyer-part1.xlsx'  # Path where the Excel file will be saved

convert_pdf_to_excel(pdf_path, start_page, end_page, output_path)