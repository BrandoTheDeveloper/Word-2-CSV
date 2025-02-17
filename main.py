import os
import csv
from docx import Document

def read_word_file(input_docx_file):
    doc = Document(input_docx_file)
    text = []
    for para in doc.paragraphs:
        if para.text.strip():
            text.append(para.text)
    return text

def parse_data_to_csv(input_docx_file):
    data = read_word_file(input_docx_file)
    output_csv_file = os.path.splitext(input_docx_file)[0] + '.csv'

    with open(output_csv_file, mode='w', newline='', encoding='utf-8') as file:
        writer = csv.writer(file)

        header = ["Potential Surplus", "Est. Resale Value", "Opening Bid", "Date Sold", "Case #",
                  "Parcel ID", "Type of Foreclosure", "First Name", "Last Name", "Mailing Address",
                  "Mailing City", "Mailing State", "Mailing Zip Code", "Property Address",
                  "Property City", "Property State", "Property Zip Code", "County"]
        writer.writerow(header)

        for line in data:
            row = [item.strip() for item in line.split(',')]
            if len(row) == len(header):
                writer.writerow(row)
            else:
                print(f"Skipping invalid row: {line}")

    return output_csv_file

if __name__ == '__main__':
    # Ask the user for the file path
    docx_file = input("Enter the path to the Word document: ")

    # Check if the provided file path exists
    if not os.path.exists(docx_file):
        print(f"Error: The file '{docx_file}' was not found.")
    # Check if the provided file has a valid extension
    elif not docx_file.lower().endswith('.docx'):
        print(f"Error: The file '{docx_file}' is not a valid Word document (.docx).")
    else:
        try:
            csv_file = parse_data_to_csv(docx_file)
            print(f"CSV file '{csv_file}' has been created successfully.")
        except PermissionError:
            print(f"Error: Permission denied when trying to create the CSV file.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")