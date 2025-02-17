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