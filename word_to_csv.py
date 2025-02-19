import os
import csv
import logging
from docx import Document

logger = logging.getLogger(__name__)

def read_word_file(input_docx_file):
    doc = Document(input_docx_file)
    text = []
    for para in doc.paragraphs:
        if para.text.strip():
            text.append(para.text)
    return text

def parse_data_to_csv(input_docx_file):
    try:
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
                    logger.warning(f"Skipping invalid row: {line}")

        logger.info(f"CSV file '{output_csv_file}' has been created successfully.")
        return output_csv_file
    except Exception as e:
        logger.error(f"An error occurred while parsing the document: {str(e)}")
        raise