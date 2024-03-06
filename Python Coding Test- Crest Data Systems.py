#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
from docx import Document
import fitz


def extract_text_from_pdf(pdf_path):
    """Extract text from PDF."""
    text = ''
    with fitz.open(pdf_path) as doc:
        for page in doc:
            text += page.get_text()
    return text


def extract_text_from_doc(doc_path):
    """Extract text from DOCX."""
    doc = Document(doc_path)
    text = ''
    for para in doc.paragraphs:
        text += para.text + '\n'
    return text


def extract_text_from_dat(dat_path):
    """Extract text from DAT file."""
    with open(dat_path, 'r') as file:
        text = file.read()
    return text


def convert_to_csv(text, csv_path):
    """Convert extracted text to CSV."""
    # Split the text into lines
    lines = text.strip().split('\n')
    # Extract headings from the first line
    headings = lines[0].split('\t')
    # Extract data from the remaining lines
    data = [line.split('\t') for line in lines[1:] if line.strip()]
    # Create a DataFrame from the data
    df = pd.DataFrame(data, columns=headings)

    # Drop duplicate rows
    df.drop_duplicates(inplace=True)

    # Convert numeric columns to numbers
    numeric_columns = ['id', 'basic_salary', 'allowances']
    df[numeric_columns] = df[numeric_columns].apply(pd.to_numeric, errors='ignore')

    # Add a new column for Gross Salary
    df['Gross Salary'] = df['basic_salary'] + df['allowances']

    # Calculate average salary and round to two decimal places
    average_salary = round(df['Gross Salary'].mean(), 2)

    # Find the second highest salary
    second_highest_salary = df['Gross Salary'].nlargest(2).iloc[-1]

    # Insert additional information as new rows below the first and second columns
    additional_info = [
        ['Second Highest Salary=' + str(second_highest_salary), 'average salary = ' + str(average_salary)]
    ]
    df_additional_info = pd.DataFrame(additional_info, columns=[headings[0], headings[1]])

    # Concatenate the additional information DataFrame with the main DataFrame
    df = pd.concat([df, df_additional_info], ignore_index=True)

    # Save the DataFrame to a CSV file
    df.to_csv(csv_path, index=False)


def convert_document_to_csv(doc_path, csv_path):
    """Convert document to CSV."""
    if doc_path.lower().endswith('.pdf'):
        text = extract_text_from_pdf(doc_path)
    elif doc_path.lower().endswith('.docx'):
        text = extract_text_from_doc(doc_path)
    elif doc_path.lower().endswith('.dat'):
        text = extract_text_from_dat(doc_path)
    else:
        print("Unsupported document type.")
        return

    if text:
        convert_to_csv(text, csv_path)
        print(f"CSV file created successfully at {csv_path}")
    else:
        print("Failed to extract text from the document.")


# Example usage:
input_document_path = r'C:\Users\A200232677\Desktop\Personal\DATA.dat'  # Provide path to your input document
output_csv_path = r'C:\Users\A200232677\Desktop\Personal\Data_Output2.csv'  # Provide path to your output file with name eg.Data_Output.csv

convert_document_to_csv(input_document_path, output_csv_path)


