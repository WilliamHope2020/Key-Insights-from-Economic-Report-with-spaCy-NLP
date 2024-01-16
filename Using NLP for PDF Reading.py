import spacy
import re
import pandas as pd
import fitz

# Load spaCy English model
nlp = spacy.load('en_core_web_sm')

# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_path):
    text = ""
    with fitz.open(pdf_path) as pdf_document:
        for page_number in range(pdf_document.page_count):
            page = pdf_document[page_number]
            text += page.get_text()
    return text

# Function to create an Excel file with extracted statistics
def create_excel_with_statistics(output_excel_path, statistics):
    # Create a DataFrame with the statistics
    df = pd.DataFrame({'Statistics': statistics})
    
    # Write the DataFrame to an Excel file
    df.to_excel(output_excel_path, index=False)

# Replace 'input.pdf' with the path to your PDF file
input_pdf_path = "C:\\Users\\savag\\Downloads\\tech_profile_report.pdf"

# Extract text from the PDF
pdf_text = extract_text_from_pdf(input_pdf_path)

# Process the text using spaCy
doc = nlp(pdf_text)

# Define a regular expression for capturing statistics (including percentages)
statistics_pattern = re.compile(r'\b([^\d]*(\d+(\.\d+)?)%[^\w\d]+?[^\d]+.*?)\b')

# Function to extract statistics based on regular expressions
def extract_statistics(text, pattern):
    matches = re.findall(pattern, text)
    return [match[0] for match in matches]

# Extract statistics
statistics = extract_statistics(pdf_text, statistics_pattern)

# Write the statistics to an Excel file
output_excel_path = "output_results.xlsx"
create_excel_with_statistics(output_excel_path, statistics)

print(f"Results written to {output_excel_path}")
