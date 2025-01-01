import bibtexparser
import pandas as pd
import re

# Define the output columns
COLUMNS = [
    "Year",
    "Journal",
    "VHB / SJR / CiteScore Rank",
    "Title",
    "Author(s)",
    "Research Problem/Gap",
    "Research Question(s)",
    "Hypothesis(es)",
    "Theorectical Model / Framework",
    "Method(s)",
    "Sample Size",
    "Main Results",
    "Conclusions",
    "DOI / Link to article"
]

def clean_text(text):
    """Remove unnecessary curly braces from text."""
    if text:
        return re.sub(r"\{(.*?)\}", r"\1", text)
    return text

def parse_bibtex_to_excel(input_bibtex, output_excel):
    # Read and parse the BibTeX file
    with open(input_bibtex, 'r') as bibtex_file:
        bib_database = bibtexparser.load(bibtex_file)

    # Prepare the data for Excel
    rows = []
    for entry in bib_database.entries:
        row = {
            "Year": clean_text(entry.get("year", "")),
            "Journal": clean_text(entry.get("journal", "")),
            "VHB / SJR / CiteScore Rank": "",  # Manually added later
            "Title": clean_text(entry.get("title", "")),
            "Author(s)": clean_text(entry.get("author", "")),
            "Research Problem/Gap": "",  # Manually added later
            "Research Question(s)": "",  # Manually added later
            "Hypothesis(es)": "",  # Manually added later
            "Theorectical Model / Framework": "",  # Manually added later
            "Method(s)": "",  # Manually added later
            "Sample Size": "",  # Manually added later
            "Main Results": "",  # Manually added later
            "Conclusions": "",  # Manually added later
            "DOI / Link to article": clean_text(entry.get("doi", ""))
        }
        rows.append(row)

    # Create a DataFrame from the rows
    df = pd.DataFrame(rows, columns=COLUMNS)

    # Write the DataFrame to an Excel file
    df.to_excel(output_excel, index=False)

if __name__ == "__main__":
    # Input BibTeX file path
    input_bibtex = "references.bib"  # Replace with your BibTeX file path

    # Output Excel file path
    output_excel = "output.xlsx"  # Replace with your desired Excel file name

    # Convert BibTeX to Excel
    parse_bibtex_to_excel(input_bibtex, output_excel)

    print(f"Conversion complete! Excel file saved as: {output_excel}")
