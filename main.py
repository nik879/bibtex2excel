import bibtexparser
import pandas as pd
import re
import requests
import os
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# Retrieve the API key from environment variables
SCOPUS_API_KEY = os.getenv("SCOPUS_API_KEY")
if not SCOPUS_API_KEY:
    raise ValueError("SCOPUS_API_KEY is not set. Please add it to the .env file.")

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

def parse_annote_field(annote):
    """Parse the annote field into a dictionary of key-value pairs."""
    if not annote:
        return {}

    # Split the annote field by semicolons and extract key-value pairs
    pairs = re.split(r";\s*", annote)
    annote_dict = {}
    for pair in pairs:
        match = re.match(r"(\w+):\s*(.+)", pair)
        if match:
            key, value = match.groups()
            annote_dict[key.strip().lower()] = value.strip()
    return annote_dict

def clean_text(text):
    """Remove unnecessary curly braces from text."""
    if text:
        return re.sub(r"\{(.*?)\}", r"\1", text)
    return text

def get_issn_from_doi(doi):
    """Retrieve the ISSN of a journal using the DOI via the Scopus API."""
    if not doi:
        return ""
    url = f"https://api.elsevier.com/content/article/doi/{doi}"
    headers = {"X-ELS-APIKey": SCOPUS_API_KEY, "Accept": "application/json"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        coredata = data.get("full-text-retrieval-response", {}).get("coredata", {})
        return coredata.get("prism:issn", "")
    else:
        print(f"Failed to retrieve ISSN for DOI {doi}: {response.status_code}, {response.text}")
        return ""

def get_issn_and_citescore_by_title(journal_title):
    """Retrieve the ISSN and CiteScore of a journal by its title via the Scopus API."""
    if not journal_title:
        return "", ""
    url = f"https://api.elsevier.com/content/serial/title?title={journal_title}"
    headers = {"X-ELS-APIKey": SCOPUS_API_KEY, "Accept": "application/json"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        entries = data.get("serial-metadata-response", {}).get("entry", [])
        if entries:
            issn = entries[0].get("prism:issn", "")
            citescore = (
                entries[0].get("citeScoreYearInfoList", {})
                .get("citeScoreCurrentMetric", "")
            )
            return issn, citescore
    return "", ""

def get_citescore(journal_issn):
    """Retrieve the CiteScore of a journal using the Scopus API."""
    if not journal_issn:
        return ""
    url = f"https://api.elsevier.com/content/serial/title/issn/{journal_issn}"
    headers = {"X-ELS-APIKey": SCOPUS_API_KEY, "Accept": "application/json"}
    response = requests.get(url, headers=headers)

    if response.status_code == 200:
        data = response.json()
        return (
            data.get("serial-metadata-response", {})
            .get("entry", [{}])[0]
            .get("citeScoreYearInfoList", {})
            .get("citeScoreCurrentMetric", "")
        )
    else:
        return ""


def parse_bibtex_to_excel(input_bibtex, output_excel):
    """Convert BibTeX file to Excel with additional data."""
    # Read and parse the BibTeX file
    with open(input_bibtex, 'r') as bibtex_file:
        bib_database = bibtexparser.load(bibtex_file)

    # Prepare the data for Excel
    rows = []
    for entry in bib_database.entries:
        journal = clean_text(entry.get("journal", ""))
        doi = clean_text(entry.get("doi", ""))

        # Retrieve ISSN and CiteScore
        issn = get_issn_from_doi(doi) if doi else ""
        citescore = get_citescore(issn) if issn else ""

        if not issn or not citescore:
            issn, citescore = get_issn_and_citescore_by_title(journal)

        # Parse the annote field
        annote_field = clean_text(entry.get("annote", ""))
        annote_data = parse_annote_field(annote_field)

        row = {
            "Year": clean_text(entry.get("year", "")),
            "Journal": journal,
            "VHB / SJR / CiteScore Rank": citescore,
            "Title": clean_text(entry.get("title", "")),
            "Author(s)": clean_text(entry.get("author", "")),
            "Research Problem/Gap": annote_data.get("researchproblem", ""),
            "Research Question(s)": annote_data.get("researchquestion", ""),
            "Hypothesis(es)": annote_data.get("hypothesis", ""),
            "Theorectical Model / Framework": annote_data.get("theoreticalmodel", ""),
            "Method(s)": annote_data.get("method", ""),
            "Sample Size": annote_data.get("samplesize", ""),
            "Main Results": annote_data.get("mainresults", ""),
            "Conclusions": annote_data.get("conclusions", ""),
            "DOI / Link to article": doi
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
