# README: Key-Value System for Parsing Manual Fields from BibTeX to Excel

This document defines a key-value system for filling the "manual" fields in the Excel file using the `annote` field from the BibTeX entries. The `annote` field is expected to contain information in a structured format that maps keys to their corresponding values.

## Key-Value System Format
The `annote` field in BibTeX will contain key-value pairs in the following format:

```
key1: value1;
key2: value2;
key3: value3;
...
```

### Example
For example, the `annote` field might look like this:

```
ResearchProblem: This study addresses a gap in understanding consumer behavior;
ResearchQuestion: How do brand perceptions influence purchasing decisions?;
Hypothesis: Positive brand perception increases likelihood of purchase;
Method: Survey-based analysis with statistical modeling;
SampleSize: 500 respondents;
MainResults: Significant correlation between brand perception and purchase likelihood;
Conclusions: Enhancing brand perception can lead to increased sales.
```

## Supported Keys
Below is the list of supported keys and their corresponding Excel columns:

| Key               | Excel Column                       |
|-------------------|------------------------------------|
| `ResearchProblem` | Research Problem/Gap              |
| `ResearchQuestion`| Research Question(s)              |
| `Hypothesis`      | Hypothesis(es)                    |
| `Method`          | Method(s)                         |
| `SampleSize`      | Sample Size                       |
| `MainResults`     | Main Results                      |
| `Conclusions`     | Conclusions                       |

### Notes:
1. Keys in the `annote` field are **case-insensitive** (e.g., `ResearchProblem` and `researchproblem` are treated the same).
2. Each key-value pair must end with a semicolon (`;`).
3. Any unsupported keys will be ignored during processing.

## Parsing Workflow
1. Check if the `annote` field exists in the BibTeX entry.
2. Extract the key-value pairs from the `annote` field.
3. Map the extracted values to their corresponding Excel columns using the key-value mapping.
4. Populate the Excel file with the extracted data.

## Error Handling
1. If the `annote` field is empty or missing, the corresponding manual fields in the Excel file will remain blank.
2. If a key does not have a value (e.g., `ResearchQuestion: ;`), the corresponding Excel column will remain blank.
3. If the `annote` field has formatting errors (e.g., missing semicolons), the entry will be skipped with a warning logged.

## Example Mapping
### Input BibTeX Entry:
```
@article{example,
    title={A Study on Consumer Behavior},
    author={John Doe},
    year={2023},
    journal={Journal of Marketing Research},
    annote={
        ResearchProblem: Understanding consumer behavior in the digital age;
        ResearchQuestion: What factors influence online purchasing?;
        Method: Qualitative interviews;
        SampleSize: 30 participants;
        Conclusions: Personalization drives online sales.;
    }
}
```

### Output in Excel:
| Column                     | Value                                      |
|----------------------------|--------------------------------------------|
| Research Problem/Gap       | Understanding consumer behavior in the digital age |
| Research Question(s)       | What factors influence online purchasing? |
| Hypothesis(es)             |                                            |
| Theoretical Model / Framework |                                       |
| Method(s)                  | Qualitative interviews                    |
| Sample Size                | 30 participants                           |
| Main Results               |                                            |
| Conclusions                | Personalization drives online sales.      |

## Implementation Notes
The parsing logic will be added to the `parse_bibtex_to_excel` function in the Python script. The function will:
1. Check for the presence of the `annote` field.
2. Parse the field using the defined key-value system.
3. Map the parsed values to the corresponding columns in the Excel file.

## Future Extensions
1. Add support for additional keys if required.
2. Handle multiline values more effectively.
3. Provide a configuration file to customize key mappings.

