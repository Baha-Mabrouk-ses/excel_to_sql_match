import argparse
import json
import unicodedata
import pandas as pd
from sentence_transformers import SentenceTransformer, util

# Load the multilingual model
model = SentenceTransformer('paraphrase-multilingual-MiniLM-L12-v2')

def normalize_text(text):
    """Normalize text by converting to lowercase and removing accents."""
    text = text.lower()
    return ''.join(
        char for char in unicodedata.normalize('NFD', text) if unicodedata.category(char) != 'Mn'
    )

def match_columns_and_output_excel(excel_file, database_content, output_excel):
    # Extract database column names dynamically from the JSON
    database_columns = list(database_content[0].keys())
    database_columns_normalized = [normalize_text(col) for col in database_columns]

    # Read the Excel file
    df = pd.read_excel(excel_file)
    excel_columns = [normalize_text(col) for col in df.columns.tolist()]

    # Encode the normalized column names
    db_embeddings = model.encode(database_columns_normalized, convert_to_tensor=True)
    excel_embeddings = model.encode(excel_columns, convert_to_tensor=True)

    # Match columns
    matches = {}
    unmatched_columns = []
    for i, excel_col in enumerate(excel_columns):
        similarities = util.pytorch_cos_sim(excel_embeddings[i], db_embeddings)
        best_match_idx = similarities.argmax()
        if similarities[0, best_match_idx] > 0.7:  # Threshold for similarity
            matches[df.columns[i]] = database_columns[best_match_idx]
        else:
            unmatched_columns.append(df.columns[i])

    # Create a DataFrame for output
    matched_df = pd.DataFrame(database_content)[list(matches.values())]
    matched_df.rename(columns={v: k for k, v in matches.items()}, inplace=True)

    # Save to Excel
    matched_df.to_excel(output_excel, index=False)
    print(f"Matched data saved to {output_excel}")

def main():
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Match columns and output database content to Excel (French compatibility)")
    parser.add_argument("excel_file", type=str, help="Path to the Excel file")
    parser.add_argument("database_content", type=str, help="Path to a JSON file containing database content as a list of dictionaries")
    parser.add_argument("--output_excel", type=str, default="output.xlsx", help="Path to save the output Excel file (default: output.xlsx)")

    # Parse the arguments
    args = parser.parse_args()
    excel_file = args.excel_file
    database_content_file = args.database_content
    output_excel = args.output_excel

    # Load database content
    try:
        with open(database_content_file, "r", encoding="utf-8") as f:
            database_content = json.load(f)
    except Exception as e:
        print(f"Error loading database content: {str(e)}")
        return

    # Match columns and output to Excel
    try:
        match_columns_and_output_excel(excel_file, database_content, output_excel)
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
