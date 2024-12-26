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

def classify_and_generate_sql(excel_file, database_columns, database_table):
    # Normalize database column names
    database_columns = [normalize_text(col) for col in database_columns]

    # Read the Excel file
    df = pd.read_excel(excel_file)
    excel_columns = [normalize_text(col) for col in df.columns.tolist()]

    # Encode the normalized column names
    db_embeddings = model.encode(database_columns, convert_to_tensor=True)
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

    # Create a DataFrame for SQL insertion
    matched_df = df[list(matches.keys())].rename(columns=matches)
    matched_df['other_information'] = df[unmatched_columns].to_dict(orient='records')

    # Generate SQL insert statements
    sql_statements = []
    for _, row in matched_df.iterrows():
        columns = ', '.join(row.index)
        values = ', '.join([
            f"'{json.dumps(value, default=str)}'"  # Use default=str to handle Timestamps
            if isinstance(value, dict) else
            f"'{str(value)}'"  # Convert other values to string
            for value in row
        ])
        sql_statements.append(f"INSERT INTO {database_table} ({columns}) VALUES ({values});")

    return "\n".join(sql_statements)

def main():
    # Set up argument parsing
    parser = argparse.ArgumentParser(description="Generate SQL script from Excel file (French compatibility)")
    parser.add_argument("excel_file", type=str, help="Path to the Excel file")
    parser.add_argument("database_columns", type=str, help="Comma-separated list of database columns")
    parser.add_argument("database_table", type=str, help="Name of the database table")
    parser.add_argument("--output", type=str, default="output.sql", help="Path to save the generated SQL script (default: output.sql)")

    # Parse the arguments
    args = parser.parse_args()
    excel_file = args.excel_file
    database_columns = args.database_columns.split(",")
    database_table = args.database_table
    output_file = args.output

    # Generate SQL script
    try:
        sql_script = classify_and_generate_sql(excel_file, database_columns, database_table)
        with open(output_file, "w", encoding="utf-8") as f:
            f.write(sql_script)
        print(f"SQL script generated successfully and saved to {output_file}")
    except Exception as e:
        print(f"Error: {str(e)}")

if __name__ == "__main__":
    main()
