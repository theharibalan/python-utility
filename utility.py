import pandas as pd
from datetime import datetime
import os

class ExcelDeltaComparator:
    def __init__(self, input_file):
        self.input_file = input_file
        self.df = pd.read_excel(input_file, skiprows=[1])  # Skip second row
        self.output_rows = []

    def safe_str(self, val):
        return '' if pd.isna(val) else str(val).strip()

    def process(self):
        column_names = self.df.columns.tolist()
        base_columns = []
        delta_columns = {}

        # Identify base columns and delta columns
        for col in column_names:
            if "(1)" in str(col):
                base = col.split(" (1)")[0].strip()
                base_columns.append(base)

            elif "delta" in str(col).lower():
                key = str(col).lower().replace("delta", "").replace("(1 - 2)", "").strip()
                delta_columns[key] = col

        for _, row in self.df.iterrows():
            position_id = row.get('position_id', None)
            if pd.isna(position_id):  # Skip row without position_id
                continue

            output_row = {'position_id': position_id}

            for base in base_columns:
                col1 = f"{base} (1)"
                col2 = f"{base} (2)"
                delta_col = delta_columns.get(base.lower(), f"delta (1 - 2)")

                val1 = self.safe_str(row.get(col1, ''))
                val2 = self.safe_str(row.get(col2, ''))
                delta_val = self.safe_str(row.get(delta_col, ''))

                if delta_val.upper() == 'DIFF':
                    # Only add if at least one value is non-empty
                    if val1 or val2:
                        output_row[base] = f"{val1} | {val2}"

            if len(output_row) > 1:  # Skip if only position_id exists
                self.output_rows.append(output_row)

    def save_output(self):
        if not self.output_rows:
            print("No differences found. No report generated.")
            return

        now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_file = f"Result_Report_{now}.xlsx"
        output_df = pd.DataFrame(self.output_rows)
        output_df.to_excel(output_file, index=False)
        print(f"✅ Output file generated: {output_file}")

# Example usage
if __name__ == "__main__":
    input_path = "your_input_file.xlsx"  # Replace with your file
    comparator = ExcelDeltaComparator(input_path)
    comparator.process()
    comparator.save_output()
