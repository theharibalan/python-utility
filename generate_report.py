import os
from xlsx_to_html_extractor.excel_reader import ExcelReader
from jinja2 import Environment, FileSystemLoader

# Config
input_excel = "sample_report.xlsx"
output_html = "output/report.html"
sheet_name = None  # Set to None to auto-detect first sheet
template_dir = "templates"

# Step 1: Extract table HTML from Excel
excel_reader = ExcelReader(input_excel)
table_html = excel_reader.get_sheet_html(sheet_name)

# Step 2: Load wrapper template (for DataTables, styles)
env = Environment(loader=FileSystemLoader(template_dir))
template = env.get_template("layout.html")
final_html = template.render(table_html=table_html)

# Step 3: Save HTML
os.makedirs("output", exist_ok=True)
with open(output_html, "w", encoding="utf-8") as f:
    f.write(final_html)

print(f"✅ Report generated: {os.path.abspath(output_html)}")
