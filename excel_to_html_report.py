import sys
from xlsx2html import xlsx2html

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Excel Report</title>
    <link rel="stylesheet" href="https://cdn.datatables.net/1.13.6/css/jquery.dataTables.min.css">
    <style>
        body {{
            font-family: Arial, sans-serif;
            padding: 20px;
        }}
        table {{
            width: 100%;
            border-collapse: collapse;
        }}
        table.dataTable thead th {{
            background-color: #f2f2f2;
        }}
    </style>
</head>
<body>

<h2>Excel Report Viewer</h2>
<div class="table-container">
    {table_content}
</div>

<script src="https://code.jquery.com/jquery-3.7.0.js"></script>
<script src="https://cdn.datatables.net/1.13.6/js/jquery.dataTables.min.js"></script>
<script>
    $(document).ready(function () {{
        $('table').DataTable({{
            "paging": true,
            "searching": true,
            "ordering": true
        }});
    }});
</script>

</body>
</html>
"""

def convert_excel_to_html(input_excel, output_html):
    try:
        # Generate raw HTML from Excel
        html_table = xlsx2html(input_excel, output=False)
        
        # Wrap with template
        final_html = HTML_TEMPLATE.format(table_content=html_table)
        
        # Write to output file
        with open(output_html, "w", encoding="utf-8") as f:
            f.write(final_html)

        print(f"[✅] HTML report generated at: {output_html}")
    except Exception as e:
        print(f"[❌] Error: {e}")

if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python excel_to_html_report.py <input.xlsx> <output.html>")
    else:
        convert_excel_to_html(sys.argv[1], sys.argv[2])
