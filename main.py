import pandas as pd

def parse_text_file(file_path):
    with open(file_path, encoding='utf-8') as f:
        lines = f.read().splitlines()
    tables, current_table = [], None
    for line in lines:
        line = line.strip()
        if not line:
            current_table = None
            continue
        if '#' not in line:
            current_table = {"name": line, "header": None, "rows": []}
            tables.append(current_table)
        else:
            header_part, data_part = [p.strip() for p in line.split('#', 1)]
            cols = [x.strip() for x in header_part.split('&')]
            vals = [x.strip() for x in data_part.split('&')]
            if current_table is None:
                current_table = {"name": "Без названия", "header": cols, "rows": []}
                tables.append(current_table)
            if current_table["header"] is None:
                current_table["header"] = cols
            current_table["rows"].append(vals)
    return tables

def write_tables_to_excel(tables, output_file):
    sheet_data = []
    for table in tables:
        sheet_data.append([table["name"]])
        sheet_data.append(table["header"] if table["header"] else [])
        sheet_data.extend(table["rows"])
        sheet_data.append([])

    max_cols = max((len(row) for row in sheet_data), default=0)
    normalized = [row + [""] * (max_cols - len(row)) for row in sheet_data]
    df = pd.DataFrame(normalized)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Tables", header=False, index=False)
    print(f"Excel файл сохранён как '{output_file}'")

if __name__ == "__main__":
    tables = parse_text_file("results.txt")
    write_tables_to_excel(tables, "output.xlsx")