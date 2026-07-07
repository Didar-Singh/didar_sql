import pandas as pd

input_file = "Objects_1000125_export Part 2.dat"
output_file = "first_10_rows2.xlsx"

delimiter = "þ\x14þ"


with open(input_file, "r", encoding="utf-8-sig", errors="ignore") as f:

    header = f.readline().strip()

    headers = [h.strip("þ ") for h in header.split(delimiter)]

    records = []

    for _ in range(10):
        line = f.readline()

        if not line:
            break

        values = [v.strip("þ \r\n") for v in line.split(delimiter)]

        if len(values) < len(headers):
            values += [""] * (len(headers) - len(values))

        records.append(values[:len(headers)])

df = pd.DataFrame(records, columns=headers)

df.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print(f"Excel file created: {output_file}")
print(f"Rows exported: {len(df)}")
print(f"Columns exported: {len(df.columns)}")