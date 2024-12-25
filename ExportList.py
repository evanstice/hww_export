# Company data provided by Poteto Copyright (c) 2017 - https://github.com/poteto/hiring-without-whiteboards/

import requests
import csv
import pandas as pd


def parse_csv(raw_data):
    company = ""
    link = ""
    location = ""
    info = ""

    csv_data = [["Company", "Link", "Location", "Info"]]
    for line in raw_data:
        count = 0
        last_count = 0
        while line[count] != "]":
            count += 1
        company = line[3:count]

        count += 2
        last_count = count

        while line[count] != ")":
            count += 1
        link = line[last_count:count]

        count += 4
        last_count = count

        try:
            while line[count] != "|":
                count += 1
            location = line[last_count:count - 1]

            count += 2

            info = line[count:]
        except IndexError:
            location = line[last_count:]
            info = ""

        csv_data.append([company, link, location, info])

    with open("data.csv", "w", newline="") as file:
        writer = csv.writer(file)
        writer.writerows(csv_data)


def main():
    # Fetch data from the URL
    url = "https://raw.githubusercontent.com/poteto/hiring-without-whiteboards/refs/heads/main/README.md"

    response = requests.get(url)
    response.raise_for_status()

    content = response.text.splitlines()

    # Remove extraneous lines
    while content:
        if content[0] == "---":
            break
        content.pop(0)
    content.pop(0)
    data = []

    for line in content:
        if line.startswith("-"):
            data.append(line)

    # Parse the raw data into CSV format
    parse_csv(data)

    # Write Excel file

    df = pd.read_csv("data.csv")

    output_file = "output.xlsx"
    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Internships", index=False, startcol=1)  # Shift table to column B

        workbook = writer.book
        worksheet = writer.sheets["Internships"]

        # Set column widths
        worksheet.set_column("A:A", 3)  # Checkboxes
        worksheet.set_column("B:B", 30)  # Company
        worksheet.set_column("C:C", 30)  # Link
        worksheet.set_column("D:D", 30)  # Location
        worksheet.set_column("E:E", 100)  # Info

        # Formatting
        table_range = f"A1:E{len(df) + 1}"
        worksheet.add_table(
            table_range,
            {
                "header_row": True,
                "style": "Table Style Medium 7",
                "columns": [{"header": col} for col in ["☑", *df.columns]],
            },
        )

        # Write hyperlinks
        for row_num in range(1, len(df) + 1):
            link = df.iloc[row_num - 1]["Link"]
            if link:
                worksheet.write_url(f"C{row_num + 1}", link)

if __name__ == "__main__":
    main()
