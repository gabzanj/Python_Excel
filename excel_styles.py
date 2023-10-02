import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Border, Side


def create_dataframe():
    df_women_inventors = {
        'Inventor Name': ['Hedy Lamarr', 'Marie Curie', 'Grace Hopper', 'Rosalind Franklin', 'Ada Lovelace'],
        'Invention': ['Secret Communication System', 'X-Rays', 'Programming Compiler', 'DNA Structure',
                      'First Algorithm'],
        'Year of Invention': [1941, 1895, 1952, 1952, 1843],
        'Field of Endeavor': ['Technology', 'Science', 'Technology', 'Science', 'Technology'],
        'Contribution': ['Developed a secret communication system for the U.S. Navy',
                         'Researched and developed the X-ray technique',
                         'Created the first compiler for the COBOL language',
                         'Contributed to the understanding of DNA structure',
                         'Wrote the first algorithm for Charles Babbage\'s analytical engine']
    }

    return pd.DataFrame(df_women_inventors)


def add_borders_to_range(worksheet):
    border_style = Side(style='thin')
    border = Border(
        left=border_style,
        right=border_style,
        top=border_style,
        bottom=border_style,
    )

    for row in worksheet.iter_rows():
        for cell in row:
            cell.border = border


def save_dataframe_to_excel_file(dataframe, excel_filename):
    workbook = Workbook()
    worksheet = workbook.active

    header_row = dataframe.columns.tolist()
    worksheet.append(header_row)
    header_font = Font(name='Arial', size=12, bold=True)
    for cell in worksheet[1]:
        cell.font = header_font

    for row_idx, row_data in dataframe.iterrows():
        row_values = row_data.tolist()
        worksheet.append(row_values)

    add_borders_to_range(worksheet)
    workbook.save(excel_filename)


def main():
    print("Routine is starting...")
    df_women_inventors = create_dataframe()

    excel_filename = 'women_inventors.xlsx'
    save_dataframe_to_excel_file(df_women_inventors, excel_filename)
    print(f"DataFrame saved successfully to '{excel_filename}'.")


if __name__ == '__main__':
    main()
