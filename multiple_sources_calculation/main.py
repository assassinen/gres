import openpyxl
from openpyxl.styles import Font, Alignment
import random

cell_size = {}

def remember_cell_size(source_sheet, cell, max_row=0):
    col_index = cell[0]
    row_index = int(cell[1])
    cell_size[col_index] = source_sheet.column_dimensions[col_index].width
    cell_size[int(row_index + max_row)] = source_sheet.row_dimensions[row_index].height
    return f'{col_index}{int(row_index + max_row)}'


def source_to_target(in_path, in_file, target_wb, template_sheets):
    source_wb = openpyxl.load_workbook(f'{in_path}/{in_file}')

    for sheet in template_sheets:
        try:
            source_sheet = source_wb[sheet]
        except:
            continue
        target_sheet = target_wb[sheet]

        max_row = 0 if in_file == '_template.xlsx' else target_sheet.max_row + 1

        for row in source_sheet:
            for cell in row:
                coordinate = remember_cell_size(source_sheet, cell.coordinate, max_row)
                font = Font(name=cell.font.name,
                            size=cell.font.size,
                            italic=cell.font.i,
                            bold=cell.font.b,
                            color=cell.font.color)
                alignment = Alignment(horizontal=cell.alignment.horizontal,
                                      vertical=cell.alignment.vertical,
                                      wrapText=cell.alignment.wrapText)
                target_sheet[coordinate].font = font
                target_sheet[coordinate].alignment = alignment

                target_sheet[coordinate] = cell.value

        for name, value in cell_size.items():
            if isinstance(name, str):
                target_sheet.column_dimensions[name].width = value
            else:
                target_sheet.row_dimensions[name].height = value

        for merge_cell in source_sheet.merged_cells:
            merge = [f'{i[0]}{int(i[1:])+max_row}' for i in str(merge_cell).split(':')]
            target_sheet.merge_cells(':'.join(merge))


def main():
    in_files = ['tula_book.xlsx', 'orel_book.xlsx']
    in_path = 'in'
    out_file = 'out/result_book.xlsx'
    skip_row = 3

    template_wb = openpyxl.load_workbook(f'{in_path}/_template.xlsx')
    template_sheets = [worksheets.title for worksheets in template_wb.worksheets]

    target_wb = openpyxl.Workbook()
    for sheet in template_sheets:
        target_wb.create_sheet(title=sheet, index=0)

    source_to_target(in_path, '_template.xlsx', target_wb, template_sheets)

    for in_file in in_files:
        source_to_target(in_path, in_file, target_wb, template_sheets)

    target_wb.remove(target_wb['Sheet'])
    target_wb.save(out_file)


if __name__ == '__main__':
    main()