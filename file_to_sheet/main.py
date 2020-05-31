import openpyxl
from openpyxl.styles import Font, Alignment

cell_size = {}

def remember_cell_size(source_sheet, cell):
    col_name = cell[0]
    row_name = int(cell[1])
    cell_size[col_name] = source_sheet.column_dimensions[col_name].width
    cell_size[row_name] = source_sheet.row_dimensions[row_name].height


def source_to_target(in_path, in_file, target_wb):
    target_wb.create_sheet(title=in_file.split('.')[0], index=0)

    source_wb = openpyxl.load_workbook(f'{in_path}/{in_file}')
    source_sheet = source_wb.active
    target_sheet = target_wb.active

    for row in source_sheet:
        for cell in row:
            remember_cell_size(source_sheet, cell.coordinate)
            font = Font(name=cell.font.name,
                        size=cell.font.size,
                        italic=cell.font.i,
                        bold=cell.font.b,
                        color=cell.font.color)
            alignment = Alignment(horizontal=cell.alignment.horizontal,
                                  vertical=cell.alignment.vertical,
                                  wrapText=cell.alignment.wrapText)
            target_sheet[cell.coordinate].font = font
            target_sheet[cell.coordinate].alignment = alignment
            target_sheet[cell.coordinate] = cell.value

    for name, value in cell_size.items():
        if isinstance(name, str):
            target_sheet.column_dimensions[name].width = value
        else:
            target_sheet.row_dimensions[name].height = value

    for merge_cell in source_sheet.merged_cells:
        target_sheet.merge_cells(str(merge_cell))


def main():
    in_files = ['alcohol.xlsx', 'cigarettes.xlsx']
    in_path = 'in'
    out_file = 'out/result_book.xlsx'


    target_wb = openpyxl.Workbook()

    for in_file in in_files:
        source_to_target(in_path, in_file, target_wb)

    target_wb.remove(target_wb['Sheet'])
    target_wb.save(out_file)


if __name__ == '__main__':
    main()