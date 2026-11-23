TARGET_HEADERS = {
    "nik", "no nik", "no. nik", "no kk", "no. kk", "nomor kk",
    "nomor kartu keluarga",
}
CHARACTER = "\u200c"


def detect_columns(workbook):
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows():
            matches = tuple(
                cell for cell in row
                if cell.data_type == "s" and isinstance(cell.value, str)
                and " ".join(cell.value.casefold().split()) in TARGET_HEADERS
            )
            if matches:
                yield sheet, row[0].row, matches
                break


def clean_columns(targets, *, identifier_format=True):
    count = 0
    for sheet, header_row, headers in targets:
        for header in headers:
            for row in range(header_row + 1, sheet.max_row + 1):
                cell = sheet.cell(row, header.column)
                if cell.data_type != "s" or not isinstance(cell.value, str):
                    continue
                if CHARACTER in cell.value:
                    cell.value = cell.value.replace(CHARACTER, "")
                    # Keep text literal even if removal exposes a leading '='.
                    cell.data_type = "s"
                    if identifier_format:
                        cell.number_format = "@"
                        cell.quotePrefix = True
                    count += 1
    return count


def save_new(workbook, output):
    with output.open("xb") as stream:
        try:
            workbook.save(stream)
        except (Exception, KeyboardInterrupt):
            stream.close()
            output.unlink(missing_ok=True)
            raise
