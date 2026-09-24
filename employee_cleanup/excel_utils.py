def norm(value):
    return "" if value is None else str(value).strip().lower()


def find_column(ws, name, header_row=1):
    target = norm(name)
    for cell in ws[header_row]:
        if norm(cell.value) == target:
            return cell.column
    return None
