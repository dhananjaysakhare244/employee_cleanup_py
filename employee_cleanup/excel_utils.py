def norm(value):
    return "" if value is None else str(value).strip().lower()


def find_column(ws, name):
    target = norm(name)
    for cell in ws[1]:
        if norm(cell.value) == target:
            return cell.column
    return None

