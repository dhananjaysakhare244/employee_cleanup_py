from openpyxl import load_workbook

from .excel_utils import find_column, norm


def get_terminated_employee_ids(ws, employee_id_column, status_name, header_row=2):
    employee_id_col = find_column(ws, employee_id_column, header_row)
    status_col = find_column(ws, status_name, header_row)

    missing = [
        name
        for name, col in [
            (employee_id_column, employee_id_col),
            (status_name, status_col),
        ]
        if not col
    ]
    if missing:
        raise ValueError("Missing employee-file column(s): " + ", ".join(missing))

    terminated_ids = set()
    terminated_count = 0

    for row in range(header_row + 1, ws.max_row + 1):
        if norm(ws.cell(row, status_col).value) == "terminated":
            terminated_count += 1
            employee_id = norm(ws.cell(row, employee_id_col).value)
            if employee_id:
                terminated_ids.add(employee_id)

    return terminated_ids, terminated_count


def process_files(
    employee_file,
    records_file,
    output_file,
    employee_sheet,
    records_sheet,
    employee_id_column,
    employee_status,
    records_employee_id_column,
    records_id,
):
    employee_wb = load_workbook(employee_file, data_only=False)
    records_wb = load_workbook(records_file, data_only=False)

    if employee_sheet not in employee_wb.sheetnames:
        raise ValueError(f"Employee sheet '{employee_sheet}' was not found.")
    if records_sheet not in records_wb.sheetnames:
        raise ValueError(f"Records sheet '{records_sheet}' was not found.")

    employee_ws = employee_wb[employee_sheet]
    records_ws = records_wb[records_sheet]

    terminated_ids, terminated_count = get_terminated_employee_ids(
        employee_ws, employee_id_column, employee_status, header_row=2
    )

    records_employee_id_col = find_column(records_ws, records_employee_id_column)
    records_id_col = find_column(records_ws, records_id)
    missing = [
        name
        for name, col in [
            (records_employee_id_column, records_employee_id_col),
            (records_id, records_id_col),
        ]
        if not col
    ]
    if missing:
        raise ValueError("Missing records-file column(s): " + ", ".join(missing))

    employee_rows = {}
    for row in range(2, records_ws.max_row + 1):
        employee_id_value = norm(
            records_ws.cell(row, records_employee_id_col).value
        )
        if employee_id_value in terminated_ids:
            employee_rows.setdefault(employee_id_value, []).append(row)

    rows_to_delete = []
    rows_to_retain = []
    for rows in employee_rows.values():
        rows_to_retain.append(rows[-1])
        rows_to_delete.extend(rows[:-1])

    for row in sorted(rows_to_delete, reverse=True):
        records_ws.delete_rows(row, 1)

    for original_row in rows_to_retain:
        new_row = original_row - sum(
            deleted_row < original_row for deleted_row in rows_to_delete
        )
        for column in range(1, records_ws.max_column + 1):
            if column != records_id_col:
                records_ws.cell(new_row, column).value = None

    if "Processing Log" in records_wb.sheetnames:
        del records_wb["Processing Log"]
    log = records_wb.create_sheet("Processing Log")
    log.append(["Metric", "Value"])

    matched_employees = len(employee_rows)
    not_found = max(len(terminated_ids) - matched_employees, 0)
    for metric, value in [
        ("Terminated employees in employee file", terminated_count),
        ("Terminated employees found in records", matched_employees),
        ("Records deleted", len(rows_to_delete)),
        ("Records retained", len(rows_to_retain)),
        ("Terminated employees not found in records", not_found),
        ("Retained row behavior", f"{records_id} kept, all other columns cleared"),
    ]:
        log.append([metric, value])

    records_wb.save(output_file)
    return {
        "terminated": terminated_count,
        "matched": matched_employees,
        "deleted": len(rows_to_delete),
        "blanked": len(rows_to_retain),
        "not_found": not_found,
    }
