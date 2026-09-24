from openpyxl import load_workbook

from .excel_utils import find_column, norm


def get_terminated_employee_ids(
    worksheet,
    head_count_employee_id_header,
    head_count_status_header,
    header_row=2,
):
    employee_id_col = find_column(worksheet, head_count_employee_id_header, header_row)
    status_col = find_column(worksheet, head_count_status_header, header_row)

    missing = [
        name
        for name, col in [
            (head_count_employee_id_header, employee_id_col),
            (head_count_status_header, status_col),
        ]
        if not col
    ]
    if missing:
        raise ValueError("Missing Head Count column(s): " + ", ".join(missing))

    terminated_ids = set()
    terminated_count = 0

    for row in range(header_row + 1, worksheet.max_row + 1):
        if norm(worksheet.cell(row, status_col).value) == "terminated":
            terminated_count += 1
            employee_id = norm(worksheet.cell(row, employee_id_col).value)
            if employee_id:
                terminated_ids.add(employee_id)

    return terminated_ids, terminated_count


def process_files(
    head_count_path,
    classic_user_path,
    output_path,
    head_count_sheet_name,
    classic_user_sheet_name,
    head_count_employee_id_header,
    head_count_status_header,
    classic_user_employee_id_header,
    classic_user_id_header,
):
    head_count_workbook = load_workbook(head_count_path, data_only=False)
    classic_user_workbook = load_workbook(classic_user_path, data_only=False)

    if head_count_sheet_name not in head_count_workbook.sheetnames:
        raise ValueError(
            f"Head Count worksheet '{head_count_sheet_name}' was not found."
        )
    if classic_user_sheet_name not in classic_user_workbook.sheetnames:
        raise ValueError(
            f"Classic User worksheet '{classic_user_sheet_name}' was not found."
        )

    head_count_worksheet = head_count_workbook[head_count_sheet_name]
    classic_user_worksheet = classic_user_workbook[classic_user_sheet_name]

    terminated_ids, terminated_count = get_terminated_employee_ids(
        head_count_worksheet,
        head_count_employee_id_header,
        head_count_status_header,
        header_row=2,
    )

    classic_user_employee_id_col = find_column(
        classic_user_worksheet, classic_user_employee_id_header
    )
    classic_user_id_col = find_column(classic_user_worksheet, classic_user_id_header)
    missing = [
        name
        for name, col in [
            (classic_user_employee_id_header, classic_user_employee_id_col),
            (classic_user_id_header, classic_user_id_col),
        ]
        if not col
    ]
    if missing:
        raise ValueError("Missing Classic User column(s): " + ", ".join(missing))

    classic_user_rows_by_employee_id = {}
    for row in range(2, classic_user_worksheet.max_row + 1):
        employee_id_value = norm(
            classic_user_worksheet.cell(row, classic_user_employee_id_col).value
        )
        if employee_id_value in terminated_ids:
            classic_user_rows_by_employee_id.setdefault(employee_id_value, []).append(row)

    rows_to_delete = []
    rows_to_retain = []
    for rows in classic_user_rows_by_employee_id.values():
        rows_to_retain.append(rows[-1])
        rows_to_delete.extend(rows[:-1])

    for row in sorted(rows_to_delete, reverse=True):
        classic_user_worksheet.delete_rows(row, 1)

    for original_row in rows_to_retain:
        new_row = original_row - sum(
            deleted_row < original_row for deleted_row in rows_to_delete
        )
        for column in range(1, classic_user_worksheet.max_column + 1):
            if column != classic_user_id_col:
                classic_user_worksheet.cell(new_row, column).value = None

    if "Processing Log" in classic_user_workbook.sheetnames:
        del classic_user_workbook["Processing Log"]
    log = classic_user_workbook.create_sheet("Processing Log")
    log.append(["Metric", "Value"])

    matched_employees = len(classic_user_rows_by_employee_id)
    not_found = max(len(terminated_ids) - matched_employees, 0)
    for metric, value in [
        ("Terminated employees in Head Count", terminated_count),
        ("Terminated employees found in Classic User", matched_employees),
        ("Classic User rows deleted", len(rows_to_delete)),
        ("Classic User rows retained", len(rows_to_retain)),
        ("Terminated employees not found in Classic User", not_found),
        (
            "Retained row behavior",
            f"{classic_user_id_header} kept, all other columns cleared",
        ),
    ]:
        log.append([metric, value])

    classic_user_workbook.save(output_path)
    return {
        "terminated": terminated_count,
        "matched": matched_employees,
        "deleted": len(rows_to_delete),
        "blanked": len(rows_to_retain),
        "not_found": not_found,
    }
