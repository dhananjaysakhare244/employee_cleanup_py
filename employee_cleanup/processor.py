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

    first_col = min(employee_id_col, status_col)
    employee_id_index = employee_id_col - first_col
    status_index = status_col - first_col

    for values in worksheet.iter_rows(
        min_row=header_row + 1,
        min_col=first_col,
        max_col=max(employee_id_col, status_col),
        values_only=True,
    ):
        if norm(values[status_index]) == "terminated":
            terminated_count += 1
            employee_id = norm(values[employee_id_index])
            if employee_id:
                terminated_ids.add(employee_id)

    return terminated_ids, terminated_count


def delete_rows_in_blocks(worksheet, rows_to_delete):
    if not rows_to_delete:
        return []

    sorted_rows = sorted(rows_to_delete)
    row_blocks = []
    block_start = previous_row = sorted_rows[0]

    for row in sorted_rows[1:]:
        if row == previous_row + 1:
            previous_row = row
            continue
        row_blocks.append((block_start, previous_row))
        block_start = previous_row = row

    row_blocks.append((block_start, previous_row))
    for start_row, end_row in reversed(row_blocks):
        worksheet.delete_rows(start_row, end_row - start_row + 1)
    return sorted_rows


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
    head_count_workbook = load_workbook(
        head_count_path,
        read_only=True,
        data_only=False,
    )
    try:
        if head_count_sheet_name not in head_count_workbook.sheetnames:
            raise ValueError(
                f"Head Count worksheet '{head_count_sheet_name}' was not found."
            )

        head_count_worksheet = head_count_workbook[head_count_sheet_name]
        terminated_ids, terminated_count = get_terminated_employee_ids(
            head_count_worksheet,
            head_count_employee_id_header,
            head_count_status_header,
            header_row=2,
        )
    finally:
        head_count_workbook.close()

    classic_user_workbook = load_workbook(classic_user_path, data_only=False)
    if classic_user_sheet_name not in classic_user_workbook.sheetnames:
        raise ValueError(
            f"Classic User worksheet '{classic_user_sheet_name}' was not found."
        )

    classic_user_worksheet = classic_user_workbook[classic_user_sheet_name]

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

    last_row_by_employee_id = {}
    rows_to_delete = []
    for row, (value,) in enumerate(
        classic_user_worksheet.iter_rows(
            min_row=2,
            min_col=classic_user_employee_id_col,
            max_col=classic_user_employee_id_col,
            values_only=True,
        ),
        start=2,
    ):
        employee_id_value = norm(value)
        if employee_id_value in terminated_ids:
            previous_row = last_row_by_employee_id.get(employee_id_value)
            if previous_row is not None:
                rows_to_delete.append(previous_row)
            last_row_by_employee_id[employee_id_value] = row

    rows_to_retain = sorted(last_row_by_employee_id.values())

    rows_to_delete = delete_rows_in_blocks(classic_user_worksheet, rows_to_delete)

    deleted_row_index = 0
    deleted_rows_before = 0
    for original_row in rows_to_retain:
        while (
            deleted_row_index < len(rows_to_delete)
            and rows_to_delete[deleted_row_index] < original_row
        ):
            deleted_rows_before += 1
            deleted_row_index += 1
        new_row = original_row - deleted_rows_before
        for column in range(1, classic_user_worksheet.max_column + 1):
            if column != classic_user_id_col:
                classic_user_worksheet.cell(new_row, column).value = None

    if "Processing Log" in classic_user_workbook.sheetnames:
        del classic_user_workbook["Processing Log"]
    log = classic_user_workbook.create_sheet("Processing Log")
    log.append(["Metric", "Value"])

    matched_employees = len(last_row_by_employee_id)
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
