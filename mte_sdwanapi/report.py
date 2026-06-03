"""Excel report generation helpers for SD-WAN data."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from mte_sdwanapi.api import parse_feature_templates_for_device_template, tprint


def save_report_to_excel(
    destination_file,
    reachable_devices,
    device_template_definitions,
    feature_template_definitions,
    policy_definitions,
    attached_device_values,
):
    """Build and save the Excel report workbook."""

    if not reachable_devices:
        tprint("No reachable devices found, Excel report was not created.")
        return

    workbook = Workbook()
    reachable_devices_worksheet = workbook.active
    reachable_devices_worksheet.title = "reachable devices"

    headers = list(reachable_devices[0].keys())

    header_fill = PatternFill(
        start_color="1F4E78", end_color="1F4E78", fill_type="solid"
    )
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, header in enumerate(headers, start=1):
        cell = reachable_devices_worksheet.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    for row_idx, device in enumerate(reachable_devices, start=2):
        for col_idx, header in enumerate(headers, start=1):
            reachable_devices_worksheet.cell(
                row=row_idx, column=col_idx, value=device.get(header)
            )

    for col_idx, header in enumerate(headers, start=1):
        max_data_length = max(
            len(str(device.get(header, ""))) for device in reachable_devices
        )
        reachable_devices_worksheet.column_dimensions[
            reachable_devices_worksheet.cell(row=1, column=col_idx).column_letter
        ].width = max(len(header), max_data_length) + 2

    reachable_devices_worksheet.freeze_panes = "A2"

    feature_template_lookup = {}
    for feature_template_definition in feature_template_definitions:
        template_id = feature_template_definition.get("templateId")
        if template_id:
            feature_template_lookup[template_id] = feature_template_definition

    policy_lookup = {}
    for policy_definition in policy_definitions:
        policy_id = policy_definition.get("policyId")
        if policy_id:
            policy_lookup[policy_id] = policy_definition

    attached_values_lookup = {}
    for template_data in attached_device_values:
        template_id = template_data.get("templateId")
        if template_id:
            attached_values_lookup[template_id] = template_data.get(
                "attachedValues", []
            )

    used_sheet_names = {reachable_devices_worksheet.title}
    for idx, device_template in enumerate(device_template_definitions, start=1):
        template_name = device_template.get("templateName") or f"device_template_{idx}"
        sheet_name = template_name[:31]
        if sheet_name in used_sheet_names:
            sheet_name = f"Template {idx}"[:31]
        used_sheet_names.add(sheet_name)

        worksheet = workbook.create_sheet(title=sheet_name)

        summary_headers = ["Field", "Value"]
        for col_idx, header in enumerate(summary_headers, start=1):
            cell = worksheet.cell(row=1, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        summary_rows = [
            ("Template Name", device_template.get("templateName", "")),
            ("Template ID", device_template.get("templateId", "")),
            (
                "Template Description",
                device_template.get("templateDescription", ""),
            ),
            ("Policy ID", device_template.get("policyId", "")),
            (
                "Policy Name",
                policy_lookup.get(device_template.get("policyId"), {}).get(
                    "policyName", "NOT_FOUND"
                ),
            ),
        ]

        for row_idx, row_data in enumerate(summary_rows, start=2):
            worksheet.cell(row=row_idx, column=1, value=row_data[0])
            worksheet.cell(row=row_idx, column=2, value=row_data[1])

        # Keep tables aligned by using a fixed starting row on each sheet.
        feature_table_header_row = 8
        feature_headers = [
            "templateType",
            "templateName",
            "templateDescription",
            "templateId",
        ]
        for col_idx, header in enumerate(feature_headers, start=1):
            cell = worksheet.cell(row=feature_table_header_row, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        template_feature_templates = parse_feature_templates_for_device_template(
            device_template
        )
        for row_idx, template_feature in enumerate(
            template_feature_templates, start=feature_table_header_row + 1
        ):
            template_id = template_feature.get("templateId")
            feature_definition = feature_template_lookup.get(template_id, {})
            worksheet.cell(
                row=row_idx,
                column=1,
                value=template_feature.get("templateType", ""),
            )
            worksheet.cell(
                row=row_idx,
                column=2,
                value=feature_definition.get("templateName", "NOT_FOUND"),
            )
            worksheet.cell(
                row=row_idx,
                column=3,
                value=feature_definition.get("templateDescription", "NOT_FOUND"),
            )
            worksheet.cell(row=row_idx, column=4, value=template_id)

        # Start attached-variable columns to the right of the feature table.
        attached_table_start_col = 9
        attached_table_header_row = 8
        attached_rows = attached_values_lookup.get(device_template.get("templateId"), [])

        preferred_attached_headers = [
            "csv-deviceIP",
            "csv-deviceId",
            "csv-host-name",
        ]

        discovered_headers = []
        for row_data in attached_rows:
            for key in row_data.keys():
                if key not in discovered_headers:
                    discovered_headers.append(key)

        attached_headers = [
            header for header in preferred_attached_headers if header in discovered_headers
        ]
        attached_headers.extend(
            sorted([header for header in discovered_headers if header not in attached_headers])
        )

        if not attached_headers:
            attached_headers = ["attached-values"]

        for col_offset, header in enumerate(attached_headers):
            cell = worksheet.cell(
                row=attached_table_header_row,
                column=attached_table_start_col + col_offset,
                value=header,
            )
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        if attached_rows:
            for row_idx, row_data in enumerate(attached_rows, start=attached_table_header_row + 1):
                for col_offset, header in enumerate(attached_headers):
                    value = row_data.get(header)
                    if isinstance(value, list):
                        value = ", ".join(str(item) for item in value)
                    worksheet.cell(
                        row=row_idx,
                        column=attached_table_start_col + col_offset,
                        value=value,
                    )
        else:
            worksheet.cell(
                row=attached_table_header_row + 1,
                column=attached_table_start_col,
                value="No attached devices/variables",
            )

        worksheet.column_dimensions["A"].width = 24
        worksheet.column_dimensions["B"].width = 50
        worksheet.column_dimensions["C"].width = 60
        worksheet.column_dimensions["D"].width = 40
        for col_offset, header in enumerate(attached_headers):
            col_letter = worksheet.cell(
                row=attached_table_header_row,
                column=attached_table_start_col + col_offset,
            ).column_letter
            if attached_rows:
                max_data_len = max(
                    len(
                        str(
                            (
                                row_data.get(header)
                                if not isinstance(row_data.get(header), list)
                                else ", ".join(
                                    str(item) for item in row_data.get(header, [])
                                )
                            )
                            or ""
                        )
                    )
                    for row_data in attached_rows
                )
            else:
                max_data_len = len("No attached devices/variables") if col_offset == 0 else 0
            worksheet.column_dimensions[col_letter].width = max(len(header), max_data_len) + 2

    workbook.save(destination_file)
    tprint(f"{destination_file} saved!")
