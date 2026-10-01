import asyncio
import csv
from io import BytesIO

import flet as ft

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from src import PyME_functions as CC #for Render, which runs from root, so src is in the path
except ImportError:
    try:
        import PyME_functions as CC
    except ImportError as e:
        print(f"Failed to import PyME_functions: {e}")
        CC = None

try:
    from docx import Document
except ImportError:
    Document = None


def stream_to_bytes(stream):
    if isinstance(stream, bytes):
        return stream
    if hasattr(stream, "getvalue"):
        return stream.getvalue()
    stream.seek(0)
    return stream.read()


def main(page: ft.Page):
    page.title = "PyME - Method Comparison"
    if not page.web:
        page.window.maximized = True

    analyte = ft.TextField(label="Measurand", value="Glucose",)
    unit = ft.TextField(label="Unit", value="g/L")
    
    error1 = ft.TextField(label="Absolute Error", value="0.2")
    error2 = ft.TextField(label="% Error", value="10")
    cutoff = ft.TextField(label="Cut-off for Error", value="2")

    hist_img = ft.Image(src=b"", visible=False, width=900)
    deming_img = ft.Image(src=b"", visible=False, width=600)
    status = ft.Text("Enter paired values in the spreadsheet, then select Generate.")
    progress = ft.ProgressBar(value=None, visible=False)
    generate_button = ft.Button(content="Generate", disabled=True)
    clipboard = ft.Clipboard()
    data_store = {"generating": False, "active_cell": (0, 0)}
    spreadsheet_rows = []
    initial_row_count = 20

    def update_generate_enabled():
        has_values = any(
            (cell.value or "").strip()
            for row in spreadsheet_rows
            for cell in row
        )
        generate_button.disabled = not has_values or data_store["generating"]
        page.update()

    def set_active_cell(row_index, column_index):
        def on_focus():
            data_store["active_cell"] = (row_index, column_index)
        return on_focus

    def add_spreadsheet_rows(count):
        start_index = len(spreadsheet_rows)
        for row_index in range(start_index, start_index + count):
            row_cells = []
            for column_index in range(2):
                row_cells.append(
                    ft.TextField(
                        value="",
                        width=190,
                        on_focus=set_active_cell(row_index, column_index),
                        on_change=update_generate_enabled,
                    )
                )
            spreadsheet_rows.append(row_cells)
            spreadsheet_grid.controls.append(
                ft.Row(
                    spacing=0,
                    controls=[
                        ft.Container(
                            width=44,
                            alignment=ft.Alignment.CENTER,
                            border=ft.Border.all(1, ft.Colors.OUTLINE),
                            content=ft.Text(str(row_index + 1)),
                        ),
                        *row_cells,
                    ],
                )
            )

    async def paste_from_clipboard():
        pasted = await clipboard.get()
        if not pasted or not pasted.strip():
            status.value = "Copy two adjacent numeric columns, then choose Paste from Excel."
            page.update()
            return

        try:
            matrix = [
                row
                for row in csv.reader(pasted.splitlines(), delimiter="\t")
                if any(value.strip() for value in row)
            ]
        except csv.Error as exc:
            status.value = f"Could not read clipboard data: {exc}"
            page.update()
            return

        start_row, start_column = data_store["active_cell"]
        needed_rows = start_row + len(matrix) - len(spreadsheet_rows)
        if needed_rows > 0:
            add_spreadsheet_rows(needed_rows)

        for row_offset, values in enumerate(matrix):
            target_row = spreadsheet_rows[start_row + row_offset]
            for column_offset, value in enumerate(values[: 2 - start_column]):
                target_row[start_column + column_offset].value = value.strip()

        data_store["active_cell"] = (start_row, start_column)
        update_generate_enabled()
        status.value = (
            f"Pasted {len(matrix)} spreadsheet rows starting at row {start_row + 1}. "
            "Reference (X) is the first column; Test (Y) is the second."
        )
        page.update()

    def add_more_rows():
        add_spreadsheet_rows(10)
        page.update()

    spreadsheet_grid = ft.ListView(
        height=360,
        spacing=0,
        scroll=ft.ScrollMode.AUTO,
        controls=[],
    )
    add_spreadsheet_rows(initial_row_count)

    async def show_progress(message):
        status.value = message
        progress.visible = True
        progress.value = None
        page.update()
        await asyncio.sleep(0.05)

    def read_spreadsheet_data():
        x_values = []
        y_values = []
        for row_index, row in enumerate(spreadsheet_rows, start=1):
            x_text = (row[0].value or "").strip()
            y_text = (row[1].value or "").strip()
            if not x_text and not y_text:
                continue
            if not x_text or not y_text:
                raise ValueError(f"Row {row_index} needs both X and Y values.")
            try:
                x_values.append(float(x_text))
                y_values.append(float(y_text))
            except ValueError as exc:
                raise ValueError(f"Row {row_index} contains a non-numeric value.") from exc

        if not x_values:
            raise ValueError("Enter or paste at least one paired X and Y row.")
        return pd.DataFrame({"Reference": x_values, "Test": y_values})

    async def generate():
        if data_store["generating"]:
            return
        if pd is None:
            status.value = "Data processing is unavailable: pandas could not be imported."
            page.update()
            return
        if CC is None:
            status.value = (
                "PyME_functions is not available. Add PyME_functions.py and its "
                "dependencies to the app before generating plots and the report."
            )
            page.update()
            return
        if Document is None:
            status.value = "The python-docx package is unavailable."
            page.update()
            return

        try:
            df = read_spreadsheet_data()
            cutoff_value = float(cutoff.value)
            error1_value = float(error1.value)
            error2_value = float(error2.value)
        except (ValueError, TypeError) as exc:
            status.value = f"Check the spreadsheet and numeric settings: {exc}"
            page.update()
            return

        data_store["generating"] = True
        generate_button.disabled = True
        paste_button.disabled = True
        add_rows_button.disabled = True
        progress.visible = True
        progress.value = None
        page.update()
        await asyncio.sleep(0.05)

        try:
            x = df["Reference"].copy()
            y = df["Test"].copy()
            x.name = f"{analyte.value} Ref."
            y.name = f"{analyte.value} Test"

            await show_progress("Assessing measurement distribution…")
            hist_stream = await asyncio.to_thread(CC.Histogram_grouped, x, y)
            hist_img.src = stream_to_bytes(hist_stream)
            hist_img.visible = True
            page.update()

            await show_progress("Calculating Deming regression…")
            deming_result = await asyncio.to_thread(
                CC.Deming_Plot_Equal_Variance_with_Error2_PX,
                x,
                y,
                Error_level_cut_off=cutoff_value,
                error1=error1_value,
                error2=error2_value,
            )
            deming_stream, n, slope, intercept, r = deming_result
            deming_img.src = stream_to_bytes(deming_stream)
            deming_img.visible = True
            page.update()

            await show_progress("Building the Word report…")
            report = await asyncio.to_thread(
                CC.MC_output,
                document=Document(),
                analyte=analyte.value,
                x=x,
                y=y,
                Unit=unit.value,
                Error_level_cut_off=cutoff_value,
                error1=error1_value,
                error2=error2_value,
            )
            report_buffer = BytesIO()
            await asyncio.to_thread(report.save, report_buffer)

            progress.visible = False
            status.value = "Report ready. Choose where to save the DOCX file."
            page.update()
            saved_path = await file_picker.save_file(
                dialog_title="Save method comparison report",
                file_name="MethodComparison.docx",
                src_bytes=report_buffer.getvalue(),
            )
            if saved_path:
                status.value = (
                    f"Report saved to {saved_path}. Generated comparison for {n} samples. "
                    f"Slope: {slope:.3f}; intercept: {intercept:.3f}; r: {r:.3f}."
                )
            else:
                status.value = "Report generated, but saving was canceled."
        except (ValueError, TypeError) as exc:
            status.value = f"Check the spreadsheet and numeric inputs: {exc}"
        except Exception as exc:
            status.value = f"Could not generate the comparison: {exc}"
        finally:
            progress.visible = False
            paste_button.disabled = False
            add_rows_button.disabled = False
            data_store["generating"] = False
            update_generate_enabled()

    paste_button = ft.Button(
        content="Paste from Excel clipboard",
        on_click=paste_from_clipboard,
    )
    add_rows_button = ft.TextButton(content="Add 10 rows", on_click=add_more_rows)
    generate_button.on_click = generate
    file_picker = ft.FilePicker()

    page.add(
        ft.Column(
            expand=True,
            scroll=ft.ScrollMode.AUTO,
            controls=[
                ft.Row(
                    scroll=ft.ScrollMode.AUTO,
                    controls=[
                        ft.Column(
                            controls=[
                                analyte,
                                unit,
                                error1,
                                error2,
                                cutoff,
                                ft.Text(
                                    "Spreadsheet input: copy two adjacent numeric columns from Excel, "
                                    "then focus a starting cell and choose Paste from Excel clipboard."
                                ),
                                ft.Row(
                                    spacing=0,
                                    controls=[
                                        ft.Container(
                                            width=44,
                                            alignment=ft.Alignment.CENTER,
                                            border=ft.Border.all(1, ft.Colors.OUTLINE),
                                            content=ft.Text("#"),
                                        ),
                                        ft.Container(
                                            width=190,
                                            alignment=ft.Alignment.CENTER,
                                            border=ft.Border.all(1, ft.Colors.OUTLINE),
                                            content=ft.Text("Reference (X)"),
                                        ),
                                        ft.Container(
                                            width=190,
                                            alignment=ft.Alignment.CENTER,
                                            border=ft.Border.all(1, ft.Colors.OUTLINE),
                                            content=ft.Text("Test (Y)"),
                                        ),
                                    ],
                                ),
                                spreadsheet_grid,
                                ft.Row(controls=[paste_button, add_rows_button]),
                                generate_button,
                                progress,
                                status,
                            ]
                        ),
                        ft.Column(controls=[hist_img, deming_img]),
                    ],
                    vertical_alignment=ft.CrossAxisAlignment.START,
                )
            ],
        )
    )


ft.run(main)
