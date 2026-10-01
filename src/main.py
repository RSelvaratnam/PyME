import asyncio
from io import BytesIO

import flet as ft

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from src import PyME_functions as CC
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

    analyte = ft.TextField(label="Measurand", value="Digoxin")
    unit = ft.TextField(label="Unit", value="nmol/L")
    cutoff = ft.TextField(label="Cut-off", value="2")
    error1 = ft.TextField(label="Absolute error", value="0.2")
    error2 = ft.TextField(label="% error", value="10")

    hist_img = ft.Image(src=b"", visible=False, width=900)
    deming_img = ft.Image(src=b"", visible=False, width=600)
    status = ft.Text("Choose an Excel workbook to begin.")
    progress = ft.ProgressBar(value=None, visible=False)
    pick_button = ft.Button(content="Pick Excel")
    generate_button = ft.Button(content="Generate", disabled=True)
    data_store = {"data": None, "generating": False}
    file_picker = ft.FilePicker()

    async def pick_excel(_):
        if pd is None:
            status.value = "Excel support is unavailable: pandas could not be imported."
            page.update()
            return

        files = await file_picker.pick_files(
            dialog_title="Choose an Excel workbook",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            allow_multiple=False,
            with_data=True,
        )
        if not files:
            return

        selected = files[0]
        if selected.bytes is None:
            status.value = "The workbook could not be read. Please choose it again."
            page.update()
            return

        try:
            data_store["data"] = pd.read_excel(
                BytesIO(selected.bytes),
                sheet_name="Accuracy",
                usecols="B:G",
                skiprows=range(0, 16),
                engine="openpyxl",
            )
            generate_button.disabled = False
            status.value = f"Loaded {selected.name}. Ready to generate the comparison."
        except Exception as exc:
            data_store["data"] = None
            generate_button.disabled = True
            status.value = f"Could not read the Accuracy sheet: {exc}"
        page.update()

    async def show_progress(message):
        status.value = message
        progress.visible = True
        progress.value = None
        page.update()
        # Yield briefly so the client can paint the indicator before the next task.
        await asyncio.sleep(0.05)

    async def generate(_):
        df = data_store["data"]
        if data_store["generating"]:
            return
        if df is None:
            status.value = "Choose a workbook before generating the comparison."
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

        data_store["generating"] = True
        pick_button.disabled = True
        generate_button.disabled = True
        progress.visible = True
        progress.value = None
        page.update()
        await asyncio.sleep(0.05)

        try:
            cutoff_value = float(cutoff.value)
            error1_value = float(error1.value)
            error2_value = float(error2.value)
            x = df.iloc[:, 3].copy()
            y = df.iloc[:, 4].copy()
            x.name = f"{analyte.value} Ref."
            y.name = f"{analyte.value} Test"

            await show_progress("Assessing measurement distributions…")
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
            status.value = f"Check the numeric inputs and workbook data: {exc}"
        except Exception as exc:
            status.value = f"Could not generate the comparison: {exc}"
        finally:
            progress.visible = False
            pick_button.disabled = False
            generate_button.disabled = data_store["data"] is None
            data_store["generating"] = False
            page.update()

    pick_button.on_click = pick_excel
    generate_button.on_click = generate
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
                                cutoff,
                                error1,
                                error2,
                                pick_button,
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
