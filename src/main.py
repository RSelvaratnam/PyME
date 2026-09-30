from io import BytesIO

import flet as ft

try:
    import pandas as pd
except ImportError:
    pd = None

try:
    import PyME_functions as CC
except ImportError:
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

    analyte = ft.TextField(label="Measurand", value="Glucose")
    unit = ft.TextField(label="Unit", value="mg/dL")
    cutoff = ft.TextField(label="Cut-off", value="2")
    error1 = ft.TextField(label="Absolute error", value="0.2")
    error2 = ft.TextField(label="% error", value="10")

    hist_img = ft.Image(src=b"", visible=False, width=900)
    deming_img = ft.Image(src=b"", visible=False, width=600)
    status = ft.Text("Choose an Excel workbook to begin.")
    generate_button = ft.Button(
        content="Generate", disabled=True, on_click=None
    )
    data_store = {"data": None}
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

    async def generate(_):
        df = data_store["data"]
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

        try:
            cutoff_value = float(cutoff.value)
            error1_value = float(error1.value)
            error2_value = float(error2.value)
            x = df.iloc[:, 3].copy()
            y = df.iloc[:, 4].copy()
            x.name = f"{analyte.value} Ref."
            y.name = f"{analyte.value} Test"

            hist_stream = CC.Histogram_grouped(x, y)
            hist_img.src = stream_to_bytes(hist_stream)
            hist_img.visible = True

            deming_stream, n, slope, intercept, r = (
                CC.Deming_Plot_Equal_Variance_with_Error2_PX(
                    x,
                    y,
                    Error_level_cut_off=cutoff_value,
                    error1=error1_value,
                    error2=error2_value,
                )
            )
            deming_img.src = stream_to_bytes(deming_stream)
            deming_img.visible = True

            CC.MC_output(
                analyte=analyte.value,
                document=Document(),
                x=x,
                y=y,
                Unit=unit.value,
                Error_level_cut_off=cutoff_value,
                error1=error1_value,
                error2=error2_value,
            )
            status.value = (
                f"Generated comparison for {n} samples. "
                f"Slope: {slope:.3f}; intercept: {intercept:.3f}; r: {r:.3f}."
            )
        except (ValueError, TypeError) as exc:
            status.value = f"Check the numeric inputs and workbook data: {exc}"
        except Exception as exc:
            status.value = f"Could not generate the comparison: {exc}"
        page.update()

    generate_button.on_click = generate
    page.add(
        ft.Column(
            controls=[
                ft.Row(
                    controls=[
                        ft.Column(
                            controls=[
                                analyte,
                                unit,
                                cutoff,
                                error1,
                                error2,
                                ft.Button(content="Pick Excel", on_click=pick_excel),
                                generate_button,
                                status,
                            ]
                        ),
                        ft.Column(controls=[hist_img, deming_img]),
                    ],
                    vertical_alignment=ft.CrossAxisAlignment.START,
                )
            ],
            scroll=ft.ScrollMode.AUTO,
        )
    )


ft.run(main)
