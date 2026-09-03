import html
import os
import re
from datetime import datetime
from zoneinfo import ZoneInfo

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

URL = "https://www.facilito.gob.pe/facilito/pages/facilito/buscadorEESS.jsp"
OUT_PATH = "historial/Chiclayo_GasoholRegular_historial.xlsx"
LIMA = ZoneInfo("America/Lima")

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF")


def scrape():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        page.goto(URL, wait_until="load", timeout=60000)

        # La pagina arranca mostrando un mapa (sin <select> en el DOM).
        # Los enlaces del mapa llaman makeAction(codigoDepartamento) directamente.
        page.evaluate("makeAction(140000)")  # LAMBAYEQUE
        page.wait_for_selector('select[name="provincia"] option[value="140100"]', timeout=15000)

        # Elegir provincia dispara un submit de formulario (navegacion completa).
        with page.expect_navigation(wait_until="load", timeout=30000):
            page.evaluate(
                """() => {
                    const el = document.querySelector('select[name="provincia"]');
                    el.value = '140100';  // CHICLAYO
                    el.dispatchEvent(new Event('change', {bubbles: true}));
                }"""
            )

        # Elegir producto carga la tabla via AJAX en la misma pagina (sin navegar).
        page.evaluate(
            """() => {
                const el = document.querySelector('select[name="producto"]');
                el.value = '126';  // Gasohol Regular
                el.dispatchEvent(new Event('change', {bubbles: true}));
            }"""
        )

        page.wait_for_function(
            "() => window.jQuery && $('table').DataTable && "
            "$('table').DataTable().rows().data().toArray().length > 0",
            timeout=30000,
        )

        raw = page.evaluate(
            "() => $('table').DataTable().rows().data().toArray()"
            ".map(r => Array.isArray(r) ? r : Object.values(r))"
        )
        browser.close()
        return raw


def clean(raw):
    records = []
    for item in raw:
        distrito = item[0].strip()
        establecimiento = html.unescape(item[1]).strip()
        direccion = html.unescape(item[2]).strip()
        precio_html = item[4]
        m = re.search(r"([\d.]+)", precio_html)
        precio = float(m.group(1))
        records.append(
            {
                "Distrito": distrito,
                "Establecimiento": establecimiento,
                "Direccion": direccion,
                "Precio": precio,
                "Key": f"{establecimiento}|{direccion}",
            }
        )
    return records


def style_header_cell(cell):
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")


def update_workbook(records, now):
    sheet_name = now.strftime("%Y-%m")
    col_header = now.strftime("%Y-%m-%d %H:%M")

    os.makedirs(os.path.dirname(OUT_PATH), exist_ok=True)

    if os.path.exists(OUT_PATH):
        wb = load_workbook(OUT_PATH)
    else:
        wb = Workbook()
        wb.remove(wb.active)

    if sheet_name not in wb.sheetnames:
        ws = wb.create_sheet(sheet_name)
        records_sorted = sorted(records, key=lambda r: r["Precio"])

        headers = ["Distrito", "Establecimiento", "Direccion", col_header]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            style_header_cell(cell)

        for r, rec in enumerate(records_sorted, start=2):
            ws.cell(row=r, column=1, value=rec["Distrito"]).font = Font(name=FONT_NAME)
            ws.cell(row=r, column=2, value=rec["Establecimiento"]).font = Font(name=FONT_NAME)
            ws.cell(row=r, column=3, value=rec["Direccion"]).font = Font(name=FONT_NAME)
            pc = ws.cell(row=r, column=4, value=rec["Precio"])
            pc.font = Font(name=FONT_NAME)
            pc.number_format = '"S/" #,##0.00'

        widths = [18, 42, 55, 16]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:D{ws.max_row}"
    else:
        ws = wb[sheet_name]
        last_row = ws.max_row
        last_col = ws.max_column
        new_col = last_col + 1

        hcell = ws.cell(row=1, column=new_col, value=col_header)
        style_header_cell(hcell)
        ws.column_dimensions[get_column_letter(new_col)].width = 16

        key_to_row = {}
        for r in range(2, last_row + 1):
            key = f"{ws.cell(row=r, column=2).value}|{ws.cell(row=r, column=3).value}"
            key_to_row[key] = r

        append_row = last_row + 1
        for rec in records:
            if rec["Key"] in key_to_row:
                row = key_to_row[rec["Key"]]
            else:
                row = append_row
                ws.cell(row=row, column=1, value=rec["Distrito"]).font = Font(name=FONT_NAME)
                ws.cell(row=row, column=2, value=rec["Establecimiento"]).font = Font(name=FONT_NAME)
                ws.cell(row=row, column=3, value=rec["Direccion"]).font = Font(name=FONT_NAME)
                key_to_row[rec["Key"]] = row
                append_row += 1
            pc = ws.cell(row=row, column=new_col, value=rec["Precio"])
            pc.font = Font(name=FONT_NAME)
            pc.number_format = '"S/" #,##0.00'

        final_last_row = append_row - 1
        ws.auto_filter.ref = f"A1:{get_column_letter(new_col)}{final_last_row}"

    if "Notas" not in wb.sheetnames:
        wsn = wb.create_sheet("Notas")
        wsn.cell(row=1, column=1, value=(
            "Fuente: Facilito - Osinergmin (facilito.gob.pe), Diesel y Gasolina > "
            "Lambayeque > Chiclayo > Gasohol Regular. Precios reportados por los propios "
            "operadores. Cada hoja mensual (AAAA-MM) contiene una fila por grifo y una "
            "columna por corrida (encabezado = fecha y hora de la consulta, hora Peru). "
            "Actualizado automaticamente via GitHub Actions cada 3 horas."
        )).font = Font(name=FONT_NAME)
        wsn.column_dimensions["A"].width = 100
        wsn.cell(row=1, column=1).alignment = Alignment(wrap_text=True)

    wb.save(OUT_PATH)
    print(f"Hoja: {sheet_name} | Columna: {col_header} | Grifos: {len(records)}")


if __name__ == "__main__":
    now = datetime.now(LIMA)
    raw = scrape()
    records = clean(raw)
    if len(records) < 50:
        raise SystemExit(f"Muy pocos registros ({len(records)}) - probablemente fallo el scraping, no se guarda.")
    update_workbook(records, now)
