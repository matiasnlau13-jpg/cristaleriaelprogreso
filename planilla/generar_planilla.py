"""Genera planilla_piezas.xlsx a partir de los datos transcriptos de las fotos del cuaderno.

Para agregar piezas, sumar filas en DATOS y volver a correr:
    python planilla/generar_planilla.py
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

SALIDA = Path(__file__).with_name("planilla_piezas.xlsx")

# Columnas de cada hoja. "P" = peso, "B" = B (como figura en el cuaderno).
COLUMNAS = {
    "VASOS": ["Nombre", "Nombre entre comillas", "Altura", "Peso (P)", "Base (B)"],
    "COPAS": ["Nombre", "Nombre entre comillas", "Pierna (estirada / pegada)", "Peso (P)", "Base (B)", "Pie"],
    "FLORERO": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "TULIPA": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "PAMELA": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "BOTELLON": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "TARRO": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "CENTRO": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "TAPA": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "JARRA": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "ACEITERA": ["Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
    "OTROS": ["Tipo", "Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
}
EXTRA = ["Sección del cuaderno", "Foto Nº", "Observaciones"]

# Filas por hoja, en el mismo orden que COLUMNAS[hoja] + EXTRA.
DATOS = {
    "COPAS": [
        ["Copa Alaska", "", "", 320, 450, "", "Prensa", 1,
         "En lápiz: \"Requem. 8\" y un \"300\" debajo del B:450 (¿B corregido a 300?)"],
        ["Copa Milk Shake", "", "", 570, 250, "", "Prensa", 1, ""],
    ],
    "OTROS": [
        ["Plato", "Plato Pizza", "", 300, 500, "Prensa", 1, ""],
        ["Faro", "Faro Puma", "", 130, 450, "Prensa", 1, ""],
        ["Faro", "Faro Ford", "", 452, 600, "Prensa", 1, ""],
        ["Filtro", "Filtro Nº 2", "", 245, 500, "Prensa", 1, ""],
        ["Filtro", "Filtro Nº 4", "SINAC", 40, 600, "Prensa", 1,
         "Escrito en lápiz; \"SINAC\" con flecha desde B:600"],
        ["Base Velador", "Base Velador Grande", "", 465, 400, "Prensa", 1, ""],
        ["Base Velador", "Base Velador Chico", "Rococo", 280, 600, "Prensa", 1,
         "\"Rococo\" está escrito debajo del Chico"],
        ["Lupa", "Lupa Nº 3", "FYW", 88, 500, "Prensa", 1, ""],
        ["Lupa", "Lupa Nº 4", "", 350, 500, "Prensa", 1, ""],
        ["Vidrio", "Vidrio", "U", 620, 300, "Prensa", 1, "Escrito en lápiz"],
        ["Baliza", "Baliza Anillada Grande", "BALIZAR", 650, 300, "Prensa", 1, ""],
    ],
}

FUENTE = "Arial"
ENCABEZADO_FILL = PatternFill("solid", start_color="1F4E78")
DUDA_FILL = PatternFill("solid", start_color="FFF2CC")
BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)


def main():
    wb = Workbook()
    wb.remove(wb.active)
    for hoja, cols in COLUMNAS.items():
        ws = wb.create_sheet(hoja)
        encabezados = cols + EXTRA
        ws.append(encabezados)
        for celda in ws[1]:
            celda.font = Font(name=FUENTE, bold=True, color="FFFFFF")
            celda.fill = ENCABEZADO_FILL
            celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            celda.border = BORDE
        for fila in DATOS.get(hoja, []):
            ws.append(fila)
            r = ws.max_row
            for celda in ws[r]:
                celda.font = Font(name=FUENTE)
                celda.border = BORDE
                celda.alignment = Alignment(vertical="center", wrap_text=True)
            if fila[-1]:
                for celda in ws[r]:
                    celda.fill = DUDA_FILL
        for i, titulo in enumerate(encabezados, start=1):
            letra = ws.cell(row=1, column=i).column_letter
            ws.column_dimensions[letra].width = 55 if titulo == "Observaciones" else max(14, len(titulo) + 4)
        ws.row_dimensions[1].height = 32
        ws.freeze_panes = "A2"
    wb.save(SALIDA)
    print(f"Guardado {SALIDA}")


if __name__ == "__main__":
    main()
