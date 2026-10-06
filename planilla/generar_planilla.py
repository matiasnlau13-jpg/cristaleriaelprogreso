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
    "COPAS": ["Nombre", "Nombre entre comillas", "Pierna (estirada / pegada)", "Corte caliente c/c (altura)", "c/",
              "Peso (P)", "Base (B)", "Pie", "Pinza", "Molde", "Destino", "Nota en lápiz"],
    "FLORERO": ["Nombre", "Nombre entre comillas", "Capacidad", "Corte caliente c/c (altura)", "Otro corte", "Peso (P)", "Base (B)", "Nota en lápiz"],
    "TULIPA": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "PAMELA": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "BOTELLON": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "TARRO": ["Nombre", "Nombre entre comillas", "Capacidad", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)", "Nota en lápiz"],
    "CENTRO": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "TAPA": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)", "Molde"],
    "JARRA": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "ACEITERA": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "OTROS": ["Nombre", "Nombre entre comillas", "Corte caliente c/c (altura)", "Peso (P)", "Base (B)"],
    "PRENSA": ["Tipo", "Nombre", "Nombre entre comillas", "Peso (P)", "Base (B)"],
}
EXTRA = ["Sección del cuaderno", "Foto Nº", "Observaciones"]

# Filas por hoja, en el mismo orden que COLUMNAS[hoja] + EXTRA.
DATOS = {
    "COPAS": [
        ["Copa Romina Vino", "", "Pegada", "15,3", "", 200, 230, "Lorena V.B", "", "", "", "", "Copas", 2, ""],
        ["Copa Presidente Champ", "", "", "", "", 180, 200, "", "", "", "", "Renato", "Copas", 2, ""],
        ["Copa Clarito Inés De los Santos", "", "", "", "", 200, 230, "Presi V.T", "", "", "", "", "Copas", 2, ""],
        ["Copa Prince Agua", "", "Pegada", "19,7", "", 220, 220, "London 98", "", "", "", "", "Copas", 2,
         "El peso está remarcado/sobrescrito: confirmar que es 220"],
        ["Copa Barón", "B", "Estirada", "", "", 250, 200, "London 98", "", "", "", "", "Copas", 2, ""],
        ["Copa Chandón 2021", "", "Estirada", "", "", 230, 200, "London 98", "", "", "", "", "Copas", 2,
         "El pie no está escrito: hay comillas (\") debajo de \"Pie London 98\" de la fila de arriba"],
        ["Copa Magnífica", "", "Estirada", "", "", 160, 200, "Delise", "", "", "", "", "Copas", 2, ""],
        ["Copa Venus Champ", "", "Estirada", "", "", 160, 220, "", "", "", "", "", "Copas", 2, ""],
        ["Copa Flavia Champ", "Renato", "Estirada", "", "", 160, 220, "", "", "", "", "", "Copas", 2, ""],
        ["Copa Ceci Flauta", "", "Estirada", "", "24", 170, 200, "Presi Agua", "", "", "", "", "Copas", 2, ""],
        ["Copa Him-Chan Agua", "", "", "", "15", 150, 230, "Venus V.B", "", "", "", "Renato", "Copas", 2, ""],
        ["Copa Retro", "", "Pegada", "", "", 180, 220, "Mod II Agua", "", "", "", "Renato", "Copas", 2, ""],
        ["Copa Premium Mediana", "", "", "", "", 230, 220, "Mini Marisco", "", "", "", "", "Copas", 2, ""],
        ["Copa Premium Chica", "", "", "", "", 220, 220, "Mod II Agua", "", "", "", "", "Copas", 2, ""],
        ["Copa Silvia Agua", "", "Pegada", "", "", 180, 220, "Silvia Agua", "", "", "", "Renato", "Copas", 2, ""],
        ["Copa Expresión", "", "Pegada", "", "", 230, 200, "Premium", "", "", "", "", "Copas", 3, ""],
        ["Copa Martini Champ", "C. Ruiz", "Alta", "", "", 260, 200, "Eva Champ", "", "", "", "", "Copas", 3, ""],
        ["Copa Nueva Progreso", "", "", "21,0", "", 300, 200, "Premium", "19", "", "", "", "Copas", 3, ""],
        ["Copa Hurricane", "", "", "20,5", "", 300, 200, "Coñac London 98", "90", "", "", "", "Copas", 3,
         "Pinza: se lee 90 (las demás dicen 19). Confirmar"],
        ["Copa Patagonia Cerveza", "", "", "15,3", "", 250, 230, "London 98", "19", "", "", "", "Copas", 3, ""],
        ["Copa Taby Cerveza", "", "", "14,3", "", 280, 200, "Venus Agua", "19", "", "", "", "Copas", 3, ""],
        ["Copa Ani", "G / German", "Pegada (pierna fina)", "", "", 280, 200, "Marisco Deisy", "",
         "Copa \"G\"", "German", "", "Copas", 3,
         "Molde copa \"G\": puede ser G o 6. Confirmar"],
        ["Copa Checo Flauta", "Renato", "Estirada", "", "22,0", 170, 230, "Presi Agua", "",
         "Erika de Rosemblet", "", "", "Copas", 3, ""],
    ],
    "TARRO": [
        ["Tarro Cónico", "", "1/2 L", "13,0", 500, 250, "", "Sin título", 4, ""],
        ["Tarro Claret", "", "1 1/2 L", "20,1", 1100, 120, "B:55", "Sin título", 4,
         "Tiene un segundo \"B:55\" al final del renglón. ¿Qué es?"],
        ["Tarro Bocha", "", "3/4 L", "11,0", 300, 250, "", "Sin título", 4, ""],
        ["Tarro M-13x8,5", "", "", "8,5", 800, 200, "", "Sin título", 4, ""],
        ["Tarro M-13x15", "", "", "15,0", 1000, 120, "", "Sin título", 4, ""],
        ["Tarro M-13x20", "", "", "20,0", 1000, 120, "", "Sin título", 4, ""],
        ["Tarro M-13x25", "", "", "25,0", 1100, 120, "", "Sin título", 4, ""],
        ["Tarro M-13x30", "", "", "30,0", 1200, 120, "", "Sin título", 4, ""],
    ],
    "FLORERO": [
        ["Florero Cónico", "", "1 L", "14,5", "", 900, 150, "", "Sin título", 4, ""],
        ["Florero Cónico", "", "2 L", "18,0", "", 1000, 100, "", "Sin título", 4, ""],
        ["Florero 12x30", "", "", "30,0", "", 1000, 100, "", "Sin título", 4, ""],
        ["Florero 17x50", "", "", "47,0", "", 2500, 100, "", "Sin título", 4, ""],
        ["Florero 17x50 2 cortes", "", "", "", "2 cortes", 3000, 100, "", "Sin título", 4, ""],
        ["Florero Lila", "", "", "15,3", "", 650, 200, "Chico", "Sin título", 4, ""],
        ["Florero Lila Diagonal", "", "", "Diagonal", "", 850, 150, "", "Sin título", 4,
         "Dice \"c/c Diagonal\" sin altura: ¿corte en diagonal?"],
        ["Florero Potiche", "", "", "23,5", "", 1200, 120, "", "Sin título", 4, ""],
        ["Florero Cuadrado Grande", "", "", "", "", 1200, 200, "", "Sin título", 5, ""],
        ["Florero Mod 11", "Cenatiempo", "", "26,5", "c/ 24 cm", 700, 200, "", "Sin título", 5,
         "Abajo dice: corte c/24cm \"Cenatiempo\". Confirmar cómo se escribe Cenatiempo"],
        ["Florero Cónico Alto", "", "", "28,0", "", 550, 250, "", "Sin título", 5, ""],
        ["Florero Cónico Bajo", "", "", "19,3", "", 400, 250, "", "Sin título", 5, ""],
    ],
    "TULIPA": [
        ["Tulipa Mini Mum", "", "", 900, 150, "Sin título", 5, ""],
        ["Tulipa Rochester", "", "20,5", 1500, 150, "Sin título", 5, ""],
    ],
    "PAMELA": [
        ["Pamela x21", "", "9,0", 1300, 150, "Sin título", 5, ""],
        ["Pamela x16", "", "7,6", 700, 150, "Sin título", 5, ""],
        ["Pamela x11", "", "", 400, 300, "Sin título", 5, ""],
        ["Pamela x10", "", "5,3", 350, 250, "Sin título", 5, ""],
    ],
    "BOTELLON": [
        ["Botellón Neo", "", "", 300, 300, "Sin título", 5, ""],
        ["Botellón Inti Chico", "", "", 400, 200, "Sin título", 5, ""],
        ["Botellón Mini", "", "", 500, 250, "Sin título", 5, ""],
        ["Botellón Decanter", "", "", 1300, 150, "Sin título", 6, ""],
        ["Botellón Decantador Progreso", "", "", 1300, 150, "Sin título", 6,
         "Escrito en lápiz debajo del Botellón Decanter. ¿Es la misma pieza con otro nombre?"],
        ["Botellón Nahuel", "", "", 1300, 100, "Sin título", 6, "Escrito en lápiz"],
    ],
    "CENTRO": [
        ["Centro Oreford Grande", "", "", 2800, 100, "Sin título", 6, ""],
        ["Centro Porta Vela", "", "9,5", 1300, 100, "Sin título", 6, ""],
    ],
    "TAPA": [
        ["Tapa Quesera 1/2 Esfera", "", "c/c (sin altura)", 1300, 120, "", "Sin título", 6,
         "Dice c/c pero no tiene la altura"],
        ["Tapa Quesera con Botón", "", "", 3000, 120, "", "Sin título", 6, ""],
        ["Tapa 17x20", "", "20,0", 1200, 120, "", "Sin título", 6, ""],
        ["Tapa 17x13", "", "13,0", 1200, 120, "", "Sin título", 6, ""],
        ["Tapa M-13", "", "", 400, 250, "Also", "Sin título", 6, "Molde: se lee \"Also\". Confirmar"],
    ],
    "JARRA": [
        ["Jarra Carafe con Pico y Asa", "", "16,0", 550, 55, "Sin título", 6, ""],
    ],
    "ACEITERA": [
        ["Aceitera Cónico Chico con Pico", "", "", 400, 250, "Sin título", 6, ""],
    ],
    "OTROS": [
        ["Bombe Gigante", "", "15,8", 700, 200, "Sin título", 6, "No entra en ninguna categoría. ¿Dónde va?"],
    ],
    "PRENSA": [
        ["Plato", "Plato Pizza", "", 300, 500, "Prensa", 1, ""],
        ["Copa", "Copa Alaska", "", 320, 450, "Prensa", 1,
         "En lápiz: \"Requem. 8\" y un \"300\" debajo del B:450 (¿B corregido a 300?)"],
        ["Faro", "Faro Puma", "", 130, 450, "Prensa", 1, ""],
        ["Faro", "Faro Ford", "", 452, 600, "Prensa", 1, ""],
        ["Copa", "Copa Milk Shake", "", 570, 250, "Prensa", 1, ""],
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
