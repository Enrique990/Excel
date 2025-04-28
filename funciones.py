
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import datetime
import openpyxl.utils

def cargar_notas():
    try:
        wb = openpyxl.load_workbook('notas_estudiantes.xlsx')
    except FileExistsError:
        print("No se encontro el archivo de notas")
        return None
    
    ws = wb.active
    ws.title = "notas"

    return  wb, ws

def automatizacion_notas():
    wb, ws = cargar_notas()
    if not wb:
        return
    
    wb = generar_reporte()

    wb.save('notas_estudiantes.xlsx')
    print("proceso completado")

def generar_reporte():
    wb, ws = cargar_notas()
    if not wb:
        return
    
    #crear una hoja nueva para el reporte
    fecha_actual = datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S")
    ws_reporte = wb.create_sheet(f"Reporte {fecha_actual}")

    #agregar encabezados
    ws_reporte.cell(row=1, column=1, value='Estadisticas').font = Font(bold=True)
    ws_reporte.cell(row=1, column=2, value='Numeros').font = Font(bold=True)
    ws_reporte.cell(row=1, column=3, value='Porcentajes').font = Font(bold=True)
    ws_reporte.cell(row=2, column=1, value='Numero total de estudiantes')
    ws_reporte.cell(row=3, column=1, value='Numero de estudiantes aprobados (nota >= 70)')
    ws_reporte.cell(row=4, column=1, value='Porcentaje de estudiantes aprobados (nota >= 70)')
    ws_reporte.cell(row=5, column=1, value='Numero de estudiantes reprobados (nota < 70)')
    ws_reporte.cell(row=6, column=1, value='Porcentaje de estudiantes reprobados (nota < 70)')
    ws_reporte.cell(row=7, column=1, value='Numero de estudiantes reprobados con notas entre 60 y 69')
    ws_reporte.cell(row=8, column=1, value='Porcentaje de estudiantes reprobados con notas entre 60 y 69')
    ws_reporte.cell(row=9, column=1, value='Media de las notas')
    ws_reporte.cell(row=10, column=1, value='Desviación estandar de las notas')

    #calcular total de estudiantes
    ws_reporte.cell(row=2, column=2, value=f"=COUNTA(notas!A2:A{ws.max_row})")

    #calcular numero de estudiantes aprobados (nota >= 70)
    ws_reporte.cell(row=3, column=2, value=f'=COUNTIF(notas!B2:B{ws.max_row},">=70")')

    #calcular porcentaje de estudiantes aprobados (nota >= 70)
    ws_reporte.cell(row=4, column=3, value=f"=B3/B2")
    ws_reporte["C4"].number_format = '0.00%'

    #calcular numero de estudiantes reprobados (nota < 70)
    ws_reporte.cell(row=5, column=2, value=f'=COUNTIF(notas!B2:B{ws.max_row},"<70")')

    #calcular porcentaje de estudiantes reprobados (nota < 70)
    ws_reporte.cell(row=6, column=2, value=f"=B5")  # opcional, repetir valor absoluto
    ws_reporte.cell(row=6, column=3, value=f"=B5/B2")
    ws_reporte["C6"].number_format = '0.00%'

    #calcular numero de estudiantes reprobados con notas entre 60 y 69
    ws_reporte.cell(row=7, column=2, value=f'=COUNTIFS(notas!B2:B{ws.max_row},">=60",notas!B2:B{ws.max_row},"<70")')

    #calcular porcentaje de estudiantes reprobados con notas entre 60 y 69
    ws_reporte.cell(row=8, column=2, value=f"=B7")  # opcional, solo para mantener consistencia
    ws_reporte.cell(row=8, column=3, value=f"=B7/B2")
    ws_reporte["C8"].number_format = '0.00%'

    #calcular media de las notas
    ws_reporte.cell(row=9, column=2, value=f"=AVERAGE(notas!B2:B{ws.max_row})")
    ws_reporte["B9"].number_format = '0.00'

    #calcular desviación estandar de las notas
    ws_reporte.cell(row=10, column=2, value=f"=STDEVP(notas!B2:B{ws.max_row})")
    ws_reporte.cell(row=10, column=2).data_type = 'f'  # Tipo fórmula
    ws_reporte["B10"].number_format = '0.00'


    ws_reporte.column_dimensions[openpyxl.utils.get_column_letter(1)].width = 50
    ws_reporte.column_dimensions[openpyxl.utils.get_column_letter(3)].width = 11

    return wb