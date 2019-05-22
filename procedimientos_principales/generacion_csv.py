import locale
import os
from os.path import dirname

import openpyxl

from procedimientos_auxiliares import busca_columnas, get_ultima_fila
from procedimientos_auxiliares import pass_to_excel

locale.setlocale(locale.LC_ALL, 'FR')
base_dir = dirname(os.path.abspath(os.path.dirname(__file__)))
db_file = os.path.join(base_dir + '\cisco.db')
CELDA_DESCUENTO = 'K3'
FIRST_ROW = str(11)
HEADERS_ROW = str(10)

def generacion_csv(oferta, light, instance):
    file_ramon = oferta
    carpeta, file_name = os.path.split(file_ramon)

    try:
        libro_oferta = openpyxl.load_workbook(file_ramon, data_only=True)
        sheet = libro_oferta.get_active_sheet()

    except:
        text = 'Fichero de oferta {} \nno procesable \n Elija otro.'.format(file_ramon)
        instance.signals.error_fichero.emit(text)
        return

    # Buscanos ahora en qué columnas está la información relevante. Lo haremos en varias tandas

    if light:  # Usamos plantilla Light
        lista_busqueda = ['Part Number real (*)', 'Fecha fin', 'Entitlement Uptime',
                          'Nombre de Backout', 'Precio de lista Backout Unitario  - ANUAL']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    else:  # Plantilla larga
        lista_busqueda = ['Part Number to quote', 'End Date', 'SKU - Entitlement Uptime/Smartnet Services',
                          'SKU Backout', 'Unit Backout Price List (EUR) - ANNUAL']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    if ok:
        code_col, end_date_col, upt_col, back_col, price_col = lista_columnas
    else:
        text = 'El formato de la oferta no es correcto\n El nombre de algún campo clave es incorrecto o no existe'
        instance.signals.error_fichero.emit(text)
        return

    if light:  # Plantilla Light
        lista_busqueda = ['Description Entitlement Uptime', 'Tech', 'Manufacturer',
                          'Fecha inicio', 'Duración (meses)']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    else:  # Plantilla larga
        lista_busqueda = ['Description Entitlement Uptime/Smartnet Services', 'LoB', 'Vendor',
                          'Start Date', 'Period (days)']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)
    if ok:
        uptime_descr_col, tech_col, manuf_col, init_date_col, durac_col = lista_columnas
    else:
        text = 'El formato de la oferta no es correcto\n El nombre de algún campo clave es incorrecto o no existe'
        instance.signals.error_fichero.emit(text)
        return

    if light:  # Plantilla Light
        lista_busqueda = ['Coste total de backout', 'Coste total mantenimiento', 'Venta total mantenimiento',
                          'Unid', 'Serial Number', 'Moneda', 'Venta total backout']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    else:  # Plantilla larga
        lista_busqueda = ['TOTAL Backout Cost (EUR) - PRO_RATED', 'Total Cost (EUR)', 'Total Sell Price (EUR)',
                          'Qty', 'Serial Number', 'Currency', 'TOTAL Backout Sell Price (EUR) - PRO_RATED']
        ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    if ok:
        cost_back_col, coste_tot_col, venta_col, qty_col, serial_col, currency_col, venta_back_col = lista_columnas
    else:
        text = 'El formato de la oferta no es correcto\n El nombre de algún campo clave es incorrecto o no existe'
        instance.signals.error_fichero.emit(text)
        return

    lista_busqueda = ['csv line (Yes/No)?']  # Si no está esta columna se trata de una plantilla <= v1.6 o Light
    #  --> All to csv
    ok, lista_columnas = busca_columnas(sheet, lista_busqueda, HEADERS_ROW)

    if ok:
        insert_in_csv_col = lista_columnas[0]  # Apunta la letra de la columna en la que está este campo
    else:
        insert_in_csv_col = ''

    last_row = str(get_ultima_fila(sheet, code_col))  # Buscamos cuál es la última fila rellena con algún código

    codes = [r[0].value
             for r in sheet[code_col + FIRST_ROW: code_col + last_row]]
    uptime_code = [r[0].value
                   for r in sheet[upt_col + FIRST_ROW:upt_col + str(last_row)]]
    uptime_descr = [r[0].value
                    for r in sheet[uptime_descr_col + FIRST_ROW:uptime_descr_col + str(last_row)]]
    tech = [r[0].value
            for r in sheet[tech_col + FIRST_ROW:tech_col + str(last_row)]]
    manufacturer = [r[0].value
                    for r in sheet[manuf_col + FIRST_ROW:manuf_col + str(last_row)]]
    init_date = [r[0].value
                 for r in sheet[init_date_col + FIRST_ROW:init_date_col + str(last_row)]]
    end_date = [r[0].value
                for r in sheet[end_date_col + FIRST_ROW:end_date_col + str(last_row)]]
    try:
        if light:  # En la plantilla Light la duración ya viene en meses
            duration = [int(r[0].value)
                        for r in sheet[durac_col + FIRST_ROW:durac_col + str(last_row)]]
        else:
            duration = [int(12 * int(r[0].value) / 365)
                        for r in sheet[durac_col + FIRST_ROW:durac_col + str(last_row)]]

    except TypeError:  # El fichero de oferta intermedia tiene valores no consolidados
        instance.signals.error_fichero.emit('Fichero con valores no definidos')
        return

    gpl = [r[0].value
           for r in sheet[price_col + FIRST_ROW:price_col + str(last_row)]]
    cost_backout = [r[0].value
                    for r in sheet[cost_back_col + FIRST_ROW:cost_back_col + str(last_row)]]
    venta_backout = [r[0].value
                     for r in sheet[venta_back_col + FIRST_ROW:venta_back_col + str(last_row)]]
    total_unit_cost = [r[0].value
                       for r in sheet[coste_tot_col + FIRST_ROW:coste_tot_col + str(last_row)]]
    total_unit_price = [r[0].value
                        for r in sheet[venta_col + FIRST_ROW:venta_col + str(last_row)]]
    backout_name = [r[0].value
                    for r in sheet[back_col + FIRST_ROW:back_col + str(last_row)]]
    qty = [r[0].value
           for r in sheet[qty_col + FIRST_ROW:qty_col + str(last_row)]]
    serial_no = [r[0].value
                 for r in sheet[serial_col + FIRST_ROW:serial_col + str(last_row)]]
    currency = [r[0].value
                for r in sheet[currency_col + FIRST_ROW:currency_col + str(last_row)]]
    if insert_in_csv_col:
        insert_in_csv = [r[0].value
                         for r in sheet[insert_in_csv_col + FIRST_ROW:insert_in_csv_col + str(last_row)]]
    else:
        insert_in_csv = []

    book1, ok1 = pass_to_excel(codes, uptime_code, uptime_descr, tech, manufacturer, init_date, end_date, duration,
                               gpl, cost_backout, venta_backout, total_unit_cost, total_unit_price, backout_name, qty,
                               serial_no, currency, 'USD', insert_in_csv, instance)
    book2, ok2 = pass_to_excel(codes, uptime_code, uptime_descr, tech, manufacturer, init_date, end_date, duration,
                               gpl, cost_backout, venta_backout, total_unit_cost, total_unit_price, backout_name,
                               qty, serial_no, currency, 'EUR', insert_in_csv, instance)

    print('Hasta aquí hemos llegado.\n Ni tan mal')
    ok_total = False

    if ok1 and ok2:
        ok_total = True

    instance.signals.fin_OK_csv.emit(book1, book2, carpeta, ok_total)

