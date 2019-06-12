import csv
import locale
import os
from os.path import dirname
import openpyxl
from PyQt5 import QtWidgets


def pass_to_excel(codes, uptime_code, uptime_descr, tech, manufacturer, init_date, end_date, duration, gpl,
                  cost_backout, venta_backout, total_unit_cost, total_unit_price, backout, qty, serial_no, moneda,
                  curr_file, insert_in_csv, instance):
    project_dir = dirname(os.path.abspath(os.path.dirname(__file__)))

    # Ahora abrimos el fichero Excel auxiliar de plantilla de MS
    filen = os.path.join(project_dir, 'plantilla_ms.xlsx')  # Este fichero no se toca. Hace de plantilla
    sheet_name = 'ms'

    if os.path.exists(filen):
        libro = openpyxl.load_workbook(filen)
        hoja = libro.get_sheet_by_name(sheet_name)

    else:
        instance.signals.error_fichero.emit('Fichero de plantilla{} no existe'.format(filen))
        return None, False

    curr_row = 2
    num_fila = 10
    cont_items = 0
    for i in range(len(codes)):
        fila = str(curr_row)
        fila_sig = str(curr_row + 1)

        if moneda[i] == curr_file:
            if (not insert_in_csv) or (
                    insert_in_csv[i] == 'Yes'):  # No hay filtrado a nivel global o este ítem no está filtrado
                cont_items += 1
                hoja['A' + fila] = num_fila
                hoja['A' + fila_sig] = num_fila + 1
                hoja['B' + fila_sig] = num_fila
                hoja['H' + fila] = 2
                hoja['H' + fila_sig] = 20
                hoja['C' + fila] = 'Dimension Data'
                hoja['C' + fila_sig] = manufacturer[i]
                hoja['E' + fila] = qty[i]
                hoja['E' + fila_sig] = qty[i]
                hoja['F' + fila] = uptime_code[i]
                hoja['F' + fila_sig] = backout[i]
                hoja['G' + fila] = uptime_descr[i]
                hoja['G' + fila_sig] = codes[i]
                hoja['I' + fila] = codes[i]
                hoja['I' + fila_sig] = 0
                hoja['J' + fila] = manufacturer[i]
                hoja['J' + fila_sig] = ''
                hoja['K' + fila] = 1
                hoja['K' + fila_sig] = 1
                hoja['L' + fila] = 'EA'
                hoja['L' + fila_sig] = 'EA'
                hoja['M' + fila] = moneda[i]
                hoja['M' + fila_sig] = moneda[i]
                hoja['N' + fila] = 'Fixed'
                hoja['N' + fila_sig] = 'Fixed'
                print(type(gpl[i]))
                if not gpl[i]:
                    wpl = '0'
                else:
                    if type(gpl[i]) == str:
                        wpl = gpl[i]

                    else:
                        wpl = str(locale.format_string('%.2f', gpl[i]))

                hoja['O' + fila] = wpl
                hoja['O' + fila_sig] = wpl
                try:
                    unit_price = float(total_unit_price[i]) / int(qty[i])
                except ValueError:
                    instance.signals.error_fichero.emit('Fichero no procesable\n Los datos finales de coste'
                                                        ' y PVP no están definidos')
                    return None, False

                unit_cost = float(total_unit_cost[i]) / int(qty[i])
                unit_cost_back = float(cost_backout[i]) / int(qty[i])
                unit_sell_back = float(venta_backout[i] / int(qty[i]))
                hoja['P' + fila_sig] = locale.format_string('%.2f', unit_sell_back)
                hoja['P' + fila] = locale.format_string('%.2f', unit_price)
                hoja['Q' + fila] = locale.format_string('%.2f', unit_cost)
                hoja['Q' + fila_sig] = locale.format_string('%.2f', unit_cost_back)
                try:
                    fecha_init = '{}/{}/{}'.format(init_date[i].day, init_date[i].month, init_date[i].year)
                    fecha_init_limpia = '{}{}{}'.format(str(init_date[i].year), str(init_date[i].month).zfill(2),
                                                        str(init_date[i].day).zfill(2))
                except AttributeError:
                    instance.signals.error_fichero.emit('Fichero no procesable\n La fecha de inicio'
                                                        ' no tiene formato correcto')
                    return None, False

                try:
                    fecha_fin = '{}/{}/{}'.format(end_date[i].day, end_date[i].month, end_date[i].year)
                except AttributeError:
                    instance.signals.error_fichero.emit('Fichero no procesable\n La fecha de final'
                                                        ' no tiene formato correcto')
                    return None, False

                hoja['X' + fila] = fecha_init
                hoja['X' + fila_sig] = fecha_init
                hoja['Y' + fila] = fecha_fin
                hoja['Y' + fila_sig] = fecha_fin
                hoja['AA' + fila] = 'StartDate=' + fecha_init_limpia + '#Duration=' + str(duration[i]) + \
                                    '#InvoiceInterval=Yearly#InvoiceMode=anticipated'
                hoja['AA' + fila_sig] = 'StartDate=' + fecha_init_limpia + '#Duration=' + str(duration[i]) + \
                                        '#InvoiceInterval=Yearly#InvoiceMode=anticipated'

                hoja['AF' + fila] = tech[i]
                hoja['AF' + fila_sig] = tech[i]
                hoja['AG' + fila] = serial_no[i]
                hoja['AG' + fila_sig] = serial_no[i]
                hoja['AI' + fila] = 226  # Nuevo campo 24/5/2019

                curr_row += 2
                num_fila += 3

    if cont_items == 0:
        return None, True
    else:
        return libro, True


def csv_from_excel(entrada, salida, instance):
    # VARIANTE CON xlrd. Pone los números de línea en float
    # print(entrada)
    # with xlrd.open_workbook(entrada) as wb:
    #     sh = wb.sheet_by_index(0)  # or wb.sheet_by_name('name_of_the_sheet_here')
    #     with open(salida, 'w', newline='') as f:
    #         c = csv.writer(f, delimiter=';', quoting=csv.QUOTE_MINIMAL)
    #         for r in range(sh.nrows):
    #             c.writerow(sh.row_values(r))
    ok = False
    while not ok:
        try:
            wb = openpyxl.load_workbook(entrada)
            sh = wb.get_active_sheet()  # or wb.sheet_by_name('name_of_the_sheet_here')

            with open(salida, 'w', newline='') as f:
                c = csv.writer(f, dialect='excel', delimiter=';')
                for r in sh.rows:
                    c.writerow([cell.value for cell in r])
            ok = True

        except PermissionError:
            text = 'Fichero {} ya abierto.\n Por favor, ciérrelo para seguir'.format(salida)
            QtWidgets.QMessageBox.warning(instance, "Error", text)
