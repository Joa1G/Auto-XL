from openpyxl import load_workbook
from package.classes import Reserve
import re

def isreservable(day_pull, day_dev, month_pull, month_dev):

    sheet_reservation = load_workbook('C:/Meus Repositórios/planilhas do auto-xl/reservas.xlsx', data_only=True)

    # vai puxar a planilha de acordo com o mês que consta na reserva.
    match month_pull:
        case '03':

            sheet_reservation_active = sheet_reservation['MAR_24']

        case '04':

            sheet_reservation_active = sheet_reservation['ABR_24']

        case _:

            sheet_reservation_active = sheet_reservation.active

    for cells in sheet_reservation_active['C']:

        # fará o loop pular a execução enquanto a linha da celula for menor que 10,
        # fiz isso por que os dados de data são o que me interessam estão a partir da linha 10.
        if cells.row < 10:
            continue    
        
        line = cells.row

        cell_value = str(sheet_reservation_active[f'C{line}'].value)
        delimiters = '[- ]'

        cell_value_splited = re.split(delimiters, cell_value)

        if cell_value_splited[2] != day_dev:
            break