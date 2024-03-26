from openpyxl import load_workbook
from package.classes import Reserve

def isreservable(day_pull, day_dev, month_pull, month_dev):
    sheet_reservation = load_workbook('C:/Meus Repositórios/planilhas do auto-xl/reservas.xlsx')
    
