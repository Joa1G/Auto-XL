# to import workbooks from certain directories, use / in place of \
# example C:\planilhas\Pasta1.xlsx turns C:/planilhas/Pasta1.xlsx
from openpyxl import *
from package.classes import Reserve
from modules.functions import isreservable
import re

sheet_of_solicitation = load_workbook('C:/Meus Repositórios/planilhas do auto-xl/solicitação.xlsx') # carrega o arquivo excel de solicitação em uma variável
sheet_of_reserve = load_workbook('C:/Meus Repositórios/planilhas do auto-xl/reservas.xlsx') # carrega o arquivo excel de reserva em outra

sheet_tab_solicitation = sheet_of_solicitation.active # ativa a primeira planilha do arquivo de solicitação em outra variável
sheet_tab_reserve = sheet_of_reserve.active

# vai pegar na planilha a data de RETIRADA que está na coluna K, irá pegar o valor contido na célula
# e irá separar cada elemento yyyy-mm-dd -> ['yyyy', 'mm', 'dd', ...]
text = str(sheet_tab_solicitation['K2'].value)
delimeters = '[- ]'
splited_date_pullout = re.split(delimeters, text)

# recebe os valores da lista com base no indice equivalente: index 2 sendo de dia, e index 1 sendo de mês.
solicitation_day_pullout = splited_date_pullout[2]
solicitation_month_pullout = splited_date_pullout[1]

# vai pegar na planilha a data de DEVOLUÇÃO que está na coluna L, irá pegar o valor contido na célula
# e irá separar cada elemento yyyy-mm-dd -> ['yyyy', 'mm', 'dd', ...]
text = str(sheet_tab_solicitation['L2'].value)
splited_date_devolution = re.split(delimeters, text)

# recebe os valores da lista com base no indice equivalente: index 2 sendo de dia, e index 1 sendo de mês.
solicitation_day_devolution = splited_date_devolution[2]
solicitation_month_devolution = splited_date_devolution[1]

isreservable(day_pull=solicitation_day_pullout, day_dev=solicitation_day_devolution, month_pull=solicitation_month_pullout, month_dev=solicitation_month_devolution)

#sheet_of_reserve.save('C:/Meus Repositórios/planilhas do auto-xl/reservas.xlsx')