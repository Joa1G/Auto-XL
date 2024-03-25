# to import workbooks from certain directories, use / in place of \
# example C:\planilhas\Pasta1.xlsx turns C:/planilhas/Pasta1.xlsx
from openpyxl import *

sheet_of_solicitation = load_workbook('solicitação.xlsx') # carrega o arquivo excel de solicitação em uma variável
sheet_of_reserve = load_workbook('reservas.xlsx') # carrega o arquivo excel de reserva em outra

sheet_tab_solicitation = sheet_of_solicitation.active # ativa a primeira planilha do arquivo de solicitação em outra variável

# irá verificar célula por célula da coluna de data de retirada na planilha de solicitação, para verificar a principio
# em qual planilha (onde cada uma representa um mês) do arquivo de reservas ele irá ativar.
for cell in sheet_tab_solicitation['K']:

    if cell.row == 1: # se a celula estiver na linha 1 vai pular o loop, pois não quero acessar o valor que está na
        continue      # que está na primeira linha, pois é o titulo da tabela.

    cell_splited = str(cell.value).split('-') # ao receber a data recebe nesse formato: yyyy-mm-dd, com isso removo os traços(-)
    
    match cell_splited[1]: # acessa o indice 1 da lista gerada, que vai armazenar o valor do mês
        case '04':
            sheet_tab_reserve  = sheet_of_reserve['ABR_24']
            matricula = sheet_tab_solicitation['F2'].value
            matricula = int(matricula)
            sheet_tab_reserve['G23'] = matricula
            break
        case '03':
            sheet_tab_reserve = sheet_of_reserve['MAR_24']
        case _:
            sheet_of_reserve = sheet_of_reserve.active

sheet_of_reserve.save('reservas.xlsx')
