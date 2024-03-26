from openpyxl import Workbook
class Reserve:
    def __init__(self, matricula, veiculo, ramal, hora_saida, hora_devolucao):
        self.matricula = matricula
        self.veiculo = veiculo
        self.ramal = ramal
        self.hora_saida = hora_saida
        self.hora_devoluacao = hora_devolucao
    def is_reservable(self, matricula, veiculo, ramal, hora_saida, hora_devolucao):
        ...