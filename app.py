import os
from openpyxl import Workbook
from openpyxl import load_workbook

class NotaFiscal:
    def __init__(self, dic_c100) -> None:
        self.dic_c100 = dic_c100
        self.lista_c185 = []
        self.nota_valida = False

    def add_c185(self, dic_c185):
        self.lista_c185.append(dic_c185)


def criar_dic_0200():
    dic_0200 = {}
    for arquivo_sped in os.listdir('SPED'):
        with open(f"SPED{os.sep}{arquivo_sped}", "r", encoding='ansi') as arquivo:
            for registro in arquivo:
                notas = registro.strip().split('|')

                if notas[1] == '9999':
                    break
                
                try:
                    if notas[1] == '0200':
                        dic_0200[notas[2]] = notas[2::]
                except IndexError:
                    pass
    return dic_0200

def ler_notas():

    for arquivo_sped in os.listdir('SPED'):
        # print(arquivo_sped)
        lista_notas = []
        with open(f"SPED{os.sep}{arquivo_sped}", "r", encoding='ansi') as arquivo:
            dic_0200 = {}
            dic_c185 = {}
            for registro in arquivo:
                notas = registro.strip().split('|')
                
                if notas[1] == '9999':
                    break

                if notas[1] == '0200':
                    dic_0200[notas[2]] = notas[2::]

                try:
                    if notas[1] == 'C100':
                        dic_c100 = {}
                        dic_c100[notas[9]] = notas
                        nota_fiscal = NotaFiscal(dic_c100)
                        lista_notas.append(nota_fiscal)
                except IndexError:
                    pass

                try:
                    if notas[1] == 'C185':
                        dic_c185 = {}
                        dic_c185[notas[2]] = notas
                        nota_fiscal.nota_valida = True
                        nota_fiscal.add_c185(dic_c185)

                except IndexError:
                    pass
            return lista_notas

        0

if __name__ == '__main__':
    wb = Workbook()
    ws = wb.active
    
    dic_0200 = criar_dic_0200()
    # print(dic_0200)
    lista_notas = ler_notas()
    for nota in lista_notas:
        if nota.nota_valida:
            for k, v_c100 in nota.dic_c100.items():
                # print(f'**** nota *****')
                # print(v)

                for c185 in nota.lista_c185:
                    for k, v_c185 in c185.items():
                        # print('*** c185 ***')
                        
                        descricao_item = dic_0200.get(v_c185[3])[1]
                        v_c185.insert(4, descricao_item)
                        # print(v_c185)
                        lista_excel = v_c100 + v_c185
                        # print(lista_excel)
                        ws.append(lista_excel)
                        # print(v)
    wb.save('c185.xlsx')
            
        

    
        