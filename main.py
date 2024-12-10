from datetime import datetime
import calendar
import random
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font

import locale
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

import copy
from Voluntários import Voluntários

from Voluntários import mes, ano

Voluntários = copy.deepcopy(Voluntários)

class diaServentia:
    def __init__ (self, dia, tipo):
        self.dia = dia
        self.tipo = tipo
        # Mídia
        self.Foto = None
        self.Story = None
        self.Mesa = None
        
        # Teens
        self.Pré = None
        self.Kids = None
        self.Babys = None
        self.Auxiliares = []

        # LOUVOR
        self.Ministro = None
        self.Backs = []
        self.Bateria = None
        self.Teclado = None
        self.Guitarra = None
        self.Baixo = None
        self.Violão = None
    
random.shuffle(Voluntários)

diasDeQuinta = []
diasDeDomingo = []
diasDeSabado = []

for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 3:  
        diasDeQuinta.append(dia)

# Listar todos os dias que são domingo no próximo mês
for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 6:  
        diasDeDomingo.append(dia)

# Listar todos os dias que são sábado no próximo mês
for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 5:  
        diasDeSabado.append(dia)

objetoDosDiasdeQuinta = []
objetoDosDiasdeDomingo = []
objetoDosDiasdeSabado = []

for dia in diasDeQuinta:
    temp = diaServentia(dia, "quinta")
    objetoDosDiasdeQuinta.append(temp)

for dia in diasDeDomingo:
    temp = diaServentia(dia, "domingo")
    objetoDosDiasdeDomingo.append(temp)

for dia in diasDeSabado:
    temp = diaServentia(dia, "sabado")
    objetoDosDiasdeSabado.append(temp)

funcao_para_atributo = {
        "Foto": "Foto",
        "Story": "Story",
        "Mesa": "Mesa",
        "Pré": "Pré",
        "Kids": "Kids",
        "Babys": "Babys",
        "Backs": "diasDeBacks",
        "Bateria": "diasDeBateria",
        "Teclado": "diasDeTeclado",
        "Guitarra": "diasDeGuitarra",
        "Baixo": "diasDeBaixo"
    }

def fazerUmaEscala(ministerio, funcao, dia):
    
    metaParaTodos = 1
    if dia == "quinta":
        for diaDeQuinta in objetoDosDiasdeQuinta:
            random.shuffle(Voluntários)
            i = 0
            while i != len(Voluntários):  
                if (
                    ministerio in Voluntários[i]["ministerios"]
                    and funcao in Voluntários[i]["funcoes"]
                    and Voluntários[i]["diasServidos"] < metaParaTodos
                    and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                    and Voluntários[i]["quinta"]
                ):    
                    if funcao in funcao_para_atributo:
                        atributo = funcao_para_atributo[funcao]
                        setattr(diaDeQuinta, atributo, Voluntários[i]["nome"])

                    Voluntários[i]["diasServidos"] += 1
                    Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                    break;
                if Voluntários[i] == Voluntários[-1]:
                    metaParaTodos += 1 
                    i = 0
                else:
                    i += 1

    if dia == "domingo":
        for diaDeDomingo in objetoDosDiasdeDomingo:
            random.shuffle(Voluntários)
            i = 0
            while i != len(Voluntários):
                if (
                    ministerio in Voluntários[i]["ministerios"]
                    and funcao in Voluntários[i]["funcoes"]
                    and Voluntários[i]["diasServidos"] < metaParaTodos
                    and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                    and Voluntários[i]["domingo"]
                ):    
                    if funcao in funcao_para_atributo:
                        atributo = funcao_para_atributo[funcao]
                        setattr(diaDeDomingo, atributo, Voluntários[i]["nome"])
                        
                    Voluntários[i]["diasServidos"] += 1
                    Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                    break;
                if Voluntários[i] == Voluntários[-1]:
                    metaParaTodos += 1 
                    i = 0
                else:
                    i += 1

diasJuntos = objetoDosDiasdeQuinta + objetoDosDiasdeDomingo
diasJuntosOrdenados = sorted(diasJuntos, key=lambda evento: evento.dia)  

def fazerEscalaPorDia():
    # random.shuffle(Voluntários)

    metaPorFunção = {
        "Foto": 1,
        "Story": 1,
        "Mesa": 1,
        "Pré": 1,
        "Kids": 1,
        "Babys": 1,
        "Backs": 1,
        "Bateria": 1,
        "Teclado": 1,
        "Guitarra": 1,
        "Baixo": 1,
        "Violão": 1
    }

    dias_por_atributo = {
        "Foto": "diasDeFoto",
        "Story": "diasDeStory",
        "Mesa": "diasDeMesa",
        "Pré": "diasDePré",
        "Kids": "diasDeKids",
        "Babys": "diasDeBabys",
        "Backs": "diasDeBacks",
        "Bateria": "diasDeBateria",
        "Teclado": "diasDeTeclado",
        "Guitarra": "diasDeGuitarra",
        "Baixo": "diasDeBaixo",
        "Violão": "diasDeViolão"
    }

    for dia in diasJuntosOrdenados:
        ministerio = ""
        i = 0
        for atributo, valor in vars(dia).items():
            if atributo in ["dia", "tipo", "Ministro", "Auxiliares"]: continue
            if atributo in ["Foto", "Story", "Mesa"]:
                ministerio = "Mídia"
            if atributo in ["Pré", "Kids", "Babys"]:
                ministerio = "Fly"
            if atributo in ["Backs", "Bateria", "Teclado", "Guitarra", "Baixo", "Violão"]:
                ministerio = "Louvor"
            while True:
                if (
                    # Se o voluntário está no ministério
                    ministerio in Voluntários[i]["ministerios"]
                    # Se o voluntário faz a função
                    and atributo in Voluntários[i]["funcoes"]
                    # Se o voluntário está dentro do limite
                    and Voluntários[i][dias_por_atributo[atributo]] < metaPorFunção[atributo]
                    # Se o voluntário já não serve no dia
                    and dia.dia not in Voluntários[i]["servindoNosDias"]
                    # Se o voluntário serve no dia da semana
                    and Voluntários[i][dia.tipo]    
                ):
                    if atributo == "Backs":
                        while True:
                            if (
                                # Se o voluntário está no ministério
                                ministerio in Voluntários[i]["ministerios"]
                                # Se o voluntário faz a função
                                and atributo in Voluntários[i]["funcoes"]
                                # Se o voluntário está dentro do limite
                                and Voluntários[i][dias_por_atributo[atributo]] < metaPorFunção[atributo]
                                # Se o voluntário já não serve no dia
                                and dia.dia not in Voluntários[i]["servindoNosDias"]
                                # Se o voluntário serve no dia da semana
                                and Voluntários[i][dia.tipo]    
                            ):

                                novaLista = valor
                                novaLista.append(Voluntários[i]["nome"])
                                setattr(dia, atributo, novaLista)
                                Voluntários[i][dias_por_atributo[atributo]] += 2
                                Voluntários[i]["diasServidos"] += 1
                                Voluntários[i]["servindoNosDias"].append(dia.dia)

                                i = 0
                                if len(valor) == 2: break

                            elif Voluntários[i] == Voluntários[-1]:
                                i = 0
                                metaPorFunção[atributo] += 1
                            else: i += 1
                        if len(valor) == 2:
                            break
                    else: 
                        setattr(dia, atributo, Voluntários[i]["nome"])
                        Voluntários[i][dias_por_atributo[atributo]] += 1
                        Voluntários[i]["diasServidos"] += 1
                        Voluntários[i]["servindoNosDias"].append(dia.dia)
                        i = 0
                        break
                
                elif Voluntários[i] == Voluntários[-1]:
                    i = 0
                    metaPorFunção[atributo] += 1
                else:
                    i += 1

workbook = Workbook()

border_style = Side(border_style="thin", color="000000")

def fazerUmaTabela(listaDosDias, folha, row, column):
    margem = 0
    for dia in listaDosDias:
        # Título do dia
        folha.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
        titleCell = folha.cell(row=row, column=column)
        titleCell.value = f"{dia.tipo.upper()} - {dia.dia}/{mes}"
        if dia.tipo == "quinta":
            titleCell.fill = PatternFill("solid", fgColor="00339966")
            if folha.title == "Louvor":
                row += 1
                folha.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
                themeCell = folha.cell(row=row, column=column, value= "(PAUTA)")
                themeCell.fill = PatternFill("solid", fgColor="00CCFFCC")
                themeCell.alignment = Alignment(horizontal="center", vertical="center")
                themeCell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        if dia.tipo == "domingo":
            titleCell.fill = PatternFill("solid", fgColor="003366FF")
        if dia.tipo == "sabado":
            titleCell.fill = PatternFill("solid", fgColor="00003366")
        titleCell.font = Font(color="00FFFFFF")
        titleCell.alignment = Alignment(horizontal="center", vertical="center")
        titleCell.border = border_style

        # Título de voluntários 
        row += 1
        folha.cell(row=row, column=column, value=f"Função")
        blackCell = folha.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        # Título de Função 
        column += 1
        folha.cell(row=row, column=column, value=f"Voluntários")
        blackCell = folha.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        # Cargos e pessoas
        column -= 1
        margem = 0
        for atributo, valor in vars(dia).items():
            if margem < 2:
                margem += 1
                continue

            if folha.title == "Mídia" and margem not in range(2, 5):
                break

            if folha.title == "Fly" and margem not in range(5, 9):
                if margem < 6:
                    margem += 1
                    continue
                else: break

            if folha.title == "Louvor" and margem not in range(9, 16):
                if margem < 10:
                    margem += 1
                    continue
                else: break    

            row += 1
            # Fazer nome do cargo
            cell = folha.cell(row=row, column=column, value=f"{atributo}")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

            column += 1
            # Fazer pessoa do cargo
            if isinstance(valor, list):
                cell = folha.cell(row=row, column=column, value=f"{', '.join(valor)}")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
            else:    
                if atributo == "Ministro":
                    cell = folha.cell(row=row, column=column, value="")
                else:
                    cell = folha.cell(row=row, column=column, value=f"{valor}")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

            if atributo == "Violão":
                column -= 1 
                folha.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

                cell = folha.cell(row=row, column=column)
                cell.value = f"Músicas"
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.fill = PatternFill("solid", fgColor="00333333")
                cell.font = Font(color="00FFFFFF")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                for i in range(5):
                    row += 1 

                    cell = folha.cell(row=row, column=column, value="Nome")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                    column += 1

                    cell = folha.cell(row=row, column=column, value="Link")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                    column -= 1

                row += 1   
              
                cell = folha.cell(row=row, column=column, value=f"Mesa")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                cell.fill = PatternFill("solid", fgColor="00CCFFCC")

                column += 1
                cell = folha.cell(row=row, column=column, value=f"{dia.Mesa}")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                cell.fill = PatternFill("solid", fgColor="00CCFFCC")

            margem += 1
            column -= 1
        row += 2

def planilhaDaMídia():

    # fazerUmaEscala("Mídia", "story", "domingo")
    # fazerUmaEscala("Mídia", "mesa", "quinta")
    # fazerUmaEscala("Mídia", "foto", "domingo")
    # fazerUmaEscala("Mídia", "story", "quinta")
    # fazerUmaEscala("Mídia", "mesa", "domingo")
    # fazerUmaEscala("Mídia", "foto", "quinta")

    mídiaSheet = workbook.active
    mídiaSheet.title = "Mídia"
    if objetoDosDiasdeDomingo[0].dia > objetoDosDiasdeQuinta[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeQuinta, mídiaSheet, row, column)

        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeDomingo, mídiaSheet, row, column)

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeDomingo, mídiaSheet, row, column)
        
        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeQuinta, mídiaSheet, row, column)

def planilhaDoFly():
    # fazerUmaEscala("Fly", "pré", "quinta")
    # fazerUmaEscala("Fly", "kids", "quinta")
    # fazerUmaEscala("Fly", "babys", "quinta")

    # fazerUmaEscala("Fly", "pré", "domingo")
    # fazerUmaEscala("Fly", "kids", "domingo")
    # fazerUmaEscala("Fly", "babys", "domingo")

    flySheet = workbook.create_sheet("Fly")

    if objetoDosDiasdeDomingo[0].dia > objetoDosDiasdeQuinta[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeQuinta, flySheet, row, column)

        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeDomingo, flySheet, row, column)

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeDomingo, flySheet, row, column)
        
        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeQuinta, flySheet, row, column)
        
def planilhaDoLouvor():
    # fazerUmaEscala("Louvor", "ministro", "quinta")
    # fazerUmaEscala("Louvor", "bateria", "quinta")
    # fazerUmaEscala("Louvor", "teclado", "quinta")

    # fazerUmaEscala("Louvor", "ministro", "domingo")
    # fazerUmaEscala("Louvor", "bateria", "domingo")
    # fazerUmaEscala("Louvor", "teclado", "domingo")

    louvorSheet = workbook.create_sheet("Louvor")

    if objetoDosDiasdeDomingo[0].dia > objetoDosDiasdeQuinta[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeQuinta, louvorSheet, row, column)

        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeDomingo, louvorSheet, row, column)

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:

        row = 1
        column = 1
        fazerUmaTabela(objetoDosDiasdeDomingo, louvorSheet, row, column)
        
        row = 1
        column = 4
        fazerUmaTabela(objetoDosDiasdeQuinta, louvorSheet, row, column)

def planilhasDosDias():
    daysSheet = workbook.create_sheet("Dias")

    column = 1

    listaJunta = objetoDosDiasdeQuinta + objetoDosDiasdeDomingo
    listaDeDiasOrdenados = sorted(listaJunta, key=lambda evento: evento.dia)

    for dia in listaDeDiasOrdenados:
        row = 1
        # Mescla as células
        daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = daysSheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"{dia.tipo.upper()} - {dia.dia}/{mes}"

        # Define a cor de fundo
        if dia.tipo == "quinta":
            titleCell.fill = PatternFill("solid", fgColor="00339966")
        if dia.tipo == "domingo":
            titleCell.fill = PatternFill("solid", fgColor="003366FF")
        if dia.tipo == "sabado":
            titleCell.fill = PatternFill("solid", fgColor="00003366")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        daysSheet.cell(row=row, column=column, value=f"Função")
        blackCell = daysSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        daysSheet.cell(row=row, column=column, value=f"Voluntário")
        blackCell = daysSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        for atributo, valor in vars(dia).items():
            if atributo in ["dia", "tipo"]: continue

            if (atributo == "Foto"):
                column -= 1
                row += 1

                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
                celulaMinisterio = daysSheet.cell(row=row, column=column)
                celulaMinisterio.value = f"MÍDIA"
                celulaMinisterio.fill = PatternFill("solid", fgColor="00969696")
                celulaMinisterio.font = Font(color="00FFFFFF")
                celulaMinisterio.alignment = Alignment(horizontal="center", vertical="center")
                
                column += 1

            if (atributo == "Pré"):
                column -= 1
                row += 1
                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
                celulaMinisterio = daysSheet.cell(row=row, column=column)
                celulaMinisterio.value = f"FLY"
                celulaMinisterio.fill = PatternFill("solid", fgColor="00969696")
                celulaMinisterio.font = Font(color="00FFFFFF")
                celulaMinisterio.alignment = Alignment(horizontal="center", vertical="center")
                
                column += 1

            if (atributo == "Ministro"):
                column -= 1
                row += 1

                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)
                celulaMinisterio = daysSheet.cell(row=row, column=column)
                celulaMinisterio.value = f"LOUVOR"
                celulaMinisterio.fill = PatternFill("solid", fgColor="00969696")
                celulaMinisterio.font = Font(color="00FFFFFF")
                celulaMinisterio.alignment = Alignment(horizontal="center", vertical="center")
                
                column += 1

            row += 1
            column -= 1
            cell = daysSheet.cell(row=row, column=column, value=f"{atributo}")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
            column += 1

            if isinstance(valor, list):
                cell = daysSheet.cell(row=row, column=column, value=f"{', '.join(valor)}")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
            else:
                if atributo == "Ministro":
                    cell = daysSheet.cell(row=row, column=column, value="")
                else:
                    cell = daysSheet.cell(row=row, column=column, value=f"{valor}")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        column += 2
        row = 1

def relatórioDePessoas():
    relatório = workbook.create_sheet("Frequência")
    row = 1 
    column = 1
    i = 0
    while i != len(Voluntários):   
        cell = relatório.cell(row=row, column=column)
        cell.value = Voluntários[i]["nome"]
        column += 1
        cell = relatório.cell(row=row, column=column)
        cell.value = Voluntários[i]["diasServidos"]
        column -= 1
        row += 1
        i += 1

fazerEscalaPorDia()

planilhaDoFly()
planilhaDaMídia()
planilhasDosDias()
planilhaDoLouvor()
relatórioDePessoas()


workbook.save(f"Escala.xlsx")