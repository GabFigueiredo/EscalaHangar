from datetime import datetime, timedelta
import calendar
import random
import inquirer
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font

import locale
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

import copy
from Voluntários import Voluntários

from Voluntários import mes, ano

Voluntários = copy.deepcopy(Voluntários)

class diaServentia:
    def __init__ (self, dia):
        self.dia = dia
        # Mídia
        self.quemTiraFoto = None
        self.quemFazStory = None
        self.quemFicaNaMesa = None
        
        # Teens
        self.professorPré = None
        self.professorKids = None
        self.professorBabys = None
        self.auxiliares = []

        # Louvor
        self.ministro = None
        self.backs = []
        self.bateria = None
        self.teclado = None
        self.guitarra = None
        self.baixo = None
        self.violão = None
    
random.shuffle(Voluntários)

diasDeQuinta = []
diasDeDomingo = []
diasDeSabado = []

# Listar todos os dias que são quintas-feiras no próximo mês
for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 3:  # 3 representa quinta-feira (segunda-feira é 0)
        diasDeQuinta.append(dia)

# Listar todos os dias que são domingo no próximo mês
for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 6:  # 3 representa quinta-feira (segunda-feira é 0)
        diasDeDomingo.append(dia)

# Listar todos os dias que são sábado no próximo mês
for dia in range(1, calendar.monthrange(ano, mes)[1] + 1):
    data = datetime(ano, mes, dia)
    if data.weekday() == 5:  # 3 representa quinta-feira (segunda-feira é 0)
        diasDeSabado.append(dia)

objetoDosDiasdeQuinta = []
objetoDosDiasdeDomingo = []
objetoDosDiasdeSabado = []

for dia in diasDeQuinta:
    temp = diaServentia(dia)
    objetoDosDiasdeQuinta.append(temp)

for dia in diasDeDomingo:
    temp = diaServentia(dia)
    objetoDosDiasdeDomingo.append(temp)

for dia in diasDeSabado:
    temp = diaServentia(dia)
    objetoDosDiasdeSabado.append(temp)

# Back-end da Mídia
def fazerEscalaDaMídia():

    ## QUINTA FEIRA ###
    # Fotos
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["foto"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemTiraFoto = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1
    # Story
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["story"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemFazStory = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                # print(f"Dia {diaDeQuinta.dia} = {Voluntários[i]["nome"]}")
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1
    # Mesa 
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["mesa"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemFicaNaMesa = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    ## DOMINGO ###
    #Fotos
    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["foto"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemTiraFoto = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["story"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemFazStory = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["mesa"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemFicaNaMesa = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    ## QUINTA FEIRA ###
    # Fotos
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["foto"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemTiraFoto = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    # Story
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["story"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemFazStory = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1
    # Mesa 
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["mesa"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.quemFicaNaMesa = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    ## DOMINGO ###
    # Fotos
    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["foto"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemTiraFoto = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["story"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemFazStory = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Mídia" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["mesa"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.quemFicaNaMesa = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

# Back-end do Fly
def fazerEscalaDoFly():

    ## QUINTA FEIRA ###
    # Pré-Teens
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):

            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["pré"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.professorPré = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    # Kids
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["kids"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.professorKids = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                # print(f"Dia {diaDeQuinta.dia} = {Voluntários[i]["nome"]}")
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1
    # Babys 
    for diaDeQuinta in objetoDosDiasdeQuinta:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["babys"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["quinta"]
            ):    
                diaDeQuinta.professorBabys = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeQuinta.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    # Pré-teens
    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["pré"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.professorPré = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    # Kids
    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["kids"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.professorKids = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

    for diaDeDomingo in objetoDosDiasdeDomingo:
        i = 0
        metaParaTodos = 1
        while i != len(Voluntários):
            if (
                "Fly" in Voluntários[i]["ministerios"]
                and Voluntários[i]["funcoes"]["babys"]
                and Voluntários[i]["diasServidos"] < metaParaTodos
                and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                and Voluntários[i]["domingo"]
            ):    
                diaDeDomingo.professorBabys = Voluntários[i]["nome"]
                Voluntários[i]["diasServidos"] += 1
                Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                break;
            if Voluntários[i] == Voluntários[-1]:
                metaParaTodos += 1 
                i = 0
            else:
                i += 1

workbook = Workbook()

border_style = Side(border_style="thin", color="000000")

def planilhaDaMídia():
    fazerUmaEscala("Mídia", "foto", "quinta")
    fazerUmaEscala("Mídia", "story", "quinta")
    fazerUmaEscala("Mídia", "mesa", "quinta")
    mídiaSheet = workbook.active
    mídiaSheet.title = "Mídia"
    if objetoDosDiasdeQuinta[0].dia < objetoDosDiasdeDomingo[0].dia:
        column = 1
    else:
        column = 4

    row = 1
    for dia in objetoDosDiasdeQuinta:
        # Mescla as células
        mídiaSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = mídiaSheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"QUINTA - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="00339966")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        mídiaSheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        mídiaSheet.cell(row=row, column=column, value=f"Função")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:
        column = 1
    else:
        column = 4

    row = 1
    for dia in objetoDosDiasdeDomingo:
        # Mescla as células
        mídiaSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = mídiaSheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"DOMINGO - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="003366FF")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        mídiaSheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        mídiaSheet.cell(row=row, column=column, value=f"Função")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

    row = 1
    column = 7
    for dia in objetoDosDiasdeSabado:
        # Mescla as células
        mídiaSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = mídiaSheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"SÁBADO - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="00003366")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        mídiaSheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        mídiaSheet.cell(row=row, column=column, value=f"Função")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

def planilhaDoFly():
    fazerEscalaDoFly()
    flySheet = workbook.active
    flySheet.title = "Fly"
    if objetoDosDiasdeQuinta[0].dia < objetoDosDiasdeDomingo[0].dia:
        column = 1
    else:
        column = 4

    row = 1
    for dia in objetoDosDiasdeQuinta:
        # Mescla as células
        flySheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = flySheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"QUINTA - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="00339966")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        flySheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        flySheet.cell(row=row, column=column, value=f"Função")
        blackCell = mídiaSheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:
        column = 1
    else:
        column = 4

    row = 1
    for dia in objetoDosDiasdeDomingo:
        # Mescla as células
        flySheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = flySheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"DOMINGO - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="003366FF")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        flySheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = flySheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        flySheet.cell(row=row, column=column, value=f"Função")
        blackCell = flySheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

    row = 1
    column = 7
    for dia in objetoDosDiasdeSabado:
        # Mescla as células
        flySheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

        # Define a célula no canto superior esquerdo da fusão
        titleCell = flySheet.cell(row=row, column=column)

        # Define o valor da célula
        titleCell.value = f"SÁBADO - {dia.dia}/{mes}"

        # Define a cor de fundo
        titleCell.fill = PatternFill("solid", fgColor="00003366")

        # Define a cor da fonte
        titleCell.font = Font(color="00FFFFFF")

        # Aplica o alinhamento
        titleCell.alignment = Alignment(horizontal="center", vertical="center")

        # Define a borda em todos os lados
        titleCell.border = border_style

        row += 1
        flySheet.cell(row=row, column=column, value=f"Voluntários")
        blackCell = flySheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        column += 1
        flySheet.cell(row=row, column=column, value=f"Função")
        blackCell = flySheet.cell(row=row, column=column)
        blackCell.fill = PatternFill("solid", fgColor = "00333333")
        blackCell.font = Font(color = "00FFFFFF")
        blackCell.alignment = Alignment(horizontal="center", vertical="center")

        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemTiraFoto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFazStory}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.quemFicaNaMesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1


def fazerUmaEscala(ministerio, funcao, dia):
    if dia == "quinta":
        for diaDeQuinta in objetoDosDiasdeQuinta:
            i = 0
            metaParaTodos = 1
            while i != len(Voluntários):
                if (
                    ministerio in Voluntários[i]["ministerios"]
                    and Voluntários[i]["funcoes"][funcao]
                    and Voluntários[i]["diasServidos"] < metaParaTodos
                    and diaDeQuinta.dia not in Voluntários[i]["servindoNosDias"]
                    and Voluntários[i]["quinta"]
                ):    
                    diaDeQuinta.quemTiraFoto = Voluntários[i]["nome"]
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
            i = 0
            metaParaTodos = 1
            while i != len(Voluntários):
                if (
                    ministerio in Voluntários[i]["ministerios"]
                    and Voluntários[i]["funcoes"][funcao]
                    and Voluntários[i]["diasServidos"] < metaParaTodos
                    and diaDeDomingo.dia not in Voluntários[i]["servindoNosDias"]
                    and Voluntários[i]["domingo"]
                ):    
                    diaDeDomingo.quemTiraFoto = Voluntários[i]["nome"]
                    Voluntários[i]["diasServidos"] += 1
                    Voluntários[i]["servindoNosDias"].append(diaDeDomingo.dia)
                    break;
                if Voluntários[i] == Voluntários[-1]:
                    metaParaTodos += 1 
                    i = 0
                else:
                    i += 1




planilhaDaMídia()

workbook.save(f"Escala.xlsx")