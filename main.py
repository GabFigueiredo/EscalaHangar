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
        self.Foto = None
        self.Story = None
        self.Mesa = None
        
        # Teens
        self.professorPré = None
        self.professorKids = None
        self.professorBabys = None
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

funcao_para_atributo = {
        "foto": "Foto",
        "story": "Story",
        "mesa": "Mesa",
        "pré": "professorPré",
        "kids": "professorKids",
        "babys": "professorBabys"
    }

def fazerUmaEscala(ministerio, funcao, dia):

    if dia == "quinta":
        for diaDeQuinta in objetoDosDiasdeQuinta:
            random.shuffle(Voluntários)
            i = 0
            metaParaTodos = 1
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
            metaParaTodos = 1
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

workbook = Workbook()

border_style = Side(border_style="thin", color="000000")

def planilhaDaMídia():

    fazerUmaEscala("Mídia", "foto", "quinta")
    fazerUmaEscala("Mídia", "foto", "domingo")

    fazerUmaEscala("Mídia", "story", "quinta")
    fazerUmaEscala("Mídia", "story", "domingo")

    fazerUmaEscala("Mídia", "mesa", "quinta")
    fazerUmaEscala("Mídia", "mesa", "domingo")


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
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Foto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Story}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Mesa}")
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
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Foto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Story}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Mesa}")
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
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Foto}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Foto")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Story}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Story")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"{dia.Mesa}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = mídiaSheet.cell(row=row, column=column, value=f"Mesa")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

def planilhaDoFly():
    fazerUmaEscala("Fly", "pré", "quinta")
    fazerUmaEscala("Fly", "kids", "quinta")
    fazerUmaEscala("Fly", "babys", "quinta")

    fazerUmaEscala("Fly", "pré", "domingo")
    fazerUmaEscala("Fly", "kids", "domingo")
    fazerUmaEscala("Fly", "babys", "domingo")

    flySheet = workbook.create_sheet("Fly")
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
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorPré}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Pré")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorKids}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Kids")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorBabys}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Babys")
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
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorPré}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Pré")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorKids}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Kids")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorBabys}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Babys")
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
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorPré}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Pré")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorKids}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Kids")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        row += 1
        column -= 1
        cell = flySheet.cell(row=row, column=column, value=f"{dia.professorBabys}")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
        column += 1
        cell = flySheet.cell(row=row, column=column, value=f"Babys")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

        row += 2
        column -= 1

def planilhasDosDias():
    daysSheet = workbook.create_sheet("Dias")

    if (len(objetoDosDiasdeDomingo) > len(objetoDosDiasdeQuinta)):
        maiorLargura = len(objetoDosDiasdeDomingo)
    else:
        maiorLargura = len(objetoDosDiasdeQuinta)
    i = 0
    column = 1

    if objetoDosDiasdeQuinta[0].dia > objetoDosDiasdeDomingo[0].dia:

        for i in range(maiorLargura):
            try:
                row = 1
                # Mescla as células
                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

                # Define a célula no canto superior esquerdo da fusão
                titleCell = daysSheet.cell(row=row, column=column)

                # Define o valor da célula
                titleCell.value = f"DOMINGO - {objetoDosDiasdeDomingo[i].dia}/{mes}"

                # Define a cor de fundo
                titleCell.fill = PatternFill("solid", fgColor="003366FF")

                # Define a cor da fonte
                titleCell.font = Font(color="00FFFFFF")

                # Aplica o alinhamento
                titleCell.alignment = Alignment(horizontal="center", vertical="center")

                # Define a borda em todos os lados
                titleCell.border = border_style

                row += 1
                daysSheet.cell(row=row, column=column, value=f"Voluntários")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                column += 1
                daysSheet.cell(row=row, column=column, value=f"Função")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                for atributo, valor in vars(objetoDosDiasdeQuinta[i]).items():

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

                    if (atributo == "professorPré"):
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
                    cell = daysSheet.cell(row=row, column=column, value=f"{valor}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                    column += 1
                    cell = daysSheet.cell(row=row, column=column, value=f"{atributo}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                column += 2

                row = 1

            except(IndexError):

                print("\nChegou ao limite\n")

            try:

                # Mescla as células
                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

                # Define a célula no canto superior esquerdo da fusão
                titleCell = daysSheet.cell(row=row, column=column)

                # Define o valor da célula
                titleCell.value = f"QUINTA - {objetoDosDiasdeQuinta[i].dia}/{mes}"

                # Define a cor de fundo
                titleCell.fill = PatternFill("solid", fgColor="00339966")

                # Define a cor da fonte
                titleCell.font = Font(color="00FFFFFF")

                # Aplica o alinhamento
                titleCell.alignment = Alignment(horizontal="center", vertical="center")

                # Define a borda em todos os lados
                titleCell.border = border_style


                row += 1
                daysSheet.cell(row=row, column=column, value=f"Voluntários")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                column += 1
                daysSheet.cell(row=row, column=column, value=f"Função")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                for atributo, valor in vars(objetoDosDiasdeQuinta[i]).items():

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

                    if (atributo == "professorPré"):
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
                    cell = daysSheet.cell(row=row, column=column, value=f"{valor}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                    column += 1
                    cell = daysSheet.cell(row=row, column=column, value=f"{atributo}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                column += 2

                row = 1
            except(IndexError):
                print("\nChegou ao limite\n") 

    else:
        for i in range(maiorLargura):
            try:    
                row = 1
                # Mescla as células
                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

                # Define a célula no canto superior esquerdo da fusão
                titleCell = daysSheet.cell(row=row, column=column)

                # Define o valor da célula
                titleCell.value = f"QUINTA - {objetoDosDiasdeQuinta[i].dia}/{mes}"

                # Define a cor de fundo
                titleCell.fill = PatternFill("solid", fgColor="00339966")

                # Define a cor da fonte
                titleCell.font = Font(color="00FFFFFF")

                # Aplica o alinhamento
                titleCell.alignment = Alignment(horizontal="center", vertical="center")

                # Define a borda em todos os lados
                titleCell.border = border_style

                        
                row += 1
                daysSheet.cell(row=row, column=column, value=f"Voluntários")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                column += 1
                daysSheet.cell(row=row, column=column, value=f"Função")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                for atributo, valor in vars(objetoDosDiasdeQuinta[i]).items():

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

                    if (atributo == "professorPré"):
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
                    cell = daysSheet.cell(row=row, column=column, value=f"{valor}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                    column += 1
                    cell = daysSheet.cell(row=row, column=column, value=f"{atributo}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                column += 2

                row = 1

            except(IndexError):
                print("\nChegou ao limite\n")
        
            try:
                # Mescla as células
                daysSheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 1)

                # Define a célula no canto superior esquerdo da fusão
                titleCell = daysSheet.cell(row=row, column=column)

                # Define o valor da célula
                titleCell.value = f"DOMINGO - {objetoDosDiasdeDomingo[i].dia}/{mes}"

                # Define a cor de fundo
                titleCell.fill = PatternFill("solid", fgColor="003366FF")

                # Define a cor da fonte
                titleCell.font = Font(color="00FFFFFF")

                # Aplica o alinhamento
                titleCell.alignment = Alignment(horizontal="center", vertical="center")

                # Define a borda em todos os lados
                titleCell.border = border_style

                row += 1
                daysSheet.cell(row=row, column=column, value=f"Voluntários")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                column += 1
                daysSheet.cell(row=row, column=column, value=f"Função")
                blackCell = daysSheet.cell(row=row, column=column)
                blackCell.fill = PatternFill("solid", fgColor = "00333333")
                blackCell.font = Font(color = "00FFFFFF")
                blackCell.alignment = Alignment(horizontal="center", vertical="center")

                for atributo, valor in vars(objetoDosDiasdeQuinta[i]).items():

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

                    if (atributo == "professorPré"):
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
                    cell = daysSheet.cell(row=row, column=column, value=f"{valor}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)
                    column += 1
                    cell = daysSheet.cell(row=row, column=column, value=f"{atributo}")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = Border(left=border_style, right=border_style, top=border_style, bottom=border_style)

                column += 2

                row = 1
            except(IndexError):
                print("\nChegou ao limite\n")

planilhaDoFly()
planilhaDaMídia()
planilhasDosDias()


workbook.save(f"Escala.xlsx")